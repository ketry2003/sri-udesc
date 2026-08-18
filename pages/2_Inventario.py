import os
import pandas as pd
import streamlit as st

from config import QUICK_FILL_WORKBOOK_NAME
from services.data_loader import load_ttd, normalize_text
from services.equivalencias import buscar_equivalencia
from services.db import (
    insert_inventory_item,
    list_setores_inventory,
    list_inventory_items_by_setor,
    delete_inventory_items,
    delete_inventory_items_by_setor,
    update_inventory_item,
    get_next_caixa_by_setor,
)
from services.forms import (
    build_quick_fill_workbook,
    parse_inventory_workbook,
    PROVENIENCIAS_PADRAO,
)
from services.exporters import dataframe_to_excel_bytes
from services.ui_helpers import dataframe_from_rows
from services.tesauro import buscar_tesauro

st.set_page_config(page_title="Inventário Documental", layout="wide")

st.title("Inventário documental")

st.subheader("Localizar documento")

busca_livre = st.text_input(
    "Digite o nome do documento, processo ou assunto"
)

tipo_busca = st.session_state.get(
    "tipo",
    "fim"
)

resultado_busca = pd.DataFrame()

if busca_livre:

    resultado_busca = buscar_tesauro(
        busca_livre,
        tipo_busca,
        limite=20
    )

    st.write(
        f"Resultados encontrados: {len(resultado_busca)}"
    )

    if not resultado_busca.empty:

        opcoes = []

        for _, row in resultado_busca.iterrows():

            documento = row.get(
                "termo_preferido_oficial",
                ""
            )

            codigo = row.get(
                "codigo_classificacao",
                ""
            )

            opcoes.append(
                f"{codigo} | {documento}"
            )

        documento_selecionado_busca = st.selectbox(
            "Selecione um documento",
            [""] + opcoes,
            key="busca_inventario"
        )

    else:

        st.warning(
            "Nenhum documento encontrado."
        )

# =============================
# SINCRONIZAÇÃO DE ESTADO
# =============================

def obter_documento_selecionado():
    """Re-sincroniza o documento selecionado do app.py a cada render."""
    return st.session_state.get("documento_selecionado", {})


documento_temporalidade = obter_documento_selecionado()

natureza_pre_preenchida = (
    documento_temporalidade.get("natureza", "")
    if documento_temporalidade
    else ""
)

if documento_temporalidade and documento_temporalidade.get("documento"):
    st.info(
        f"✓ Documento selecionado da Consulta de Temporalidade: "
        f"**{documento_temporalidade.get('documento')}**"
    )
    st.caption(
        "Este documento será usado para pré-preencher a busca assistida abaixo. "
        "Você pode usar este formulário ou fazer uma nova busca."
    )
    st.divider()



@st.cache_data(show_spinner=False)
def carregar_ttd():
    return load_ttd("todos")


df_ttd = carregar_ttd()


def calcular_ano_eliminacao(
    ano_emissao,
    prazo_corrente,
    prazo_intermediario,
    destinacao_final,
):
    if not ano_emissao:
        return "-"

    destino = str(destinacao_final or "").lower()

    if "permanente" in destino:
        return "-"

    try:
        ano = int(str(ano_emissao).strip())
        return str(
            ano
            + int(prazo_corrente or 0)
            + int(prazo_intermediario or 0)
        )
    except Exception:
        return "-"


def calcular_total(prazo_corrente, prazo_intermediario):
    try:
        return int(prazo_corrente or 0) + int(prazo_intermediario or 0)
    except Exception:
        return ""


def filtrar_ttd(df, codigo_busca="", tipo_busca="", natureza=""):
    out = df.copy()

    if natureza and "natureza_documental" in out.columns:
        out = out[
            out["natureza_documental"]
            .astype(str)
            .str.strip()
            .str.lower()
            == natureza.strip().lower()
        ]

    if codigo_busca.strip() and "codigo_classificacao" in out.columns:
        codigo_norm = normalize_text(codigo_busca)
        out = out[
            out["codigo_classificacao"]
            .astype(str)
            .map(normalize_text)
            .str.contains(codigo_norm, na=False)
        ]

    if tipo_busca.strip() and "item_documental" in out.columns:
        tipo_norm = normalize_text(tipo_busca)
        out = out[
            out["item_documental"]
            .astype(str)
            .map(normalize_text)
            .str.contains(tipo_norm, na=False)
        ]

    colunas_ordenacao = [
        c
        for c in ["source_priority", "codigo_classificacao", "item_documental"]
        if c in out.columns
    ]

    if colunas_ordenacao:
        out = out.sort_values(colunas_ordenacao)

    return out.head(30)


def valor_registro(registro, campo, padrao=""):
    if registro is None:
        return padrao

    try:
        valor = registro.get(campo, padrao)
    except AttributeError:
        valor = registro[campo] if campo in registro else padrao

    return valor if valor is not None else padrao


aba1, aba2 = st.tabs(["Adicionar item", "Inventário salvo"])


with aba1:
    st.caption(
        "Baixe o formulário oficial do inventário em Excel ou faça o "
        "preenchimento assistido abaixo. A busca pode ser feita "
        "primeiro pelo código de classificação e depois pelo tipo "
        "documental."
    )

    if documento_temporalidade and documento_temporalidade.get("documento"):
        col_info, col_clear = st.columns([4, 1])
        with col_info:
            st.info(
                f"📄 **Documento carregado:** {documento_temporalidade.get('documento')}"
            )
        with col_clear:
            if st.button("Limpar documento", key="btn_limpar_documento"):
                st.session_state["documento_selecionado"] = {}
                st.rerun()

    st.divider()
    st.subheader("Preenchimento assistido")

    c1, c2, c3 = st.columns(3)

    opcoes_natureza = [
        "",
        "Atividade-fim",
        "Atividade-meio"
    ]

    indice_natureza = (
        opcoes_natureza.index(natureza_pre_preenchida)
        if natureza_pre_preenchida in opcoes_natureza
        else 0
    )

    natureza_escolhida = c1.selectbox(
        "Natureza documental",
        opcoes_natureza,
        index=indice_natureza,
    )

    codigo_busca = c2.text_input(
        "Buscar pelo código de classificação",
        placeholder="Ex.: 06.24.01",
    )

    tipo_busca = c3.text_input(
        "Digite o nome do documento, processo ou assunto",
        value=documento_temporalidade.get("documento", "") or "",
        placeholder=(
            "Ex.: contrato administrativo, licitação, "
            "férias, almoxarifado, edital, processo"
        ),
        help=(
            "Se um documento foi selecionado na página anterior, "
            "ele aparecerá aqui. Modifique ou deixe como está."
        ),
    )

    st.divider()

workbook_bytes = build_quick_fill_workbook(df_ttd)

# ==================================
# EQUIVALÊNCIAS HISTÓRICAS
# ==================================

if tipo_busca:
    termo_equivalente = buscar_equivalencia(tipo_busca)

    if termo_equivalente:
        st.info(
            f"Equivalência histórica encontrada: "
            f"{termo_equivalente}"
        )
        tipo_busca = termo_equivalente


# ==================================
# BUSCA NORMAL
# ==================================

sugestoes = pd.DataFrame()

registro = None 

if tipo_busca.strip():

    tipo_atividade = (
        "meio"
        if natureza_escolhida == "Atividade-meio"
         else "fim"
    )

    sugestoes = buscar_tesauro(
        tipo_busca,
        tipo_atividade,
        limite=20
    )

    if (
        documento_temporalidade
        and not sugestoes.empty
    ):
        registro = sugestoes.iloc[0]

if not sugestoes.empty:

    opcoes = [
        f"{row.get('codigo_classificacao', '')} | "
        f"{row.get('item_documental', '')} | "
        f"{row.get('natureza_documental', '')}"
        for _, row in sugestoes.iterrows()
]

else:

    sugestoes = filtrar_ttd(
        df_ttd,
        codigo_busca,
        "",
        natureza_escolhida,
    )

if not sugestoes.empty:

    opcoes = [
        f"{row.get('codigo_classificacao', '')} | "
        f"{row.get('item_documental', '')} | "
        f"{row.get('natureza_documental', '')}"
        for _, row in sugestoes.iterrows()
    ]

    selecionado = st.selectbox(
        "Resultado da TTD",
        [""] + opcoes
    )

    if selecionado:
        registro = sugestoes.iloc[
            opcoes.index(selecionado)
        ]
        
    if registro is not None:
        st.success("Classificação localizada.")

        a, b, c = st.columns(3)

        a.write(
            f"**Código:** "
            f"{valor_registro(registro, 'codigo_classificacao', '-') or '-'}"
        )

        b.write(
            f"**Tipo documental:** "
            f"{valor_registro(registro, 'item_documental', '-') or '-'}"
        )

        c.write(
            f"**Assunto:** "
            f"{valor_registro(registro, 'assunto', '-') or '-'}"
        )

        d, e = st.columns(2)

        d.write(
            f"**Subsérie:** "
            f"{valor_registro(registro, 'subserie', '-') or '-'}"
        )

        e.write(
            f"**Dossiê/Processo:** "
            f"{valor_registro(registro, 'dossie_processo', '-') or '-'}"
        )

    with st.form("form_inventario_assistido"):
        st.markdown("### Dados do inventário")

        permitir_sem_classificacao = st.checkbox(
            "Documento sem classificação definida",
            help=(
                "Use quando o documento não se encaixar claramente em nenhuma "
                "categoria da TTD. O item ficará marcado como A avaliar."
            ),
        )

        col1, col2, col3 = st.columns(3)

        opcoes_proveniencia = (
            ["Selecione..."]
            + PROVENIENCIAS_PADRAO
            + ["Outro / informar manualmente"]
        )

        proveniencia_opcao = col1.selectbox(
            "Proveniência / Setor *",
            opcoes_proveniencia,
            index=0,
            help=(
                "Campo obrigatório. Selecione o setor produtor da "
                "documentação."
            ),
        )

        proveniencia_outro = ""

        if proveniencia_opcao == "Outro / informar manualmente":
            proveniencia_outro = st.text_input(
                "Informe o nome do setor/proveniência",
                placeholder="Ex.: Setor de Obras",
            )

        proveniencia = (
            proveniencia_outro.strip()
            if proveniencia_opcao == "Outro / informar manualmente"
            else proveniencia_opcao
        )

        ano_emissao = col2.text_input(
            "Ano de emissão",
            placeholder="Ex.: 2024",
        )

        referencia = col3.text_input(
            "Referência",
            placeholder="Ex.: nº processo, turma, matrícula",
        )

        col4, col5, col6 = st.columns(3)

        assunto_digitado = col4.text_input(
            "Assunto",
            value=(
                valor_registro(registro, "assunto", "")
                if registro is not None
                else ""
            ),
            placeholder="Descreva resumidamente o assunto",
        )

        caixa_sugerida = ""

        if proveniencia and proveniencia != "Selecione...":
            try:
                caixa_sugerida = get_next_caixa_by_setor(proveniencia)
            except Exception:
                caixa_sugerida = ""

        try:
            valor_inicial = int(str(caixa_sugerida).strip())
        except Exception:
            valor_inicial = 1

        numero_caixa = col5.number_input(
            "Nº Caixa",
            min_value=1,
            step=1,
            value=valor_inicial,
            help="Número da caixa.",
        )

        caixa_formatada = f"{numero_caixa:03d}"

        quantidade = col6.number_input(
            "Quantidade",
            min_value=1,
            value=1,
        )

        st.markdown("### Campos preenchidos automaticamente")

        codigo_classificacao = valor_registro(
            registro,
            "codigo_classificacao",
            "",
        )

        tipo_documental = valor_registro(
            registro,
            "item_documental",
            "",
        )

        classe_ttd = valor_registro(
            registro,
            "item_documental",
            "",
        )

        prazo_corrente = valor_registro(
            registro,
            "prazo_corrente",
            0,
        )

        prazo_intermediario = valor_registro(
            registro,
            "prazo_intermediario",
            0,
        )

        destinacao_final = valor_registro(
            registro,
            "destinacao_final",
            "",
        )

        subserie = valor_registro(
            registro,
            "subserie",
            "",
        )

        dossie_processo = valor_registro(
            registro,
            "dossie_processo",
            "",
        )

        total_prazo = calcular_total(
            prazo_corrente,
            prazo_intermediario,
        )

        ano_eliminacao = calcular_ano_eliminacao(
            ano_emissao,
            prazo_corrente,
            prazo_intermediario,
            destinacao_final,
        )

        guarda_permanente = (
            "Guarda permanente"
            if "permanente" in str(destinacao_final).lower()
            else "-"
        )

        x1, x2, x3 = st.columns(3)

        x1.text_input(
            "Código segundo o Plano de Classificação",
            value=str(codigo_classificacao or ""),
            disabled=True,
        )

        x2.text_input(
            "Tipo documental localizado",
            value=str(tipo_documental or ""),
            disabled=True,
        )

        x3.text_input(
            "Classe segundo a TTD",
            value=str(classe_ttd or ""),
            disabled=True,
        )

        x4, x5 = st.columns(2)

        x4.text_input(
            "Subsérie",
            value=str(subserie or ""),
            disabled=True,
        )

        x5.text_input(
            "Dossiê/Processo",
            value=str(dossie_processo or ""),
            disabled=True,
        )

        y1, y2, y3 = st.columns(3)

        y1.text_input(
            "Fase Corrente em anos",
            value=str(prazo_corrente or ""),
            disabled=True,
        )

        y2.text_input(
            "Fase Intermediária em anos",
            value=str(prazo_intermediario or ""),
            disabled=True,
        )

        y3.text_input(
            "Total",
            value=str(total_prazo),
            disabled=True,
        )

        z1, z2, z3 = st.columns(3)

        z1.text_input(
            "Destinação",
            value=str(destinacao_final or ""),
            disabled=True,
        )

        z2.text_input(
            "Eliminação no ano de",
            value=str(ano_eliminacao),
            disabled=True,
        )

        z3.text_input(
            "Guarda Permanente",
            value=str(guarda_permanente),
            disabled=True,
        )

        observacoes = st.text_area(
            "Observações",
            placeholder=(
                "Use este campo para complementar a proveniência, detalhar "
                "referência ou informar particularidades."
            ),
        )

        enviado = st.form_submit_button("Adicionar ao inventário")

        if enviado:
            if proveniencia == "Selecione..." or not proveniencia.strip():
                st.error("Selecione ou informe a Proveniência / Setor.")

            elif registro is None and not permitir_sem_classificacao:
                st.error(
                    "Primeiro localize um documento na TTD ou marque a opção "
                    "Documento sem classificação definida."
                )

            else:
                if registro is None and permitir_sem_classificacao:
                    registro = {
                        "natureza_documental": "Não classificado",
                        "grupo": "",
                        "subgrupo": "",
                        "serie": "",
                        "subserie": "",
                        "dossie_processo": "",
                        "item_documental": (
                            assunto_digitado
                            or "Documento sem classificação definida"
                        ),
                        "codigo_classificacao": "",
                        "prazo_corrente": "",
                        "prazo_intermediario": "",
                        "destinacao_final": "A avaliar",
                    }

                    codigo_classificacao = ""
                    tipo_documental = (
                        assunto_digitado
                        or "Documento sem classificação definida"
                    )
                    classe_ttd = tipo_documental
                    prazo_corrente = ""
                    prazo_intermediario = ""
                    destinacao_final = "A avaliar"

                texto_obs = []

                if referencia:
                    texto_obs.append(f"Referência: {referencia}")

                if assunto_digitado:
                    texto_obs.append(f"Assunto: {assunto_digitado}")

                if observacoes:
                    texto_obs.append(observacoes)

                payload = {
                    "setor": proveniencia,
                    "tipo_documental": tipo_documental,
                    "natureza_documental": valor_registro(
                        registro,
                        "natureza_documental",
                        "",
                    ),
                    "grupo": valor_registro(registro, "grupo", ""),
                    "subgrupo": valor_registro(registro, "subgrupo", ""),
                    "serie": valor_registro(registro, "serie", ""),
                    "subserie": valor_registro(registro, "subserie", ""),
                    "dossie_processo": valor_registro(
                        registro,
                        "dossie_processo",
                        "",
                    ),
                    "item_documental": classe_ttd,
                    "codigo_classificacao": codigo_classificacao,
                    "prazo_corrente": prazo_corrente,
                    "prazo_intermediario": prazo_intermediario,
                    "destinacao_final": destinacao_final,
                    "datas_limite": str(ano_emissao or ""),
                    "quantidade": quantidade,
                    "caixa": caixa_formatada,
                    "observacoes": " | ".join([t for t in texto_obs if t]),
                }

                try:
                    insert_inventory_item(payload)
                    st.success(
                        "✅ Item adicionado ao inventário com preenchimento assistido."
                    )
                except Exception as e:
                    st.error(
                        f"❌ Erro ao adicionar item ao inventário: {str(e)}\n\n"
                        f"Verifique se o banco de dados está acessível e tente novamente."
                    )

    st.divider()
    st.subheader("Importar planilha preenchida")

    st.info(
        "A importação abaixo substitui apenas os itens do setor escolhido, "
        "sem afetar os demais setores."
    )

    opcoes_importacao = (
        [""]
        + PROVENIENCIAS_PADRAO
        + ["Outro / informar manualmente"]
    )

    setor_importacao_opcao = st.selectbox(
        "Setor/proveniência para importação",
        opcoes_importacao,
        index=0,
        key="setor_importacao_opcao",
    )

    setor_importacao_outro = ""

    if setor_importacao_opcao == "Outro / informar manualmente":
        setor_importacao_outro = st.text_input(
            "Digite o nome do setor/proveniência para "
            "importação",
            placeholder="Ex.: Setor de Obras",
            key="setor_importacao_outro",
        )

    setor_importacao = (
        setor_importacao_outro.strip()
        if setor_importacao_opcao == "Outro / informar manualmente"
        else setor_importacao_opcao
    ) if setor_importacao_opcao else ""

    arquivo = st.file_uploader(
        "Selecione o arquivo Excel do inventário",
        type=["xlsx"],
    )

    if arquivo is not None:
        try:
            df_importado = parse_inventory_workbook(arquivo, df_ttd)

            count = len(df_importado)
            st.write(f"{count} item(ns) localizado(s) na planilha.")
            st.dataframe(df_importado, use_container_width=True)

            if not df_importado.empty:
                if not setor_importacao.strip():
                    st.warning(
                        "Selecione ou informe o setor/proveniência "
                        "antes de importar a planilha."
                    )
                else:
                    st.warning(
                        f"Confirme a importação para substituir "
                        f"apenas o inventário do setor: "
                        f"{setor_importacao}."
                    )

            btn_importar = st.button(
                "Importar planilha para este setor"
            )
            if btn_importar:
                if not setor_importacao.strip():
                    st.error(
                        "Selecione ou informe o setor/proveniência "
                        "antes de importar."
                    )

                elif df_importado.empty:
                    st.warning("A planilha não possui itens para importar.")

                else:
                    delete_inventory_items_by_setor(setor_importacao)

                    total = 0
                    erros = []

                    for idx, row in df_importado.iterrows():
                        try:
                            payload = {
                                "setor": setor_importacao,
                                "tipo_documental": row.get("tipo_documental", ""),
                                "natureza_documental": row.get(
                                    "natureza_documental",
                                    "",
                                ),
                                "grupo": row.get("grupo", ""),
                                "subgrupo": row.get("subgrupo", ""),
                                "serie": row.get("serie", ""),
                                "subserie": row.get("subserie", ""),
                                "dossie_processo": row.get(
                                    "dossie_processo",
                                    "",
                                ),
                                "item_documental": row.get("item_documental", ""),
                                "codigo_classificacao": row.get(
                                    "codigo_classificacao",
                                    "",
                                ),
                                "prazo_corrente": row.get("prazo_corrente", ""),
                                "prazo_intermediario": row.get(
                                    "prazo_intermediario",
                                    "",
                                ),
                                "destinacao_final": row.get(
                                    "destinacao_final",
                                    "",
                                ),
                                "datas_limite": row.get("datas_limite", ""),
                                "quantidade": int(row.get("quantidade", 1) or 1),
                                "caixa": row.get("caixa", ""),
                                "observacoes": row.get("observacoes", ""),
                            }

                            insert_inventory_item(payload)
                            total += 1
                        except Exception as e:
                            erros.append(f"Linha {idx + 1}: {str(e)}")

                    if total > 0:
                        st.success(
                            f"✅ Inventário do setor {setor_importacao} atualizado com {total} item(ns)."
                        )

                    if erros:
                        with st.expander(f"⚠️ {len(erros)} linha(s) com erro"):
                            for erro in erros:
                                st.caption(f"❌ {erro}")

                    if total > 0:
                        st.rerun()

        except Exception as e:
            st.error(f"Erro ao ler planilha: {e}")

# FINAL DA ABA 1

st.divider()

st.subheader("Download do formulário")

st.download_button(
        "Baixar formulário oficial em Excel (preenchimento rápido)",
        data=workbook_bytes,
        file_name=QUICK_FILL_WORKBOOK_NAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help=(
            "Modelo oficial do inventário documental com preenchimento assistido "
            "e classificação automática."
        ),
    )

with st.expander("Ajuda de busca avançada", expanded=False):
                st.markdown(
                    """
                    **Como pesquisar**
                    - Use o **código de classificação** quando souber o código exato
                    ou parte dele.
                    - Use o **tipo documental** para procurar pelo nome do documento.
                    - Você pode combinar os dois campos para refinar o resultado.
                    - A busca aceita trechos do texto, não precisa digitar o nome
                    completo.
                    - A natureza documental pode ser usada como filtro adicional.
                    """
                )

with aba2:
    st.subheader("Inventário salvo por setor / proveniência")

    setores = list_setores_inventory()

    if not setores:
        st.info("Nenhum item salvo ainda.")

    else:
        setor_escolhido = st.selectbox(
            "Selecione o setor",
            [""] + setores,
            index=0,
        )

        if setor_escolhido:
            rows = list_inventory_items_by_setor(setor_escolhido)
            df_inv = dataframe_from_rows(rows)

            if df_inv.empty:
                st.info("Nenhum item salvo para este setor.")

            else:
                st.dataframe(df_inv, use_container_width=True)

                export_cols = [
                    "setor",
                    "tipo_documental",
                    "subserie",
                    "dossie_processo",
                    "item_documental",
                    "codigo_classificacao",
                    "prazo_corrente",
                    "prazo_intermediario",
                    "destinacao_final",
                    "datas_limite",
                    "quantidade",
                    "caixa",
                    "observacoes",
                    "criado_em",
                ]

                export_df = df_inv[
                    [c for c in export_cols if c in df_inv.columns]
                ].copy()

                st.download_button(
                    "Baixar inventário deste setor em Excel",
                    data=dataframe_to_excel_bytes(export_df, "Inventário"),
                    file_name=f"inventario_{setor_escolhido}.xlsx".replace(
                        " ",
                        "_",
                    ),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            st.markdown("### Editar ou excluir item salvo")

            opcoes_edicao = []
            mapa_edicao = {}

            for _, row in df_inv.iterrows():
                rotulo = (
                    f"#{row['id']} | "
                    f"{row.get('tipo_documental', '-') or '-'} | "
                    f"Caixa {row.get('caixa', '-') or '-'} | "
                    f"{row.get('datas_limite', '-') or '-'}"
                )

                opcoes_edicao.append(rotulo)
                mapa_edicao[rotulo] = row

            item_para_editar = st.selectbox(
                    "Selecione o item",
                    [""] + opcoes_edicao,
                )

            @st.dialog("Confirmar exclusão")
            def confirmar_exclusao_item(item_id, descricao_item):
                st.error(
                    "⚠️ ATENÇÃO: Esta ação é irreversível. O item será excluído "
                    "permanentemente do banco de dados."
                )

                st.write("Confira o item a ser excluído:")
                st.write(f"**{descricao_item}**")

                st.divider()

                # Inicializar contador de tentativas na session
                if f"tentativas_exclusao_{item_id}" not in st.session_state:
                    st.session_state[f"tentativas_exclusao_{item_id}"] = 0

                max_tentativas = 3
                tentativas_restantes = max_tentativas - st.session_state[f"tentativas_exclusao_{item_id}"]

                if tentativas_restantes <= 0:
                    st.error("🔒 Limite de tentativas excedido. Feche este diálogo e tente novamente.")
                    st.stop()

                if tentativas_restantes <= 1:
                    st.warning(f"⚠️ {tentativas_restantes} tentativa restante!")

                senha_confirmacao = st.text_input(
                    "Informe a senha de administrador para confirmar",
                    type="password",
                    key=f"senha_confirmacao_exclusao_{item_id}",
                )

                st.caption(
                    "A senha é armazenada na variável de ambiente "
                    "INVENTARIO_DELETE_PASSWORD (configuração de segurança)."
                )

                col_confirmar, col_cancelar = st.columns(2)

                with col_confirmar:
                    if st.button(
                        "Confirmar exclusão permanente",
                        type="primary",
                        key=f"confirmar_exclusao_{item_id}",
                    ):
                        senha_esperada = os.getenv("INVENTARIO_DELETE_PASSWORD", "")

                        if not senha_esperada:
                            st.error(
                                "🔧 Erro de configuração: senha de administrador não definida. "
                                "Entre em contato com o administrador do sistema."
                            )
                        elif not senha_confirmacao:
                            st.error("Digite a senha para confirmar.")
                        elif senha_confirmacao != senha_esperada:
                            st.session_state[f"tentativas_exclusao_{item_id}"] += 1
                            tentativas_restantes = max_tentativas - st.session_state[f"tentativas_exclusao_{item_id}"]
                            st.error(
                                f"❌ Senha incorreta. "
                                f"Tentativas restantes: {tentativas_restantes}. "
                                "Exclusão cancelada."
                            )
                        else:
                            try:
                                total = delete_inventory_items([int(item_id)]) or 0

                                if total > 0:
                                    st.success("✅ Item excluído com sucesso do banco de dados.")
                                    st.session_state[f"tentativas_exclusao_{item_id}"] = 0
                                    st.rerun()
                                else:
                                    st.warning("⚠️ Nenhum item foi excluído (verifique se o ID ainda existe).")
                            except Exception as e:
                                st.error(f"❌ Erro ao excluir item: {str(e)}")

                with col_cancelar:
                    if st.button(
                        "Cancelar",
                        key=f"cancelar_exclusao_{item_id}",
                    ):
                        st.session_state[f"tentativas_exclusao_{item_id}"] = 0
                        st.rerun()

            total = None

            if item_para_editar:
                item = mapa_edicao[item_para_editar]

                st.markdown("#### Editar item")

                with st.form("form_editar_item"):
                    novo_ano = st.text_input(
                        "Ano de emissão / Datas-limite",
                        value=str(item.get("datas_limite", "") or ""),
                    )

                    try:
                        caixa_atual = int(item.get("caixa", 1) or 1)
                    except Exception:
                        caixa_atual = 1

                    nova_caixa_numero = st.number_input(
                        "Nº Caixa",
                        min_value=1,
                        value=caixa_atual,
                    )

                    nova_caixa = f"{nova_caixa_numero:03d}"

                    try:
                        quantidade_atual = int(
                            item.get("quantidade", 1) or 1
                        )
                    except Exception:
                        quantidade_atual = 1

                    nova_quantidade = st.number_input(
                        "Quantidade",
                        min_value=1,
                        value=quantidade_atual,
                    )

                    novas_observacoes = st.text_area(
                        "Observações",
                        value=str(item.get("observacoes", "") or ""),
                    )

                    salvar_edicao = st.form_submit_button(
                        "Salvar alterações"
                    )

                    if salvar_edicao:
                        payload = {
                            "datas_limite": novo_ano,
                            "quantidade": nova_quantidade,
                            "caixa": nova_caixa,
                            "observacoes": novas_observacoes,
                        }

                        total = update_inventory_item(
                            int(item["id"]),
                            payload,
                        )

                        if total is not None and total > 0:
                            st.success("Item atualizado com sucesso.")
                            st.rerun()
                        else:
                            st.warning("Nenhuma alteração foi salva.")

                    st.markdown("#### Excluir item")

                    descricao_item = (
                        f"#{item['id']} | "
                        f"{item.get('tipo_documental', '-') or '-'} | "
                        f"Caixa {item.get('caixa', '-') or '-'} | "
                        f"{item.get('datas_limite', '-') or '-'}"
                    )

                    st.warning(
                        "A exclusão agora é feita somente item por item."
                    )

                    if st.button(
                        "Excluir este item",
                        type="secondary",
                        key=f"abrir_confirmacao_exclusao_{item['id']}",
                    ):
                        confirmar_exclusao_item(
                            int(item["id"]),
                            descricao_item,
                        )
