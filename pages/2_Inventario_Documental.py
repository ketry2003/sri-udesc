import pandas as pd
import streamlit as st

from config import QUICK_FILL_WORKBOOK_NAME
from services.data_loader import load_ttd, normalize_text
from services.db import (
    insert_inventory_item,
    list_setores_inventory,
    list_inventory_items_by_setor,
    delete_inventory_items,
    delete_inventory_items_by_setor,
    replace_inventory_from_dataframe_by_setor,
)
from services.forms import (
    build_quick_fill_workbook,
    parse_inventory_workbook,
    PROVENIENCIAS_PADRAO,
)
from services.exporters import dataframe_to_excel_bytes
from services.ui_helpers import dataframe_from_rows


st.set_page_config(page_title="Inventário Documental", layout="wide")

st.title("Inventário documental")

# Base consolidada: atividade-fim + atividade-meio
df_ttd = load_ttd("todos")


def calcular_ano_eliminacao(
    ano_emissao,
    prazo_corrente,
    prazo_intermediario,
    destinacao_final
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

    if natureza:
        out = out[out["natureza_documental"] == natureza]

    if codigo_busca.strip():
        codigo_norm = normalize_text(codigo_busca)
        out = out[
            out["codigo_classificacao"]
            .astype(str)
            .map(normalize_text)
            .str.contains(codigo_norm, na=False)
        ]

    if tipo_busca.strip():
        tipo_norm = normalize_text(tipo_busca)
        out = out[
            out["item_documental"]
            .astype(str)
            .map(normalize_text)
            .str.contains(tipo_norm, na=False)
        ]

    colunas_ordenacao = [
        c for c in ["source_priority", "codigo_classificacao", "item_documental"]
        if c in out.columns
    ]

    if colunas_ordenacao:
        out = out.sort_values(colunas_ordenacao)

    return out.head(30)


aba1, aba2 = st.tabs(["Adicionar item", "Inventário salvo"])

with aba1:
    st.caption(
        "Baixe o formulário oficial do inventário em Excel ou faça o preenchimento "
        "assistido abaixo. A busca pode ser feita primeiro pelo código de "
        "classificação e depois pelo tipo documental."
    )

    with st.expander("Ajuda de busca avançada", expanded=False):
        st.markdown(
            """
            **Como pesquisar**
            - Use o **código de classificação** quando souber o código exato ou parte dele.
            - Use o **tipo documental** para procurar pelo nome do documento.
            - Você pode combinar os dois campos para refinar o resultado.
            - A busca aceita trechos do texto, não precisa digitar o nome completo.
            - A natureza documental pode ser usada como filtro adicional.
            """
        )

    workbook_bytes = build_quick_fill_workbook(df_ttd)
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

    st.divider()
    st.subheader("Preenchimento assistido")

    c1, c2, c3 = st.columns(3)
    natureza_escolhida = c1.selectbox(
        "Natureza documental",
        ["", "Atividade-fim", "Atividade-meio"],
        index=0,
        help=(
            "Use atividade-fim como principal. "
            "Use atividade-meio como apoio ou segunda fonte."
        ),
    )
    codigo_busca = c2.text_input(
        "Buscar pelo código de classificação",
        placeholder="Ex.: 06.24.01",
    )
    tipo_busca = c3.text_input(
        "Buscar pelo tipo documental",
        placeholder="Ex.: diário de classe, ofício, ata",
    )

    sugestoes = filtrar_ttd(df_ttd, codigo_busca, tipo_busca, natureza_escolhida)

    registro = None
    if not sugestoes.empty:
        opcoes = [
            f"{row.get('codigo_classificacao', '')} | "
            f"{row.get('item_documental', '')} | "
            f"{row.get('natureza_documental', '')}"
            for _, row in sugestoes.iterrows()
        ]

        selecionado = st.selectbox("Resultado da TTD", [""] + opcoes)

        if selecionado:
            registro = sugestoes.iloc[opcoes.index(selecionado)]

    if registro is not None:
        st.success("Classificação localizada.")

        a, b, c = st.columns(3)
        a.write(f"**Código:** {registro.get('codigo_classificacao', '') or '-'}")
        b.write(f"**Tipo documental:** {registro.get('item_documental', '') or '-'}")
        c.write(f"**Assunto:** {registro.get('assunto', '') or '-'}")

        d, e = st.columns(2)
        d.write(f"**Subsérie:** {registro.get('subserie', '') or '-'}")
        e.write(f"**Dossiê/Processo:** {registro.get('dossie_processo', '') or '-'}")

    with st.form("form_inventario_assistido"):
        st.markdown("### Dados do inventário")

        col1, col2, col3 = st.columns(3)
        proveniencia = col1.selectbox(
            "Proveniência",
            [""] + PROVENIENCIAS_PADRAO,
            index=0,
        )
        ano_emissao = col2.text_input("Ano de emissão", placeholder="Ex.: 2024")
        referencia = col3.text_input(
            "Referência",
            placeholder="Ex.: nº processo, turma, matrícula",
        )

        col4, col5, col6 = st.columns(3)
        assunto_digitado = col4.text_input(
            "Assunto",
            value=registro["assunto"] if registro is not None and registro.get("assunto") else "",
            placeholder="Descreva resumidamente o assunto",
        )
        numero_caixa = col5.text_input("Nº Caixa")
        quantidade = col6.number_input("Quantidade", min_value=1, value=1)

        st.markdown("### Campos preenchidos automaticamente")

        codigo_classificacao = registro["codigo_classificacao"] if registro is not None else ""
        tipo_documental = registro["item_documental"] if registro is not None else ""
        classe_ttd = registro["item_documental"] if registro is not None else ""
        prazo_corrente = registro["prazo_corrente"] if registro is not None and "prazo_corrente" in registro else 0
        prazo_intermediario = registro["prazo_intermediario"] if registro is not None and "prazo_intermediario" in registro else 0
        destinacao_final = registro["destinacao_final"] if registro is not None else ""
        subserie = registro["subserie"] if registro is not None and registro.get("subserie") else ""
        dossie_processo = registro["dossie_processo"] if registro is not None and registro.get("dossie_processo") else ""

        total_prazo = calcular_total(prazo_corrente, prazo_intermediario)
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
            value=codigo_classificacao,
            disabled=True,
        )
        x2.text_input(
            "Tipo documental localizado",
            value=tipo_documental,
            disabled=True,
        )
        x3.text_input(
            "Classe segundo a TTD",
            value=classe_ttd,
            disabled=True,
        )

        x4, x5 = st.columns(2)
        x4.text_input(
            "Subsérie",
            value=subserie,
            disabled=True,
        )
        x5.text_input(
            "Dossiê/Processo",
            value=dossie_processo,
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
        y3.text_input("Total", value=str(total_prazo), disabled=True)

        z1, z2, z3 = st.columns(3)
        z1.text_input("Destinação", value=destinacao_final, disabled=True)
        z2.text_input("Eliminação no ano de", value=ano_eliminacao, disabled=True)
        z3.text_input("Guarda Permanente", value=guarda_permanente, disabled=True)

        observacoes = st.text_area(
            "Observações",
            placeholder=(
                "Use este campo para complementar a proveniência, detalhar "
                "referência ou informar particularidades."
            ),
        )

        enviado = st.form_submit_button("Adicionar ao inventário")

        if enviado:
            if registro is None:
                st.error(
                    "Primeiro localize um documento na TTD pelo código de "
                    "classificação ou pelo tipo documental."
                )
            elif not proveniencia.strip():
                st.error("Informe a proveniência antes de adicionar o item.")
            else:
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
                    "natureza_documental": registro.get("natureza_documental", ""),
                    "grupo": registro.get("grupo", ""),
                    "subgrupo": registro.get("subgrupo", ""),
                    "serie": registro.get("serie", ""),
                    "subserie": registro.get("subserie", ""),
                    "dossie_processo": registro.get("dossie_processo", ""),
                    "item_documental": classe_ttd,
                    "codigo_classificacao": codigo_classificacao,
                    "prazo_corrente": prazo_corrente,
                    "prazo_intermediario": prazo_intermediario,
                    "destinacao_final": destinacao_final,
                    "datas_limite": str(ano_emissao or ""),
                    "quantidade": quantidade,
                    "caixa": numero_caixa,
                    "observacoes": " | ".join([t for t in texto_obs if t]),
                }

                insert_inventory_item(payload)
                st.success("Item adicionado ao inventário com preenchimento assistido.")

    st.divider()
    st.subheader("Importar planilha preenchida")

    st.info(
        "A importação abaixo substitui apenas os itens do setor escolhido, "
        "sem afetar os demais setores."
    )

    setor_importacao = st.selectbox(
        "Setor/proveniência para importação",
        [""] + PROVENIENCIAS_PADRAO,
        index=0,
        key="setor_importacao",
    )

    arquivo = st.file_uploader(
        "Selecione o arquivo Excel do inventário",
        type=["xlsx"],
    )

    if arquivo is not None:
        try:
            df_importado = parse_inventory_workbook(arquivo, df_ttd)
            st.write(f"{len(df_importado)} item(ns) localizado(s) na planilha.")
            st.dataframe(df_importado, use_container_width=True)

            if not df_importado.empty:
                if not setor_importacao.strip():
                    st.warning(
                        "Selecione o setor/proveniência antes de importar a planilha."
                    )
                else:
                    st.warning(
                        f"Confirme a importação para substituir apenas o inventário do setor: {setor_importacao}."
                    )

                    if st.button("Importar planilha para este setor"):
                        total = replace_inventory_from_dataframe_by_setor(
                            df_importado,
                            setor_importacao,
                        )
                        st.success(
                            f"Inventário do setor {setor_importacao} atualizado com {total} item(ns)."
                        )
        except Exception as e:
            st.error(f"Erro ao ler planilha: {e}")

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
                    file_name=f"inventario_{setor_escolhido}.xlsx".replace(" ", "_"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.markdown("### Exclusão em lote")

                opcoes = []
                mapa_ids = {}

                for _, row in df_inv.iterrows():
                    rotulo = (
                        f"#{row['id']} | "
                        f"{row.get('tipo_documental', '-') or '-'} | "
                        f"Caixa {row.get('caixa', '-') or '-'} | "
                        f"{row.get('datas_limite', '-') or '-'}"
                    )

                    opcoes.append(rotulo)
                    mapa_ids[rotulo] = int(row["id"])

                selecionar_todos = st.checkbox(
                    "Selecionar todos os itens deste setor"
                )

                selecionados = st.multiselect(
                    "Itens para excluir",
                    options=opcoes,
                    default=opcoes if selecionar_todos else [],
                )

                col1, col2 = st.columns(2)

                with col1:
                    if st.button("Excluir itens selecionados"):
                        ids = [mapa_ids[item] for item in selecionados]

                        total = delete_inventory_items(ids)

                        if total > 0:
                            st.success(f"{total} item(ns) excluído(s).")
                            st.rerun()
                        else:
                            st.warning("Nenhum item selecionado.")

                with col2:
                    if st.button("Excluir tudo deste setor"):
                        total = delete_inventory_items_by_setor(
                            setor_escolhido
                        )

                        if total > 0:
                            st.success(
                                f"Todos os {total} item(ns) do setor foram excluídos."
                            )
                            st.rerun()
                        else:
                            st.warning("Não havia itens para excluir.")