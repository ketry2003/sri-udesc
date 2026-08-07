from pathlib import Path

import streamlit as st

from services.db import list_setores_inventory, list_inventory_items_by_setor
from services.ui_helpers import dataframe_from_rows
from services.exporters import (
    build_elimination_listing_dataframe,
    build_elimination_pdf,
    build_edital_pdf,
    build_termo_pdf,
    dataframe_to_excel_bytes,
)


def salvar_arquivo_local(nome_arquivo: str, conteudo: bytes) -> str:
    """
    Salva arquivo apenas no ambiente onde o app está rodando.
    Em ambiente local, salva na pasta ./saidas.
    Em nuvem, salva no servidor temporário.
    """
    pasta_saida = Path("saidas")
    pasta_saida.mkdir(exist_ok=True)
    caminho = pasta_saida / nome_arquivo
    with open(caminho, "wb") as f:
        f.write(conteudo)
    return str(caminho.resolve())


st.set_page_config(page_title="Eliminação Documental", layout="wide")
st.title("Eliminação documental")

st.caption(
    "Módulo baseado na IN SEA nº 10/2024: gera Anexo I (Listagem de Eliminação), "
    "Anexo II (Edital de Ciência de Eliminação) e Anexo III (Termo de Eliminação)."
)

setores = list_setores_inventory()

if not setores:
    st.info("Nenhum item disponível no inventário.")
    st.stop()

setor_escolhido = st.selectbox(
    "Selecione o setor / proveniência",
    [""] + setores,
    index=0,
)

if not setor_escolhido:
    st.info("Selecione um setor para visualizar os itens.")
    st.stop()

rows = list_inventory_items_by_setor(setor_escolhido)
df = dataframe_from_rows(rows)

if df.empty:
    st.info("Nenhum item disponível no inventário para este setor.")
    st.stop()

# Monta opções de seleção com identificação mais clara
opcoes = {
    (
        f"#{row.get('id', '-') or '-'} | "
        f"{row.get('tipo_documental', '-') or '-'} | "
        f"Caixa {row.get('caixa', '-') or '-'} | "
        f"{row.get('datas_limite', '-') or '-'}"
    ): row.to_dict()
    for _, row in df.iterrows()
}

selecionados = st.multiselect(
    "Selecione os itens para eliminação",
    options=list(opcoes.keys()),
    help="Somente itens com destinação final eliminável devem ser incluídos."
)

if not selecionados:
    st.info("Selecione um ou mais itens para montar a eliminação documental.")
    st.stop()

records = [opcoes[k] for k in selecionados]

# Bloqueia itens de guarda permanente
itens_bloqueados = []
for item in records:
    destinacao = str(item.get("destinacao_final", "")).lower()
    if "permanente" in destinacao:
        itens_bloqueados.append(item.get("tipo_documental", "Item sem nome"))

if itens_bloqueados:
    st.error(
        "Os itens abaixo possuem destinação de guarda permanente e não podem ser incluídos "
        "na eliminação documental:\n\n- "
        + "\n- ".join(itens_bloqueados)
        + "\n\nRevise a seleção no inventário e confirme a temporalidade no módulo "
        "'Consulta de temporalidade'."
    )
    st.stop()

with st.expander("Pré-visualizar itens selecionados", expanded=False):
    preview_df = build_elimination_listing_dataframe(records)
    st.dataframe(preview_df, use_container_width=True)

st.markdown("## Dados do processo e dos anexos")

col1, col2, col3 = st.columns(3)
orgao_entidade = col1.text_input("Órgão/Entidade", value="UDESC")
unidade_setor = col2.text_input("Unidade/Setor", value=setor_escolhido)
processo_numero = col3.text_input(
    "Processo (sigla, número/ano)",
    placeholder="Ex.: UDESC 12345/2026"
)

col4, col5, col6, col7 = st.columns(4)
listagem_numero = col4.text_input("Nº da Listagem", value="01")
listagem_ano = col5.text_input("Ano da Listagem", value="2026")
edital_numero = col6.text_input("Nº do Edital", value="01")
edital_ano = col7.text_input("Ano do Edital", value="2026")

col8, col9 = st.columns(2)
nome_titular_orgao = col8.text_input("Nome do Titular do Órgão")
cargo_titular_orgao = col9.text_input(
    "Cargo do Titular do Órgão",
    placeholder="Ex.: Diretor-Geral, Reitor(a)"
)

col10, col11 = st.columns(2)
presidente_cpad = col10.text_input("Nome do Presidente da CPAD")
cargo_presidente_cpad = col11.text_input(
    "Cargo do Presidente da CPAD",
    value="Presidente da CPAD"
)

col12, col13 = st.columns(2)
responsavel_selecao = col12.text_input("Responsável pela seleção")
cargo_responsavel_selecao = col13.text_input("Cargo do responsável pela seleção")

col14, col15, col16 = st.columns(3)
local = col14.text_input("Local", value="Joinville/SC")
data_local = col15.text_input("Data local", placeholder="Ex.: 08/04/2026")
doe_numero_data = col16.text_input(
    "DOE nº e data de publicação",
    placeholder="Ex.: DOE nº 22401, de 22/11/2024"
)

col17, col18 = st.columns(2)
data_eliminacao_extenso = col17.text_input(
    "Data da eliminação (para o Termo)",
    placeholder="Ex.: 08 de abril de 2026"
)
conjuntos_documentais = col18.text_input(
    "Conjuntos documentais (para o Edital)",
    placeholder="Ex.: cronogramas de horário de aula; atas; relatórios"
)

datas_limite_gerais = st.text_input(
    "Datas-limite gerais",
    placeholder="Ex.: 2010-2015/2018"
)

meta = {
    "orgao_entidade": orgao_entidade,
    "unidade_setor": unidade_setor,
    "processo_numero": processo_numero,
    "listagem_numero": listagem_numero,
    "listagem_ano": listagem_ano,
    "edital_numero": edital_numero,
    "edital_ano": edital_ano,
    "nome_titular_orgao": nome_titular_orgao,
    "cargo_titular_orgao": cargo_titular_orgao,
    "responsavel_orgao": nome_titular_orgao,
    "cargo_responsavel_orgao": cargo_titular_orgao,
    "presidente_cpad": presidente_cpad,
    "cargo_presidente_cpad": cargo_presidente_cpad,
    "responsavel_selecao": responsavel_selecao,
    "cargo_responsavel_selecao": cargo_responsavel_selecao,
    "local": local,
    "data_local": data_local,
    "doe_numero_data": doe_numero_data,
    "data_eliminacao_extenso": data_eliminacao_extenso,
    "conjuntos_documentais": conjuntos_documentais,
    "datas_limite_gerais": datas_limite_gerais,
}

st.markdown("## Gerar arquivos")

col_a, col_b = st.columns([1, 1])

with col_a:
    if st.button("Gerar Anexo I, II e III"):
        campos_obrigatorios = {
            "Órgão/Entidade": orgao_entidade,
            "Unidade/Setor": unidade_setor,
            "Processo": processo_numero,
            "Nº da Listagem": listagem_numero,
            "Ano da Listagem": listagem_ano,
            "Nº do Edital": edital_numero,
            "Ano do Edital": edital_ano,
            "Local": local,
        }

        faltantes = [
            campo for campo, valor in campos_obrigatorios.items()
            if not str(valor).strip()
        ]

        if faltantes:
            st.error(
                "Preencha os campos obrigatórios antes de gerar os arquivos:\n\n- "
                + "\n- ".join(faltantes)
            )
        else:
            try:
                anexo_i_pdf = build_elimination_pdf(records, meta)
                anexo_ii_pdf = build_edital_pdf(records, meta)
                anexo_iii_pdf = build_termo_pdf(records, meta)
                eliminacao_excel = dataframe_to_excel_bytes(
                    build_elimination_listing_dataframe(records),
                    sheet_name="Listagem"
                )

                nome_anexo_i = (
                    f"anexo_i_listagem_eliminacao_{listagem_numero}_{listagem_ano}.pdf"
                )
                nome_anexo_ii = (
                    f"anexo_ii_edital_ciencia_{edital_numero}_{edital_ano}.pdf"
                )
                nome_anexo_iii = (
                    f"anexo_iii_termo_eliminacao_{edital_numero}_{edital_ano}.pdf"
                )
                nome_excel = (
                    f"listagem_eliminacao_{listagem_numero}_{listagem_ano}.xlsx"
                )

                st.session_state["anexo_i_pdf"] = anexo_i_pdf
                st.session_state["anexo_ii_pdf"] = anexo_ii_pdf
                st.session_state["anexo_iii_pdf"] = anexo_iii_pdf
                st.session_state["eliminacao_excel"] = eliminacao_excel

                st.session_state["caminho_anexo_i"] = salvar_arquivo_local(
                    nome_anexo_i, anexo_i_pdf
                )
                st.session_state["caminho_anexo_ii"] = salvar_arquivo_local(
                    nome_anexo_ii, anexo_ii_pdf
                )
                st.session_state["caminho_anexo_iii"] = salvar_arquivo_local(
                    nome_anexo_iii, anexo_iii_pdf
                )
                st.session_state["caminho_excel"] = salvar_arquivo_local(
                    nome_excel, eliminacao_excel
                )

                st.success("Arquivos gerados com sucesso.")
            except Exception as e:
                st.error(f"Erro ao gerar os arquivos: {e}")

with col_b:
    st.info(
        "Preencha os dados do processo e clique em 'Gerar Anexo I, II e III'. "
        "Depois use os botões de download abaixo."
    )

if "anexo_i_pdf" in st.session_state:
    st.markdown("### Downloads")

    st.download_button(
        "Baixar Anexo I - Listagem (PDF)",
        data=st.session_state["anexo_i_pdf"],
        file_name=f"anexo_i_listagem_eliminacao_{listagem_numero}_{listagem_ano}.pdf",
        mime="application/pdf",
        key="download_anexo_i_pdf",
    )

    st.download_button(
        "Baixar Anexo II - Edital (PDF)",
        data=st.session_state["anexo_ii_pdf"],
        file_name=f"anexo_ii_edital_ciencia_{edital_numero}_{edital_ano}.pdf",
        mime="application/pdf",
        key="download_anexo_ii_pdf",
    )

    st.download_button(
        "Baixar Anexo III - Termo (PDF)",
        data=st.session_state["anexo_iii_pdf"],
        file_name=f"anexo_iii_termo_eliminacao_{edital_numero}_{edital_ano}.pdf",
        mime="application/pdf",
        key="download_anexo_iii_pdf",
    )

    st.download_button(
        "Baixar planilha base da listagem (Excel)",
        data=st.session_state["eliminacao_excel"],
        file_name=f"listagem_eliminacao_{listagem_numero}_{listagem_ano}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_eliminacao_excel",
    )

    with st.expander("Detalhes técnicos dos arquivos gerados", expanded=False):
        st.warning(
            "Os caminhos abaixo referem-se ao ambiente onde o app está rodando. "
            "Na nuvem, eles não representam o seu computador."
        )
        st.code(st.session_state.get("caminho_anexo_i", ""))
        st.code(st.session_state.get("caminho_anexo_ii", ""))
        st.code(st.session_state.get("caminho_anexo_iii", ""))
        st.code(st.session_state.get("caminho_excel", ""))