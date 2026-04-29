from pathlib import Path

import pandas as pd
import streamlit as st

from services.search import load_ttd, get_filter_options, search_records
from services.ui_helpers import status_badge


st.set_page_config(page_title="Consulta de Temporalidade", layout="wide")


@st.cache_data
def carregar_tesauro():
    arquivo = (
        Path(__file__).resolve().parent.parent
        / "data"
        / "reference"
        / "vocabulario.xlsx"
    )

    if not arquivo.exists():
        st.error("Arquivo vocabulario.xlsx não encontrado em data/reference.")
        return pd.DataFrame()

    return pd.read_excel(arquivo, sheet_name="Busca Geral")


def buscar_tesauro(texto):
    tesauro = carregar_tesauro()

    if tesauro.empty:
        return pd.DataFrame()

    termo = texto.lower().strip()

    if termo == "":
        return pd.DataFrame()

    resultado = tesauro[
        tesauro.apply(
            lambda linha: linha.astype(str)
            .str.lower()
            .str.contains(termo, na=False, regex=False)
            .any(),
            axis=1
        )
    ]

    return resultado.head(10)


st.title("Consulta de temporalidade")

if "tipo" not in st.session_state:
    st.session_state.tipo = "fim"

tipo = st.selectbox(
    "Selecione o tipo de atividade:",
    options=["meio", "fim"],
    index=1,
    format_func=lambda x: "Atividade-meio" if x == "meio" else "Atividade-fim",
    help="""
Atividade-meio: funções administrativas
(RH, compras, contratos, patrimônio, etc).
Atividade-fim: funções ligadas ao ensino, pesquisa, extensão
e atividades acadêmicas.
"""
)

st.session_state.tipo = tipo

df = load_ttd(tipo)

st.caption("Pesquise por tipo documental ou pelo número/código de classificação.")

query = st.text_input(
    "Busca livre por tipo documental ou código de classificação",
    placeholder="Ex.: diário de classe | CI | PTI | 06.24.01 | 06.24.01.02.01.001"
)

# Sugestões do Tesauro somente para atividade-fim
if query and tipo == "fim":
    sugestoes = buscar_tesauro(query)

    if not sugestoes.empty:
        with st.expander("🔎 Sugestões do Tesauro UDESC", expanded=True):
            colunas_exibir = [
                "Termo encontrado / possível busca",
                "Termo padronizado sugerido",
                "Área",
                "Subárea",
                "Observação para inventário",
            ]

            colunas_existentes = [
                coluna for coluna in colunas_exibir if coluna in sugestoes.columns
            ]

            st.dataframe(
                sugestoes[colunas_existentes],
                use_container_width=True,
                hide_index=True
            )

            termo_sugerido = sugestoes.iloc[0].get(
                "Termo padronizado sugerido",
                ""
            )

            if termo_sugerido:
                st.info(
                    f"Termo padronizado mais provável: **{termo_sugerido}**"
                )

    else:
        st.warning(
            "Nenhuma equivalência encontrada no tesauro. "
            "Se necessário, registrar como pendente de classificação."
        )

filters = {}

if tipo == "meio":
    cols = st.columns(3)

    filters["natureza_documental"] = cols[0].selectbox(
        "Natureza",
        [""] + get_filter_options(df, "natureza_documental"),
        index=0
    )

    filters["grupo"] = cols[1].selectbox(
        "Grupo",
        [""] + get_filter_options(
            df,
            "grupo",
            {"natureza_documental": filters["natureza_documental"]}
        ),
        index=0
    )

    filters["subgrupo"] = cols[2].selectbox(
        "Subgrupo",
        [""] + get_filter_options(
            df,
            "subgrupo",
            {k: v for k, v in filters.items() if v}
        ),
        index=0
    )

    cols2 = st.columns(3)

    filters["serie"] = cols2[0].selectbox(
        "Série",
        [""] + get_filter_options(
            df,
            "serie",
            {k: v for k, v in filters.items() if v}
        ),
        index=0
    )

    filters["subserie"] = cols2[1].selectbox(
        "Subsérie",
        [""] + get_filter_options(
            df,
            "subserie",
            {k: v for k, v in filters.items() if v}
        ),
        index=0
    )

    filters["dossie_processo"] = cols2[2].selectbox(
        "Dossiê / Processo",
        [""] + get_filter_options(
            df,
            "dossie_processo",
            {k: v for k, v in filters.items() if v}
        ),
        index=0
    )

else:
    st.info(
        "Na atividade-fim, a consulta foi simplificada para focar em "
        "subsérie e dossiê/processo."
    )

    cols = st.columns(2)

    filters["subserie"] = cols[0].selectbox(
        "Subsérie",
        [""] + get_filter_options(df, "subserie"),
        index=0
    )

    filters["dossie_processo"] = cols[1].selectbox(
        "Dossiê / Processo",
        [""] + get_filter_options(
            df,
            "dossie_processo",
            {k: v for k, v in filters.items() if v}
        ),
        index=0
    )

results = search_records(
    df,
    query=query,
    filters={k: v for k, v in filters.items() if v},
    limit=30
)

st.write(f"{len(results)} resultado(s)")

if results.empty:
    st.info("Nenhum resultado encontrado na TTD.")
else:
    for _, row in results.iterrows():
        with st.container(border=True):
            left, right = st.columns([4, 1.3])

            left.subheader(row.get("item_documental", "") or "-")
            left.write(
                f"**Código de classificação:** {row.get('codigo_classificacao', '') or '-'}"
            )

            if tipo == "meio":
                left.write(f"**Natureza:** {row.get('natureza_documental', '') or '-'}")
                left.write(f"**Grupo:** {row.get('grupo', '') or '-'}")
                left.write(f"**Subgrupo:** {row.get('subgrupo', '') or '-'}")
                left.write(f"**Série:** {row.get('serie', '') or '-'}")

            left.write(f"**Subsérie:** {row.get('subserie', '') or '-'}")
            left.write(f"**Dossiê / Processo:** {row.get('dossie_processo', '') or '-'}")

            with st.expander("Ver detalhamento da classificação"):
                if tipo == "meio":
                    left_codigo_grupo = row.get("grupo_codigo", "") or "-"
                    left_codigo_subgrupo = row.get("subgrupo_codigo", "") or "-"
                    left_codigo_serie = row.get("serie_codigo", "") or "-"

                    st.write(f"**Código do grupo:** {left_codigo_grupo}")
                    st.write(f"**Código do subgrupo:** {left_codigo_subgrupo}")
                    st.write(f"**Código da série:** {left_codigo_serie}")

                st.write(
                    f"**Código da subsérie:** {row.get('subserie_codigo', '') or '-'}"
                )
                st.write(
                    f"**Código do dossiê/processo:** {row.get('dossie_processo_codigo', '') or '-'}"
                )
                st.write(
                    f"**Código do item documental:** {row.get('item_documental_codigo', '') or '-'}"
                )

            right.write("**Temporalidade**")
            right.write(f"Corrente: {row.get('prazo_corrente', '') or '-'}")
            right.write(f"Intermediário: {row.get('prazo_intermediario', '') or '-'}")

            status_badge(row.get("destinacao_final", ""))

            if row.get("observacao"):
                st.info(f"Observação: {row.get('observacao')}")