import streamlit as st
from services.search import load_ttd, get_filter_options, search_records
from services.ui_helpers import status_badge

st.set_page_config(page_title="Consulta de Temporalidade", layout="wide")

st.title("Consulta de temporalidade")

df = load_ttd()

st.caption("Pesquise por tipo documental ou pelo número/código de classificação.")

query = st.text_input(
    "Busca livre por tipo documental ou código de classificação",
    placeholder="Ex.: diário de classe | 06.24.01 | 06.24.01.02.01.001"
)

filters = {}

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

results = search_records(
    df,
    query=query,
    filters={k: v for k, v in filters.items() if v},
    limit=30
)

st.write(f"{len(results)} resultado(s)")

if results.empty:
    st.info("Nenhum resultado encontrado.")
else:
    for _, row in results.iterrows():
        with st.container(border=True):
            left, right = st.columns([4, 1.3])

            left.subheader(row.get("item_documental", "") or "-")
            left.write(f"**Código de classificação:** {row.get('codigo_classificacao', '') or '-'}")
            left.write(f"**Natureza:** {row.get('natureza_documental', '') or '-'}")
            left.write(f"**Grupo:** {row.get('grupo', '') or '-'}")
            left.write(f"**Subgrupo:** {row.get('subgrupo', '') or '-'}")
            left.write(f"**Série:** {row.get('serie', '') or '-'}")
            left.write(f"**Subsérie:** {row.get('subserie', '') or '-'}")
            left.write(f"**Dossiê / Processo:** {row.get('dossie_processo', '') or '-'}")

            with st.expander("Ver detalhamento da classificação"):
                st.write(f"**Código do grupo:** {row.get('grupo_codigo', '') or '-'}")
                st.write(f"**Código do subgrupo:** {row.get('subgrupo_codigo', '') or '-'}")
                st.write(f"**Código da série:** {row.get('serie_codigo', '') or '-'}")
                st.write(f"**Código da subsérie:** {row.get('subserie_codigo', '') or '-'}")
                st.write(f"**Código do dossiê/processo:** {row.get('dossie_processo_codigo', '') or '-'}")
                st.write(f"**Código do item documental:** {row.get('item_documental_codigo', '') or '-'}")

            right.write("**Temporalidade**")
            right.write(f"Corrente: {row.get('prazo_corrente', '') or '-'}")
            right.write(f"Intermediário: {row.get('prazo_intermediario', '') or '-'}")

            status_badge(row.get("destinacao_final", ""))

            if row.get("observacao"):
                st.info(f"Observação: {row.get('observacao')}")