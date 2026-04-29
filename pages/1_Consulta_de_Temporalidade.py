from pathlib import Path

import pandas as pd
import streamlit as st

from services.search import load_ttd, get_filter_options, search_records
from services.ui_helpers import status_badge


st.set_page_config(page_title="Consulta de Temporalidade", layout="wide")


@st.cache_data
def carregar_tesauro(tipo):
    arquivo = (
        Path(__file__).resolve().parent.parent
        / "data"
        / "reference"
        / "vocabulario_controlado.xlsx"
    )

    if not arquivo.exists():
        st.error("Arquivo vocabulario_controlado.xlsx não encontrado em data/reference.")
        return pd.DataFrame()

    df = pd.read_excel(arquivo)

    # Padroniza nomes das colunas
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("/", "_")
    )

    # Identifica a coluna que informa se é meio ou fim
    coluna_tipo = None

    for possivel in ["tipo", "tipo_de_atividade", "atividade"]:
        if possivel in df.columns:
            coluna_tipo = possivel
            break

    # Filtra pelo tipo selecionado: meio ou fim
    if coluna_tipo:
        df[coluna_tipo] = (
            df[coluna_tipo]
            .astype(str)
            .str.lower()
            .str.strip()
            .str.replace("atividade-", "")
            .str.replace("atividade_", "")
            .str.replace("atividade ", "")
        )

        df = df[df[coluna_tipo] == tipo]

    return df


def buscar_tesauro(texto, tipo):
    tesauro = carregar_tesauro(tipo)

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
    placeholder="Ex.: PTI | férias | CI | contrato | estágio | diário de classe"
)

# TESAURO / VOCABULÁRIO CONTROLADO
if query:
    sugestoes = buscar_tesauro(query, tipo)

    if not sugestoes.empty:
        with st.expander("🔎 Sugestões do Vocabulário Controlado", expanded=True):

            termo_sugerido = sugestoes.iloc[0].get("termo_padronizado", "")

            if termo_sugerido:
                st.success(
                    f"Termo padronizado mais provável: **{termo_sugerido}**"
                )

            colunas_exibir = [
                "termo_encontrado",
                "termo_padronizado",
                "codigo_ttd",
                "subarea",
                "observacao",
                "observacao_para_inventario",
            ]

            colunas_existentes = [
                c for c in colunas_exibir if c in sugestoes.columns
            ]

            st.dataframe(
                sugestoes[colunas_existentes],
                use_container_width=True,
                hide_index=True
            )

    else:
        st.warning(
            "Nenhum termo encontrado no vocabulário controlado. "
            "Tente pesquisar por assunto geral."
        )

filters = {}

# FILTROS
if tipo == "meio":
    cols = st.columns(3)

    filters["natureza_documental"] = cols[0].selectbox(
        "Natureza",
        [""] + get_filter_options(df, "natureza_documental")
    )

    filters["grupo"] = cols[1].selectbox(
        "Grupo",
        [""] + get_filter_options(
            df,
            "grupo",
            {"natureza_documental": filters["natureza_documental"]}
        )
    )

    filters["subgrupo"] = cols[2].selectbox(
        "Subgrupo",
        [""] + get_filter_options(
            df,
            "subgrupo",
            {k: v for k, v in filters.items() if v}
        )
    )

    cols2 = st.columns(3)

    filters["serie"] = cols2[0].selectbox(
        "Série",
        [""] + get_filter_options(
            df,
            "serie",
            {k: v for k, v in filters.items() if v}
        )
    )

    filters["subserie"] = cols2[1].selectbox(
        "Subsérie",
        [""] + get_filter_options(
            df,
            "subserie",
            {k: v for k, v in filters.items() if v}
        )
    )

    filters["dossie_processo"] = cols2[2].selectbox(
        "Dossiê / Processo",
        [""] + get_filter_options(
            df,
            "dossie_processo",
            {k: v for k, v in filters.items() if v}
        )
    )

else:
    st.info(
        "Na atividade-fim, a consulta foi simplificada para focar em "
        "subsérie e dossiê/processo."
    )

    cols = st.columns(2)

    filters["subserie"] = cols[0].selectbox(
        "Subsérie",
        [""] + get_filter_options(df, "subserie")
    )

    filters["dossie_processo"] = cols[1].selectbox(
        "Dossiê / Processo",
        [""] + get_filter_options(
            df,
            "dossie_processo",
            {k: v for k, v in filters.items() if v}
        )
    )

# RESULTADOS TTD
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

            right.write("**Temporalidade**")
            right.write(f"Corrente: {row.get('prazo_corrente', '') or '-'}")
            right.write(f"Intermediário: {row.get('prazo_intermediario', '') or '-'}")

            status_badge(row.get("destinacao_final", ""))

            if row.get("observacao"):
                st.info(f"Observação: {row.get('observacao')}")