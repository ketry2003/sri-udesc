from pathlib import Path

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
import unicodedata

from services.search import load_ttd, get_filter_options, search_records
from services.ui_helpers import status_badge
from services.equivalencias import (
    buscar_equivalencia,
    salvar_equivalencia,
)


st.set_page_config(page_title="Consulta de Temporalidade", layout="wide")


def caminho_vocabulario():
    return (
        Path(__file__).resolve().parent.parent
        / "data"
        / "reference"
        / "planilha_atualizada.xlsx"
    )


def normalizar_colunas(df):
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("/", "_")
        .str.replace("-", "_")
    )
    return df


def remover_acentos(texto):
    return "".join(
        caractere
        for caractere in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(caractere) != "Mn"
    )


def normalizar_texto(texto):
    texto = (
        str(texto)
        .lower()
        .strip()
        .replace("\xa0", " ")
    )

    texto = remover_acentos(texto)

    return texto


@st.cache_data
def carregar_tesauro(tipo, arquivo_modificado=None):

    from services.db import carregar_vocabulario

    df = carregar_vocabulario(tipo)

    if df.empty:
        return pd.DataFrame()

    df = normalizar_colunas(df)

    colunas_busca = [
        "termo_encontrado",
        "termo_padronizado",
        "area",
        "subarea",
        "atividade",
        "assunto",
        "codigo_classificacao",
    ]

    colunas_existentes = [
        c for c in colunas_busca
        if c in df.columns
    ]

    if colunas_existentes:
        df["texto_busca"] = (
            df[colunas_existentes]
            .fillna("")
            .astype(str)
            .agg(" ".join, axis=1)
            .apply(normalizar_texto)
        )
    else:
        df["texto_busca"] = ""

    return df

def buscar_tesauro(texto, tipo, limite=10, corte=35):

    tesauro = carregar_tesauro(tipo)

    if tesauro.empty:
        return pd.DataFrame()

    termo = normalizar_texto(texto)

    if termo == "":
        return pd.DataFrame()

    tesauro = tesauro.copy()

    tesauro["score"] = tesauro["texto_busca"].apply(
        lambda texto_base: fuzz.WRatio(termo, texto_base)
    )

    resultado = (
        tesauro[tesauro["score"] >= corte]
        .sort_values("score", ascending=False)
        .head(limite)
    )

    st.write("DEBUG RESULTADO")
    st.write(resultado.head())
    st.write(f"Registros encontrados: {len(resultado)}")

    return resultado



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

st.caption(
    "Pesquise pelo nome do documento, processo, formulário, ata, edital, "
    "portaria, relatório ou assunto."
)

query = st.text_input(
    "Digite o nome do documento, processo ou assunto",
    placeholder=(
        "Ex.: edital de monitoria | ata de defesa | "
        "termo de compromisso de estágio | relatório final | portaria de banca"
    )
)

query_original = query

if query:

    termo_equivalente = buscar_equivalencia(query)

    if termo_equivalente:

        st.success(
            f"Equivalência histórica encontrada: "
            f"{termo_equivalente}"
        )

        query = termo_equivalente

# =========================
# SUGESTÕES DO VOCABULÁRIO
# =========================

if query:
    sugestoes = buscar_tesauro(query, tipo)

    if not sugestoes.empty:
        with st.expander("🔎 Sugestões do vocabulário controlado", expanded=True):

            primeira_linha = sugestoes.iloc[0]

            documento = primeira_linha.get(
            "termo_padronizado",
            ""
            )

    if not documento:
            documento = primeira_linha.get(
                "termo_encontrado",
        ""
        )

            # ==================================
            # SALVAR EQUIVALÊNCIA HISTÓRICA
            # ==================================

            if (
                query_original
                and documento
                and normalizar_texto(query_original)
                != normalizar_texto(documento)
            ):

                if st.button(
                    "💾 Salvar equivalência histórica",
                    key=f"salvar_eq_{query_original}"
                ):

                    salvar_equivalencia(
                        query_original,
                        documento
                    )

                    st.success(
                        f"""
Equivalência salva com sucesso:

{query_original}

→

{documento}
"""
                    )
            tipo_doc = primeira_linha.get(
                "atividade",
                ""
            )

            assunto = primeira_linha.get(
                "assunto",
                ""
            )

            codigo = primeira_linha.get(
                "codigo_classificacao",
                ""
            )

            st.success(
                f"""
Documento mais provável encontrado: **{documento}**

    if st.button(
        "📁 Usar esta classificação no Inventário",
        key=f"inventario_{codigo}"
    ):
        st.session_state.documento_selecionado = {
            "codigo_classificacao": codigo,
            "documento": documento,
            "assunto": assunto,
            "atividade": tipo_doc,
        }

        st.success(
            "Classificação enviada para o Inventário."
        )

Tipo documental: {tipo_doc}  
Assunto técnico: {assunto}  
Código de classificação: {codigo}
"""
            )

            colunas_exibir = [
                "termo_padronizado",
                "assunto",
                "codigo_classificacao",
            ]

            colunas_existentes = [
                c for c in colunas_exibir if c in sugestoes.columns
            ]

            sugestoes_exibir = sugestoes[colunas_existentes].rename(columns={
                "termo_padronizado": "Documento oficial",
                "assunto": "Assunto",
                "codigo_classificacao": "Código",
            })

            st.dataframe(
                sugestoes_exibir,
                use_container_width=True,
                hide_index=True
            )

    else:
        st.warning(
            "Nenhum termo encontrado no vocabulário controlado. "
            "Tente pesquisar pelo nome do documento, processo ou peça administrativa. "
            "Exemplos: 'ata de defesa', 'edital de monitoria', "
            "'termo de compromisso de estágio', 'relatório final', "
            "'portaria de banca', 'histórico escolar', "
            "'processo de jubilação', 'certificado de monitoria'."
        )

# =========================
# FILTROS AVANÇADOS
# =========================

filters = {}

with st.expander("Filtros avançados"):
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
            "Filtros técnicos para refinamento da atividade-fim."
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


filtros_ativos = {k: v for k, v in filters.items() if v}

if not query and not filtros_ativos:
    st.info("Digite um termo ou use os filtros avançados para iniciar a consulta.")
    st.stop()


# =========================
# RESULTADOS DA TTD
# =========================

results = search_records(
    df,
    query=query,
    filters=filtros_ativos,
    limit=100
)

st.write(f"{len(results)} resultado(s) exibido(s)")

if results.empty:
    st.info("Nenhum resultado encontrado na TTD.")

else:
    for _, row in results.iterrows():
        with st.container(border=True):

            left, right = st.columns([4, 1.3])

            left.subheader(row.get("item_documental", "") or "-")

            left.write(
                f"**Código de classificação:** "
                f"{row.get('codigo_classificacao', '') or '-'}"
            )

            if tipo == "meio":
                left.write(
                    f"**Natureza:** "
                    f"{row.get('natureza_documental', '') or '-'}"
                )
                left.write(
                    f"**Grupo:** "
                    f"{row.get('grupo', '') or '-'}"
                )
                left.write(
                    f"**Subgrupo:** "
                    f"{row.get('subgrupo', '') or '-'}"
                )
                left.write(
                    f"**Série:** "
                    f"{row.get('serie', '') or '-'}"
                )

            left.write(
                f"**Subsérie:** "
                f"{row.get('subserie', '') or '-'}"
            )

            left.write(
                f"**Dossiê / Processo:** "
                f"{row.get('dossie_processo', '') or '-'}"
            )

            if row.get("assunto_tecnico"):
                left.write(
                    f"**Assunto técnico:** "
                    f"{row.get('assunto_tecnico', '') or '-'}"
                )

            if row.get("funcao"):
                left.write(
                    f"**Função:** "
                    f"{row.get('funcao', '') or '-'}"
                )

            if row.get("subfuncao"):
                left.write(
                    f"**Subfunção:** "
                    f"{row.get('subfuncao', '') or '-'}"
                )

            if row.get("atividade"):
                left.write(
                    f"**Atividade:** "
                    f"{row.get('atividade', '') or '-'}"
                )

            right.write("**Temporalidade**")
            right.write(f"Corrente: {row.get('prazo_corrente', '') or '-'}")
            right.write(f"Intermediário: {row.get('prazo_intermediario', '') or '-'}")

            status_badge(row.get("destinacao_final", ""))

            if row.get("observacao"):
                st.info(f"Observação: {row.get('observacao')}")