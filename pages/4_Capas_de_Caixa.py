from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate

from services.db import (
    list_setores_inventory,
    list_inventory_items_by_setor,
)

from services.ui_helpers import dataframe_from_rows


# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="Capas de Caixa",
    layout="wide"
)

st.title("Capas / Etiquetas de Caixa")


# =========================================================
# TEMPLATE WORD
# =========================================================

TEMPLATE_PATH = Path("data/templates/capa_caixa.docx")

if not TEMPLATE_PATH.exists():
    st.error(
        "Modelo Word não encontrado em:\n"
        "data/templates/capa_caixa.docx"
    )
    st.stop()


# =========================================================
# BUSCA INVENTÁRIO
# =========================================================

setores = list_setores_inventory()

if not setores:
    st.warning("Nenhum setor encontrado.")
    st.stop()

setor_escolhido = st.selectbox(
    "Selecione o setor",
    options=setores
)

rows = list_inventory_items_by_setor(setor_escolhido)

df = dataframe_from_rows(rows)

if df.empty:
    st.warning("Nenhum item encontrado.")
    st.stop()


# =========================================================
# CAIXAS DISPONÍVEIS
# =========================================================

df["caixa"] = df["caixa"].astype(str).str.strip()

caixas = sorted(
    df["caixa"]
    .dropna()
    .unique()
    .tolist()
)

caixas_escolhidas = st.multiselect(
    "Selecione as caixas",
    caixas
)

if not caixas_escolhidas:
    st.info("Selecione ao menos uma caixa.")
    st.stop()


# =========================================================
# PREVIEW
# =========================================================

preview = []

records = []

for caixa in caixas_escolhidas:

    caixa_df = df[df["caixa"] == str(caixa)].copy()

    if caixa_df.empty:
        continue

    # -------------------------------------
    # DATAS-LIMITE
    # -------------------------------------

    datas = [
        str(x).strip()
        for x in caixa_df["datas_limite"]
        .fillna("")
        .tolist()
        if str(x).strip()
    ]

    datas_limite = " / ".join(sorted(set(datas)))

    if not datas_limite:
        datas_limite = "-"

    # -------------------------------------
    # DESTINAÇÃO
    # -------------------------------------

    destinos = (
        caixa_df["destinacao_final"]
        .fillna("")
        .astype(str)
        .tolist()
    )

    permanente = any(
        "permanente" in d.lower()
        for d in destinos
    )

    destinacao = (
        "GUARDA PERMANENTE"
        if permanente
        else "ELIMINAÇÃO"
    )

    # -------------------------------------
    # CÓDIGO CLASSIFICAÇÃO
    # -------------------------------------

    codigo = "-"

    if "codigo_classificacao" in caixa_df.columns:

        codigos = (
            caixa_df["codigo_classificacao"]
            .fillna("")
            .astype(str)
            .tolist()
        )

        codigos = [
            c.strip()
            for c in codigos
            if c.strip()
        ]

        if codigos:
            codigo = codigos[0]

    # -------------------------------------
    # ASSUNTO
    # -------------------------------------

    assuntos = (
        caixa_df["tipo_documental"]
        .fillna("")
        .astype(str)
        .tolist()
    )

    assuntos = [
        a.strip()
        for a in assuntos
        if a.strip()
    ]

    assunto = " / ".join(sorted(set(assuntos[:4])))

    if not assunto:
        assunto = "-"

    # -------------------------------------
    # UNIDADE DOCUMENTAÇÃO
    # -------------------------------------

    unidade = setor_escolhido

    if "setor" in caixa_df.columns:

        setores_caixa = (
            caixa_df["setor"]
            .fillna("")
            .astype(str)
            .tolist()
        )

        setores_caixa = [
            s.strip()
            for s in setores_caixa
            if s.strip()
        ]

        if setores_caixa:
            unidade = setores_caixa[0]

    # -------------------------------------
    # RECORD TEMPLATE
    # -------------------------------------

    record = {
        "unidade_documentacao": unidade,
        "datas_limite_resumo": datas_limite,
        "destinacao": destinacao,
        "codigo_classificacao": codigo,
        "assunto": assunto,
        "numero_caixa": f"CAIXA {caixa}",
    }

    records.append(record)
    preview.append(record)


# =========================================================
# PREVIEW TABELA
# =========================================================

st.subheader("Pré-visualização")

preview_df = pd.DataFrame(preview)

st.dataframe(
    preview_df,
    use_container_width=True
)


# =========================================================
# GERAR DOCX
# =========================================================

def gerar_docx(records):

    doc = DocxTemplate(str(TEMPLATE_PATH))

    # renderiza primeira capa
    doc.render(records[0])

    # se quiser múltiplas caixas,
    # gera uma por vez em arquivos separados

    buffer = BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# DOWNLOAD
# =========================================================

if st.button("Gerar capas"):

    if len(records) == 1:

        arquivo = gerar_docx(records)

        st.download_button(
            "Baixar DOCX",
            data=arquivo,
            file_name=f"capa_caixa_{records[0]['numero_caixa']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    else:

        st.warning(
            "No momento gere uma caixa por vez.\n\n"
            "Depois podemos implementar múltiplas capas no mesmo DOCX."
        )