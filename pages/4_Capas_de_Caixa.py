import streamlit as st

from config import FUNDO_PADRAO, OFFICIAL_BOX_COVER_NAME
from services.db import list_setores_inventory, list_inventory_items_by_setor
from services.exporters import build_box_covers_from_template_docx_bytes
from services.ui_helpers import dataframe_from_rows

st.set_page_config(page_title="Capas / Etiquetas de Caixa", layout="wide")
st.title("Capas / etiquetas de caixa")

st.caption(
    "As capas aproveitam automaticamente os dados do inventário e da classificação já associada aos documentos."
)

setores = list_setores_inventory()

if not setores:
    st.info("Cadastre itens no inventário antes de gerar capas.")
    st.stop()

setor_escolhido = st.selectbox(
    "Selecione o setor / proveniência",
    [""] + setores,
    index=0,
)

if not setor_escolhido:
    st.info("Selecione um setor para visualizar as caixas.")
    st.stop()

rows = list_inventory_items_by_setor(setor_escolhido)
inv_df = dataframe_from_rows(rows)

if inv_df.empty:
    st.info("Nenhum item salvo para este setor.")
    st.stop()

# Mantém apenas linhas com número de caixa
inv_df = inv_df.copy()
inv_df["caixa"] = inv_df["caixa"].astype(str).str.strip()
inv_df = inv_df[inv_df["caixa"] != ""]

if inv_df.empty:
    st.info("Nenhum item do inventário deste setor possui número de caixa informado.")
    st.stop()

fundo = st.text_input("Fundo", value=FUNDO_PADRAO)
orgao_entidade = st.text_input("Órgão/Entidade", value="UDESC")
unidade_macro = st.text_input("Centro / Reitoria", value=setor_escolhido)
setor_responsavel = st.text_input("Setor responsável", value=setor_escolhido)

caixas_disponiveis = sorted(inv_df["caixa"].dropna().astype(str).unique().tolist())

selecionadas = st.multiselect(
    "Selecione as caixas para gerar a capa",
    caixas_disponiveis,
    format_func=lambda x: f"Caixa {x}"
)

if not selecionadas:
    st.info("Selecione uma ou mais caixas.")
    st.stop()

records = []
preview_rows = []

for caixa in selecionadas:
    caixa_df = inv_df[inv_df["caixa"].astype(str) == str(caixa)].copy()

    if caixa_df.empty:
        continue

    # Datas-limite consolidadas
    datas = [
        str(v).strip()
        for v in caixa_df["datas_limite"].fillna("").tolist()
        if str(v).strip()
    ]
    datas_limite = " / ".join(sorted(set(datas))) if datas else "-"

    # Se algum item for permanente, a caixa vira permanente
    destinos = caixa_df["destinacao_final"].fillna("").astype(str).tolist()
    guarda_permanente = any("permanente" in d.lower() for d in destinos)

    # Código e tipo documental principal
    codigo_classificacao = (
        caixa_df["codigo_classificacao"].fillna("").astype(str).iloc[0]
        if "codigo_classificacao" in caixa_df.columns and not caixa_df.empty
        else ""
    )

    tipos = [
        str(v).strip()
        for v in caixa_df["tipo_documental"].fillna("").tolist()
        if str(v).strip()
    ]
    assunto = " ; ".join(sorted(set(tipos[:3]))) if tipos else "-"
    if len(set(tipos)) > 3:
        assunto += " ..."

    # Temporalidade consolidada
    prazos_correntes = (
        caixa_df["prazo_corrente"].fillna("").astype(str).tolist()
        if "prazo_corrente" in caixa_df.columns
        else []
    )
    prazos_intermediarios = (
        caixa_df["prazo_intermediario"].fillna("").astype(str).tolist()
        if "prazo_intermediario" in caixa_df.columns
        else []
    )

    prazo_corrente = " / ".join(
        sorted(set([p for p in prazos_correntes if p.strip()]))
    ) or "-"
    prazo_intermediario = " / ".join(
        sorted(set([p for p in prazos_intermediarios if p.strip()]))
    ) or "-"

    destinacao = "GUARDA PERMANENTE" if guarda_permanente else "ELIMINAÇÃO"
    destaque_permanente = "GUARDA PERMANENTE" if guarda_permanente else ""

    # Proveniência / setor da caixa
    unidade_documentacao = (
        caixa_df["setor"].fillna("").astype(str).iloc[0]
        if "setor" in caixa_df.columns and not caixa_df.empty
        else ""
    )

    observacoes = [
        str(v).strip()
        for v in caixa_df["observacoes"].fillna("").tolist()
        if str(v).strip()
    ]
    observacao = " | ".join(sorted(set(observacoes[:3]))) if observacoes else ""

    record = {
        "fundo": fundo,
        "orgao_entidade": orgao_entidade,
        "unidade_macro": unidade_macro,
        "setor_responsavel": setor_responsavel,
        "unidade_documentacao": unidade_documentacao,
        "codigo_classificacao": codigo_classificacao or "-",
        "assunto": assunto,
        "datas_limite_resumo": datas_limite,
        "datas_limite_detalhadas": datas_limite,
        "prazo_corrente": prazo_corrente,
        "prazo_intermediario": prazo_intermediario,
        "destinacao": destinacao,
        "destaque_permanente": destaque_permanente,
        "numero_caixa": str(caixa),
        "observacao": observacao,
        "lista_itens_caixa": "\n".join(sorted(set(tipos))) if tipos else "",
    }
    records.append(record)
    preview_rows.append(record)

st.subheader("Pré-visualização das capas")
st.dataframe(dataframe_from_rows(preview_rows), use_container_width=True)

st.download_button(
    "Baixar capas em DOCX",
    data=build_box_covers_from_template_docx_bytes(records),
    file_name=OFFICIAL_BOX_COVER_NAME,
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)