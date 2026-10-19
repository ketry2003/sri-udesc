import pandas as pd
import streamlit as st

from services.db import search_inventory_global, list_setores_inventory
from services.ui_helpers import dataframe_from_rows
from services.exporters import dataframe_to_excel_bytes

st.set_page_config(page_title="Consulta Geral do Inventário", layout="wide")

st.title("🔍 Consulta de Documentos no Inventário")
st.caption(
    "Utilize os campos abaixo para pesquisar por termos gerais em todo o inventário "
    "ou faça uma busca direcionada no campo de observações, referências e assuntos."
)

st.divider()

# Layout de filtros de busca
col1, col2 = st.columns(2)

with col1:
    busca_geral = st.text_input(
        "Busca simples (todos os campos)",
        placeholder="Ex.: Orbenk, Licitação, Caixa 001, 2024",
        help="Procura o termo digitado em qualquer informação cadastrada no inventário."
    )

with col2:
    busca_focada = st.text_input(
        "Busca focada (Observações / Detalhes)",
        placeholder="Ex.: Contrato 2000, Pregão 12",
        help="Filtra especificamente dentro dos textos informados nas observações."
    )

# Botão para disparar a consulta
if st.button("Pesquisar", type="primary"):
    resultados = search_inventory_global(busca_geral, busca_focada)
    
    # Salva os resultados na sessão para manter na tela caso interaja com a página
    st.session_state["resultados_consulta"] = resultados

# Recupera os resultados da sessão se existirem
resultados_atuais = st.session_state.get("resultados_consulta", [])

if resultados_atuais:
    df_resultados = dataframe_from_rows(resultados_atuais)
    
    st.success(f"✨ Foram encontrados **{len(df_resultados)}** item(ns) correspondentes.")
    
    # Exibe a tabela interativa
    st.dataframe(df_resultados, use_container_width=True)
    
    # Botão de exportação para Excel dos resultados filtrados
    st.download_button(
        "📥 Baixar resultados da consulta em Excel",
        data=dataframe_to_excel_bytes(df_resultados, "Consulta_Inventario"),
        file_name="consulta_inventario_documental.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Baixe uma planilha contendo exatamente os registros encontrados nesta busca."
    )
elif "resultados_consulta" in st.session_state:
    st.warning("⚠️ Nenhum registro encontrado com os critérios informados.")