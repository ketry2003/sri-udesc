
import streamlit as st

from config import FUNDO_PADRAO, OFFICIAL_BOX_COVER_NAME
from services.db import list_inventory_items
from services.exporters import build_box_covers_docx_bytes
from services.ui_helpers import dataframe_from_rows

st.title("Capas / etiquetas de caixa")
rows = list_inventory_items()
inv_df = dataframe_from_rows(rows)

st.caption("As capas aproveitam automaticamente os dados do inventário e da classificação já associada ao item.")

if inv_df.empty:
    st.info("Cadastre itens no inventário antes de gerar capas.")
else:
    fundo = st.text_input("Fundo", value=FUNDO_PADRAO)
    ids = inv_df["id"].astype(str).tolist()
    selecionados = st.multiselect(
        "Itens para gerar capa",
        ids,
        format_func=lambda x: f"{x} - {inv_df.loc[inv_df['id'].astype(str) == x, 'tipo_documental'].iloc[0]}"
    )
    if selecionados:
        records = []
        for item_id in selecionados:
            row = inv_df[inv_df["id"].astype(str) == item_id].iloc[0].to_dict()
            row["fundo"] = fundo
            records.append(row)
        st.dataframe(dataframe_from_rows(records), use_container_width=True)
        st.download_button(
            "Baixar capas em DOCX",
            data=build_box_covers_docx_bytes(records),
            file_name=OFFICIAL_BOX_COVER_NAME,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
