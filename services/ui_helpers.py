from __future__ import annotations

import pandas as pd
import streamlit as st


def status_badge(destinacao: str):
    destino = str(destinacao or "").lower()
    if "permanente" in destino:
        st.success(f"Status: {destinacao}")
    elif "elimina" in destino:
        st.error(f"Status: {destinacao}")
    else:
        st.warning(f"Status: {destinacao or 'Exige análise'}")


def dataframe_from_rows(rows) -> pd.DataFrame:
    return pd.DataFrame([dict(row) for row in rows]) if rows else pd.DataFrame()