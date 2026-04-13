from __future__ import annotations

import pandas as pd
import streamlit as st

from services.data_loader import load_ttd


df = load_ttd()

st.title("SRI Arquivístico UDESC - CCT")
st.caption("Ferramenta de apoio para consulta de temporalidade, inventário, eliminação documental e capas de caixa.")

col1, col2, col3 = st.columns(3)
col1.metric("Registros TTD carregados", len(df))
col2.metric(
    "Atividade-meio",
    int((df.get("natureza_documental", pd.Series(index=df.index, dtype="object")) == "Atividade-meio").sum())
)
col3.metric(
    "Atividade-fim",
    int((df.get("natureza_documental", pd.Series(index=df.index, dtype="object")) == "Atividade-fim").sum())
)

st.markdown(
    """
### Módulos
Use o menu lateral para acessar:
- Consulta de temporalidade
- Inventário
- Eliminação documental
- Capas de caixa
"""
)