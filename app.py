import streamlit as st

from services.db import init_db
from services.data_loader import load_ttd

st.set_page_config(page_title="SRI Arquivístico UDESC - CCT", page_icon="📚", layout="wide")

init_db()

df = load_ttd()

st.title("SRI Arquivístico UDESC - CCT")
st.caption("Ferramenta de apoio para consulta de temporalidade, inventário, eliminação documental e capas de caixa.")

col1, col2, col3 = st.columns(3)
col1.metric("Registros TTD carregados", len(df))
col2.metric("Atividade-meio", int((df["natureza_documental"] == "Atividade-meio").sum()))
col3.metric("Atividade-fim", int((df["natureza_documental"] == "Atividade-fim").sum()))

st.markdown(
    """
### Módulos disponíveis
- **Consulta de temporalidade**: pesquise o documento e veja classificação, prazos e destinação.
- **Inventário**: monte um inventário com apoio automático da TTD.
- **Eliminação documental**: gere a listagem a partir dos itens validados.
- **Capas de caixa**: emita capas/etiquetas com a classificação completa.

Use o menu lateral para navegar.
"""
)
