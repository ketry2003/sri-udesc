import pandas as pd
from pathlib import Path

ARQ = (
    Path(__file__).resolve().parent.parent
    / "data"
    / "reference"
    / "equivalencias_historicas.xlsx"
)

def buscar_equivalencia(termo):

    if not ARQ.exists():
        return None

    df = pd.read_excel(
        ARQ,
        engine="openpyxl"
    )

    termo = str(termo).strip().lower()

    resultado = df[
        df["termo_historico"]
        .astype(str)
        .str.strip()
        .str.lower()
        == termo
    ]

    if resultado.empty:
        return None

    return resultado.iloc[0]["termo_oficial"]