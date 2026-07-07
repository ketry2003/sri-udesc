import pandas as pd
from pathlib import Path

ARQ = (
    Path(__file__).resolve().parent.parent
    / "data"
    / "reference"
    / "equivalencias_historicas.xlsx"
)

def salvar_equivalencia(
    termo_historico,
    termo_oficial,
    observacao=""
):

    df = carregar_equivalencias()

    nova_linha = pd.DataFrame([
        {
            "termo_historico": termo_historico,
            "termo_oficial": termo_oficial,
            "validado": "SIM",
            "observacao": observacao,
        }
    ])

    df = pd.concat(
        [df, nova_linha],
        ignore_index=True
    )

    df.drop_duplicates(
        subset=["termo_historico"],
        keep="last",
        inplace=True
    )

    df.to_excel(
        ARQ,
        index=False
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