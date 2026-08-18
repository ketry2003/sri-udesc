import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent / "data" / "reference"

df = pd.read_excel(
    BASE / "planilha_atualizada_v2.xlsx",
    engine="openpyxl"
)

novos = df[df["id_sistema"].isna()]

print("NOVOS:", len(novos))

print(
    novos[
        [
            "termo_preferido_oficial",
            "prazo_corrente",
            "prazo_intermediario",
            "destinacao"
        ]
    ].head(30)
)