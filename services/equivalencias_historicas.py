import pandas as pd

df = pd.DataFrame(
    columns=[
        "termo_historico",
        "termo_oficial",
        "validado",
        "observacao",
    ]
)

df.to_excel(
    "data/reference/equivalencias_historicas.xlsx",
    index=False
)

print("Arquivo criado")