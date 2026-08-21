from pathlib import Path
import re
import unicodedata

import pandas as pd
from rapidfuzz import fuzz


BASE = Path(__file__).resolve().parent / "data" / "reference"

ARQUIVO_TTD = BASE / "ttd_atual.xlsx"
ARQUIVO_VOC = BASE / "planilha_atualizada.xlsx"


def normalizar(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto).strip().lower()

    texto = "".join(
        c
        for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )

    texto = re.sub(
        r"^\d+\s*-\s*",
        "",
        texto
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto
    )

    return texto.strip()


print("Carregando arquivos...")

ttd = pd.read_excel(
    ARQUIVO_TTD,
    engine="openpyxl"
)

voc = pd.read_excel(
    ARQUIVO_VOC,
    sheet_name="base_adaptada",
    engine="openpyxl"
)

print("\nEXEMPLOS TTD:")
print(
    ttd["item_documental"]
    .dropna()
    .head(10)
    .tolist()
)

print("\nEXEMPLOS VOCABULÁRIO:")
print(
    voc["termo_padronizado"]
    .dropna()
    .head(10)
    .tolist()
)

# --------------------------------------------------
# DOCUMENTOS DA TTD
# --------------------------------------------------

docs_ttd = []

for valor in ttd["item_documental"].dropna():
    docs_ttd.append(
        (
            str(valor).strip(),
            normalizar(valor)
        )
    )

# --------------------------------------------------
# VOCABULÁRIO
# --------------------------------------------------

docs_vocabulario = set()

for coluna in [
    "termo_padronizado",
    "termo_encontrado",
    "termos_populares_sugeridos",
    "assunto_tecnico"
]:
    if coluna in voc.columns:

        docs_vocabulario.update(
            {
                normalizar(x)
                for x in voc[coluna].dropna()
            }
        )

# --------------------------------------------------
# AUDITORIA POR SIMILARIDADE
# --------------------------------------------------

faltantes = []

for original, documento in docs_ttd:

    melhor_score = 0

    for termo in docs_vocabulario:

        score = fuzz.token_sort_ratio(
            documento,
            termo
        )

        melhor_termo = ""

        if score > melhor_score:
            melhor_score = score
            melhor_termo = termo

        if score >= 95:
            break

    if melhor_score < 95:
        faltantes.append(
        {
            "documento_faltante":
                original,

            "melhor_correspondencia":
                melhor_score,

            "termo_semelhante":
                melhor_termo
        }
    )

# --------------------------------------------------
# RELATÓRIO
# --------------------------------------------------

relatorio = pd.DataFrame(faltantes)

from services.db import get_supabase_client

supabase = get_supabase_client()

for _, row in relatorio.iterrows():

    supabase.table(
        "vocabulario_pendente"
    ).insert(
        {
            "documento_faltante":
                row["documento_faltante"],

            "score":
                row["melhor_correspondencia"]
        }
    ).execute()

saida = BASE / "relatorio_lacunas.xlsx"

relatorio.to_excel(
    saida,
    index=False
)

print()
print(f"Documentos na TTD: {len(docs_ttd)}")
print(f"Termos no vocabulário: {len(docs_vocabulario)}")
print(f"Lacunas encontradas: {len(relatorio)}")

print("\nPrimeiras lacunas:")
print(relatorio.head(30))

print("\n=== DOCUMENTOS COM EXAME ===")

exame = relatorio[
    relatorio["documento_faltante"]
    .astype(str)
    .str.contains(
        "exame",
        case=False,
        na=False
    )
]

print("\n=== EXAME NA TTD ===")

exame_ttd = ttd[
    ttd["item_documental"]
    .astype(str)
    .str.contains("exame", case=False, na=False)
]

print(
    exame_ttd[
        ["item_documental"]
    ]
)

print("\n=== EXAME NO VOCABULÁRIO ===")

exame_voc = voc[
    voc["termo_padronizado"]
    .astype(str)
    .str.contains("exame", case=False, na=False)
]

print(
    exame_voc[
        ["termo_padronizado"]
    ]
)

print(exame)

print()
print("Relatório gerado:")
print(saida)