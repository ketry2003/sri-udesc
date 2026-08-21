import pandas as pd
from pathlib import Path
import unicodedata
import re

BASE = Path(__file__).resolve().parent / "data" / "reference"

ARQ_PLAN = BASE / "planilha_atualizada.xlsx"
ARQ_TTD = BASE / "ttd_atual.xlsx"


def normalizar(texto):

    if pd.isna(texto):
        return ""

    texto = str(texto)

    # remove prefixos do tipo:
    # 001 - Documento
    # 025 - Documento
    texto = re.sub(
        r"^\d+\s*-\s*",
        "",
        texto
    )

    texto = texto.lower().strip()

    texto = "".join(
        c
        for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )

    return texto


# ==================================
# LEITURA
# ==================================

plan = pd.read_excel(
    ARQ_PLAN,
    sheet_name="base_adaptada",
    engine="openpyxl"
)

plan_original = len(plan)

ttd = pd.read_excel(
    ARQ_TTD,
    engine="openpyxl"
)

print(f"Planilha original: {len(plan)}")
print(f"TTD nova: {len(ttd)}")


# ==================================
# CHAVES
# ==================================

plan["chave"] = (
    plan["termo_padronizado"]
    .fillna("")
    .apply(normalizar)
)

ttd["chave"] = (
    ttd["item_documental"]
    .fillna("")
    .apply(normalizar)
)

indice = {
    chave: idx
    for idx, chave in plan["chave"].items()
}

novos = []
atualizados = 0

nao_encontrados = []

# ==================================
# MERGE
# ==================================

for _, linha in ttd.iterrows():

    chave = linha["chave"]

    # ----------------------
    # DESTINAÇÃO
    # ----------------------

    destinacao = ""

    if pd.notna(linha["guarda_permanente"]):
        destinacao = "Guarda permanente"

    elif pd.notna(linha["eliminacao"]):
        destinacao = "Eliminação"

    # ----------------------
    # JÁ EXISTE
    # ----------------------

    if chave in indice:

        idx = indice[chave]

        if "subfuncao" in plan.columns:
            plan.at[idx, "subfuncao"] = linha["subfuncao"]

        if "atividade" in plan.columns:
            plan.at[idx, "atividade"] = linha["atividade"]

        plan.at[idx, "prazo_corrente"] = (
            linha["prazo_corrente"]
        )

        plan.at[idx, "prazo_intermediario"] = (
            linha["prazo_intermediario"]
        )

        plan.at[idx, "destinacao"] = destinacao

        plan.at[idx, "observacao"] = (
            linha["observacao"]
        )

        # Preencher assunto apenas se vazio

        if (
            pd.isna(plan.at[idx, "assunto_tecnico"])
            or str(
                plan.at[idx, "assunto_tecnico"]
            ).strip() == ""
        ):
            plan.at[idx, "assunto_tecnico"] = (
                linha["assunto"]
            )

        # Preencher código apenas se vazio

        if (
            pd.isna(plan.at[idx, "codigo_classificacao"])
            or str(
                plan.at[idx, "codigo_classificacao"]
            ).strip() == ""
        ):
            plan.at[idx, "codigo_classificacao"] = (
                linha["codigo_classificacao"]
            )

        atualizados += 1

    # ----------------------
    # DOCUMENTO NOVO
    # ----------------------

    else:

        novo = {
            col: None
            for col in plan.columns
        }

        novo["tipo_atividade"] = "fim"

        novo["termo_padronizado"] = (
            linha["item_documental"]
        )

        novo["subfuncao"] = linha["subfuncao"]

        novo["atividade"] = linha["atividade"]

        novo["assunto_tecnico"] = (
            linha["assunto"]
        )

        novo["codigo_classificacao"] = (
            linha["codigo_classificacao"]
        )

        novo["prazo_corrente"] = (
            linha["prazo_corrente"]
        )

        novo["prazo_intermediario"] = (
            linha["prazo_intermediario"]
        )

        novo["destinacao"] = destinacao

        novo["observacao"] = (
            linha["observacao"]
        )

        novos.append(novo)

        nao_encontrados.append(
            linha["item_documental"]
        )


# ==================================
# INSERIR NOVOS
# ==================================

if novos:

    plan = pd.concat(
        [
            plan,
            pd.DataFrame(novos)
        ],
        ignore_index=True
    )

# ==================================
# LIMPEZA
# ==================================

if "chave" in plan.columns:
    plan.drop(
        columns=["chave"],
        inplace=True
    )

# ==================================
# SALVAR BASE ATUALIZADA
# ==================================

saida = (
    BASE
    / "planilha_atualizada_v2.xlsx"
)

plan.to_excel(
    saida,
    index=False,
    engine="openpyxl"
)

# ==================================
# RELATÓRIO NOVOS
# ==================================

pd.DataFrame(
    {
        "item_documental": nao_encontrados
    }
).to_excel(
    BASE / "novos_documentos.xlsx",
    index=False
)

# ==================================
# RESUMO
# ==================================

print("\n========================")
print("MERGE FINALIZADO")
print("========================")

print(
    f"Registros originais: {plan_original}"
)

print(
    f"Registros TTD: {len(ttd)}"
)

print(
    f"Atualizados: {atualizados}"
)

print(
    f"Novos documentos: {len(novos)}"
)

print(
    f"Total final: {len(plan)}"
)

print(
    f"Arquivo: {saida}"
)

print(
    "Relatório: novos_documentos.xlsx"
)