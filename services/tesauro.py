from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz
import unicodedata

from services.db import carregar_vocabulario


def caminho_vocabulario():
    return (
        Path(__file__).resolve().parent.parent
        / "data"
        / "reference"
        / "planilha_atualizada.xlsx"
    )


def normalizar_colunas(df):
    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("/", "_")
        .str.replace("-", "_")
    )

    return df


def remover_acentos(texto):
    return "".join(
        caractere
        for caractere in unicodedata.normalize(
            "NFD",
            str(texto)
        )
        if unicodedata.category(caractere) != "Mn"
    )


def normalizar_texto(texto):
    texto = (
        str(texto)
        .lower()
        .strip()
        .replace("\xa0", " ")
    )

    texto = remover_acentos(texto)

    return texto


def carregar_tesauro(tipo):
    arquivo = caminho_vocabulario()

    if not arquivo.exists():
        return pd.DataFrame()

    df = carregar_vocabulario(tipo)

    if df.empty:
        return df

    df = normalizar_colunas(df)

    coluna_tipo = None

    for possivel in [
        "tipo",
        "tipo_atividade",
        "tipo_de_atividade",
        "atividade",
    ]:
        if possivel in df.columns:
            coluna_tipo = possivel
            break

    if coluna_tipo:

        df[coluna_tipo] = (
            df[coluna_tipo]
            .astype(str)
            .str.lower()
            .str.strip()
            .str.replace("atividade-", "", regex=False)
            .str.replace("atividade_", "", regex=False)
            .str.replace("atividade ", "", regex=False)
            .str.replace("\xa0", "", regex=False)
        )

        if tipo in ["fim", "meio"]:
            df = df[
                df[coluna_tipo]
                == tipo
            ]

    colunas_busca = [
        "item_documental",
        "tipo_documental",
        "termo_preferido_oficial",
        "termos_populares_sugeridos",
        "pergunta_guia_usuario",
        "assunto_tecnico",
        "funcao",
        "subfuncao",
        "atividade",
        "codigo_classificacao",
        "observacao",
        "texto_busca_sistema",
    ]

    colunas_existentes = [
        c
        for c in colunas_busca
        if c in df.columns
    ]

    if colunas_existentes:

        df["texto_busca"] = (
            df[colunas_existentes]
            .fillna("")
            .astype(str)
            .agg(" ".join, axis=1)
            .apply(normalizar_texto)
        )

    else:

        df["texto_busca"] = ""

    return df


def buscar_tesauro(
    texto,
    tipo,
    limite=20,
    corte=55,
):
    tesauro = carregar_tesauro(tipo)

    if tesauro.empty:
        return pd.DataFrame()

    termo = normalizar_texto(texto)

    if not termo:
        return pd.DataFrame()

    tesauro = tesauro.copy()

    tesauro["score"] = tesauro["texto_busca"].apply(
        lambda texto_base: fuzz.WRatio(termo, texto_base)
    )

    resultado = (
        tesauro[tesauro["score"] >= corte]
        .sort_values("score", ascending=False)
        .head(limite)
    )

    return resultado
