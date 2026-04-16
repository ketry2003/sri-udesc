from __future__ import annotations

from functools import lru_cache
from pathlib import Path
import re
import unicodedata

import pandas as pd


REQUIRED_COLUMNS = [
    "natureza_documental",
    "grupo",
    "subgrupo",
    "serie",
    "subserie",
    "dossie_processo",
    "item_documental",
    "assunto",
    "codigo_classificacao",
    "grupo_codigo",
    "subgrupo_codigo",
    "serie_codigo",
    "subserie_codigo",
    "dossie_processo_codigo",
    "item_documental_codigo",
    "grupo_descricao",
    "subgrupo_descricao",
    "serie_descricao",
    "subserie_descricao",
    "dossie_processo_descricao",
    "item_documental_descricao",
    "prazo_corrente",
    "prazo_intermediario",
    "destinacao_final",
    "observacao",
    "source_priority",
]


def normalize(text):
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^\w_]", "", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def normalize_text(text):
    if text is None:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"\s+", " ", text).strip()
    return text


def natural_key(text):
    text = str(text or "")
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", text)
    ]


COL_MAP = {
    "natureza_documental": "natureza_documental",
    "natureza": "natureza_documental",
    "grupo": "grupo",
    "subgrupo": "subgrupo",
    "serie": "serie",
    "subserie": "subserie",
    "dossie_processo": "dossie_processo",
    "item_documental": "item_documental",
    "assunto": "assunto",
    "codigo_classificacao": "codigo_classificacao",
    "grupo_codigo": "grupo_codigo",
    "subgrupo_codigo": "subgrupo_codigo",
    "serie_codigo": "serie_codigo",
    "subserie_codigo": "subserie_codigo",
    "dossie_processo_codigo": "dossie_processo_codigo",
    "item_documental_codigo": "item_documental_codigo",
    "grupo_descricao": "grupo_descricao",
    "subgrupo_descricao": "subgrupo_descricao",
    "serie_descricao": "serie_descricao",
    "subserie_descricao": "subserie_descricao",
    "dossie_processo_descricao": "dossie_processo_descricao",
    "item_documental_descricao": "item_documental_descricao",
    "prazo_corrente": "prazo_corrente",
    "prazo_intermediario": "prazo_intermediario",
    "destinacao_final": "destinacao_final",
    "observacao": "observacao",
}


def _prepare_df(df, natureza_padrao, prioridade):
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.dropna(how="all").copy()

    rename = {}
    for col in df.columns:
        rename[col] = COL_MAP.get(normalize(col), normalize(col))

    df = df.rename(columns=rename)

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    for col in REQUIRED_COLUMNS:
        df[col] = df[col].fillna("").astype(str).str.strip()

    if df["natureza_documental"].eq("").all():
        df["natureza_documental"] = natureza_padrao

    if df["source_priority"].eq("").all():
        df["source_priority"] = str(prioridade)

    return df


@lru_cache(maxsize=3)
def load_ttd(tipo="todos", path=None):
    base_dir = Path(__file__).resolve().parent.parent
    excel_path = Path(path) if path else base_dir / "data" / "reference" / "ttd.xlsx"

    if not excel_path.exists():
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    df_meio = pd.read_excel(excel_path, sheet_name="ativ_meio", dtype=str)
    df_meio = _prepare_df(df_meio, "Atividade-meio", 2)

    df_fim = pd.read_excel(excel_path, sheet_name="ativ_fim", dtype=str)
    df_fim = _prepare_df(df_fim, "Atividade-fim", 1)

    if tipo == "meio":
        return df_meio

    if tipo == "fim":
        return df_fim

    return pd.concat([df_fim, df_meio], ignore_index=True)


def apply_filters(df, filters=None):
    if not filters:
        return df

    out = df.copy()

    for col, value in filters.items():
        if col not in out.columns:
            continue
        if not value:
            continue

        out = out[
            out[col].fillna("").astype(str).str.strip()
            == str(value).strip()
        ]

    return out


def get_filter_options(df, column=None, filters=None):
    df = apply_filters(df, filters)

    if not column or column not in df.columns:
        return []

    vals = df[column].fillna("").astype(str).str.strip()
    vals = vals[vals != ""].unique().tolist()

    return sorted(vals, key=natural_key)


def search_records(df, query="", filters=None, limit=30):
    df = apply_filters(df, filters)

    if query:
        q = normalize_text(query)

        campos_busca = [
            "item_documental",
            "codigo_classificacao",
            "assunto",
            "subserie",
            "dossie_processo",
            "observacao",
        ]

        mask = None

        for col in campos_busca:
            if col not in df.columns:
                continue

            col_text = df[col].fillna("").astype(str).map(normalize_text)
            atual = col_text.str.contains(q, na=False, regex=False)

            mask = atual if mask is None else (mask | atual)

        if mask is not None:
            df = df[mask]

    return df.sort_values(
        by=["source_priority", "codigo_classificacao", "item_documental"],
        key=lambda col: col.map(natural_key) if col.name == "codigo_classificacao" else col
    ).head(limit)