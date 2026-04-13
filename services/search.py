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


COL_MAP = {
    "natureza_documental": "natureza_documental",
    "natureza": "natureza_documental",
    "grupo": "grupo",
    "subgrupo": "subgrupo",
    "serie": "serie",
    "subserie": "subserie",
    "dossie_processo": "dossie_processo",
    "item_documental": "item_documental",
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

    "subfuncao": "subserie",
    "atividade": "dossie_processo",
    "documento": "item_documental",
    "prazo_intermediaria": "prazo_intermediario",
    "destinacao_resumida": "destinacao_final",

    "eliminacao": "eliminacao",
    "guarda_permanente": "guarda_permanente",
    "marcado_eliminacao": "marcado_eliminacao",
    "marcado_guarda_permanente": "marcado_guarda_permanente",
}


@lru_cache(maxsize=1)
def load_ttd(path=None):
    base_dir = Path(__file__).resolve().parent.parent
    excel_path = Path(path) if path else base_dir / "data" / "reference" / "TTD.xlsx"

    if not excel_path.exists():
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    df = pd.read_excel(
        excel_path,
        sheet_name="Todos",
        engine="openpyxl",
        dtype=str,
    )

    if df is None or df.empty:
        raise ValueError("A planilha 'Todos' está vazia")

    df = df.dropna(how="all").copy()

    rename = {}
    for col in df.columns:
        rename[col] = COL_MAP.get(normalize(col), normalize(col))

    df = df.rename(columns=rename)

    for col in df.columns:
        df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    for col in REQUIRED_COLUMNS:
        df[col] = df[col].fillna("").astype(str).str.strip()

    if df["item_documental_descricao"].eq("").all():
        df["item_documental_descricao"] = df["item_documental"]

    if df["dossie_processo_descricao"].eq("").all():
        df["dossie_processo_descricao"] = df["dossie_processo"]

    if df["subserie_descricao"].eq("").all():
        df["subserie_descricao"] = df["subserie"]

    if df["serie_descricao"].eq("").all():
        df["serie_descricao"] = df["serie"]

    if df["source_priority"].eq("").all():
        df["source_priority"] = "1"

    vazia = df["destinacao_final"].eq("")

    if "guarda_permanente" in df.columns:
        gp = df["guarda_permanente"].fillna("").astype(str).str.strip()
        df.loc[vazia & gp.ne(""), "destinacao_final"] = "Guarda permanente"

    vazia = df["destinacao_final"].eq("")

    if "eliminacao" in df.columns:
        el = df["eliminacao"].fillna("").astype(str).str.strip()
        df.loc[vazia & el.ne(""), "destinacao_final"] = "Eliminação"

    return df


def apply_filters(df, filters=None):
    if not filters:
        return df

    filtered_df = df.copy()

    for col, value in filters.items():
        if col not in filtered_df.columns:
            continue
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue

        filtered_df = filtered_df[
            filtered_df[col].fillna("").astype(str).str.strip() == str(value).strip()
        ]

    return filtered_df


def get_filter_options(df, column=None, filters=None):
    filtered_df = apply_filters(df, filters)

    if column is not None:
        if column not in filtered_df.columns:
            return []

        values = filtered_df[column].fillna("").astype(str).str.strip()
        values = values[values != ""]
        return sorted(values.unique().tolist())

    filter_columns = [
        "natureza_documental",
        "grupo",
        "subgrupo",
        "serie",
        "subserie",
        "dossie_processo",
        "item_documental",
        "prazo_corrente",
        "prazo_intermediario",
        "destinacao_final",
    ]

    options = {}
    for col in filter_columns:
        if col in filtered_df.columns:
            values = filtered_df[col].fillna("").astype(str).str.strip()
            values = values[values != ""]
            options[col] = sorted(values.unique().tolist())
        else:
            options[col] = []

    return options

def search_records(df, query="", filters=None, limit=30):
    filtered_df = apply_filters(df, filters)

    if query and str(query).strip():
        q = normalize_text(query)
        search_cols = [
            "item_documental",
            "item_documental_descricao",
            "dossie_processo",
            "dossie_processo_descricao",
            "subserie",
            "subserie_descricao",
            "serie",
            "serie_descricao",
            "subgrupo",
            "subgrupo_descricao",
            "grupo",
            "grupo_descricao",
            "codigo_classificacao",
            "grupo_codigo",
            "subgrupo_codigo",
            "serie_codigo",
            "subserie_codigo",
            "dossie_processo_codigo",
            "item_documental_codigo",
            "observacao",
        ]

        mask = False
        for col in search_cols:
            if col in filtered_df.columns:
                col_text = filtered_df[col].fillna("").astype(str).map(normalize_text)
                col_mask = col_text.str.contains(q, na=False, regex=False)
                mask = col_mask if isinstance(mask, bool) else (mask | col_mask)

        if not isinstance(mask, bool):
            filtered_df = filtered_df[mask]

    if "source_priority" in filtered_df.columns:
        filtered_df = filtered_df.sort_values(
            by=["source_priority", "codigo_classificacao", "item_documental"],
            ascending=[True, True, True],
        )
    else:
        filtered_df = filtered_df.sort_values(
            by=["codigo_classificacao", "item_documental"],
            ascending=[True, True],
        )

    return filtered_df.head(limit).copy()
