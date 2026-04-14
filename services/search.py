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
    "eliminacao",
    "guarda_permanente",
    "marcado_eliminacao",
    "marcado_guarda_permanente",
    "destinacao_resumida",
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

    # seu novo modelo Excel
    "subfuncao": "subserie",
    "atividade": "dossie_processo",
    "documento": "item_documental",
    "prazo_intermediaria": "prazo_intermediario",
    "destinacao_resumida": "destinacao_resumida",

    "eliminacao": "eliminacao",
    "guarda_permanente": "guarda_permanente",
    "marcado_eliminacao": "marcado_eliminacao",
    "marcado_guarda_permanente": "marcado_guarda_permanente",
}


def _prepare_df(df, natureza_padrao, prioridade):
    if df is None or df.empty:
        return pd.DataFrame()

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

    # natureza
    vazia = df["natureza_documental"] == ""
    df.loc[vazia, "natureza_documental"] = natureza_padrao

    # prioridade
    vazia = df["source_priority"] == ""
    df.loc[vazia, "source_priority"] = str(prioridade)

    # descrições
    if df["item_documental_descricao"].eq("").all():
        df["item_documental_descricao"] = df["item_documental"]

    # destino
    vazia = df["destinacao_final"] == ""

    gp = df["guarda_permanente"]
    df.loc[vazia & gp.ne(""), "destinacao_final"] = "Guarda permanente"

    vazia = df["destinacao_final"] == ""

    el = df["eliminacao"]
    df.loc[vazia & el.ne(""), "destinacao_final"] = "Eliminação"

    vazia = df["destinacao_final"] == ""

    dr = df["destinacao_resumida"]
    df.loc[vazia & dr.ne(""), "destinacao_final"] = dr

    return df


@lru_cache(maxsize=3)
def load_ttd(tipo="todos", path=None):
    base_dir = Path(__file__).resolve().parent.parent
    excel_path = Path(path) if path else base_dir / "data" / "reference" / "ttd.xlsx"

    if not excel_path.exists():
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    if tipo not in {"meio", "fim", "todos"}:
        raise ValueError("Tipo deve ser 'meio', 'fim' ou 'todos'")

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

        out = out[out[col].fillna("").astype(str).str.strip() == str(value).strip()]

    return out


def get_filter_options(df, column=None, filters=None):
    df = apply_filters(df, filters)

    if column:
        if column not in df.columns:
            return []

        vals = df[column].fillna("").astype(str).str.strip()
        vals = vals[vals != ""]
        return sorted(vals.unique())

    return {}


def search_records(df, query="", filters=None, limit=30):
    df = apply_filters(df, filters)

    if query:
        q = normalize_text(query)

        mask = False
        for col in [
            "item_documental",
            "codigo_classificacao",
            "assunto",
            "observacao",
        ]:
            if col in df.columns:
                col_text = df[col].fillna("").astype(str).map(normalize_text)
                col_mask = col_text.str.contains(q, na=False, regex=False)
                mask = col_mask if isinstance(mask, bool) else (mask | col_mask)

        if not isinstance(mask, bool):
            df = df[mask]

    return df.sort_values(
        ["source_priority", "codigo_classificacao", "item_documental"]
    ).head(limit)