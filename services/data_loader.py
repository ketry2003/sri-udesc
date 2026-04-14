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
    "source_priority": "source_priority",

    # mapeamentos da sua planilha padronizada
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


def _prepare_dataframe(df: pd.DataFrame, natureza_padrao: str, prioridade: int) -> pd.DataFrame:
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

    # Natureza documental
    if df["natureza_documental"].eq("").all():
        df["natureza_documental"] = natureza_padrao
    else:
        vazias = df["natureza_documental"].eq("")
        df.loc[vazias, "natureza_documental"] = natureza_padrao

    # Prioridade da fonte
    if df["source_priority"].eq("").all():
        df["source_priority"] = str(prioridade)
    else:
        vazias = df["source_priority"].eq("")
        df.loc[vazias, "source_priority"] = str(prioridade)

    # Descrições auxiliares
    if df["item_documental_descricao"].eq("").all():
        df["item_documental_descricao"] = df["item_documental"]

    if df["dossie_processo_descricao"].eq("").all():
        df["dossie_processo_descricao"] = df["dossie_processo"]

    if df["subserie_descricao"].eq("").all():
        df["subserie_descricao"] = df["subserie"]

    if df["serie_descricao"].eq("").all():
        df["serie_descricao"] = df["serie"]

    # Destinação final a partir das colunas novas
    vazia = df["destinacao_final"].eq("")

    if "guarda_permanente" in df.columns:
        gp = df["guarda_permanente"].fillna("").astype(str).str.strip()
        df.loc[vazia & gp.ne(""), "destinacao_final"] = "Guarda permanente"

    vazia = df["destinacao_final"].eq("")

    if "eliminacao" in df.columns:
        el = df["eliminacao"].fillna("").astype(str).str.strip()
        df.loc[vazia & el.ne(""), "destinacao_final"] = "Eliminação"

    vazia = df["destinacao_final"].eq("")

    if "destinacao_resumida" in df.columns:
        dr = df["destinacao_resumida"].fillna("").astype(str).str.strip()
        df.loc[vazia & dr.ne(""), "destinacao_final"] = dr

    return df


@lru_cache(maxsize=3)
def load_ttd(tipo="todos", path=None):
    base_dir = Path(__file__).resolve().parent.parent
    arquivo_ttd = base_dir / "data" / "reference" / "ttd.xlsx"
    excel_path = Path(path) if path else arquivo_ttd

    if not excel_path.exists():
        raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")

    mapa_abas = {
        "meio": "ativ_meio",  # SEA
        "fim": "ativ_fim",    # UDESC
    }

    if tipo not in {"meio", "fim", "todos"}:
        raise ValueError("Tipo deve ser 'meio', 'fim' ou 'todos'")

    df_meio = pd.read_excel(
        excel_path,
        sheet_name=mapa_abas["meio"],
        engine="openpyxl",
        dtype=str,
    )
    df_meio = _prepare_dataframe(df_meio, "Atividade-meio", 2)

    df_fim = pd.read_excel(
        excel_path,
        sheet_name=mapa_abas["fim"],
        engine="openpyxl",
        dtype=str,
    )
    df_fim = _prepare_dataframe(df_fim, "Atividade-fim", 1)

    if tipo == "meio":
        if df_meio.empty:
            raise ValueError("A aba 'ativ_meio' está vazia")
        return df_meio

    if tipo == "fim":
        if df_fim.empty:
            raise ValueError("A aba 'ativ_fim' está vazia")
        return df_fim

    # tipo == "todos"
    if df_meio.empty and df_fim.empty:
        raise ValueError("As abas 'ativ_meio' e 'ativ_fim' estão vazias")

    return pd.concat([df_fim, df_meio], ignore_index=True)


def load_ttd_completo(path=None):
    return load_ttd("todos", path)


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
            filtered_df[col].fillna("").astype(str).str.strip()
            == str(value).strip()
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
        "assunto",
        "codigo_classificacao",
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