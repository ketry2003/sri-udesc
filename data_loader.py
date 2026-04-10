
from __future__ import annotations

from functools import lru_cache
from pathlib import Path
import re
import unicodedata

import pandas as pd

from config import TTD_PATH

COL_MAP = {
    "PDF": "fonte_pdf",
    "Página": "pagina",
    "Grupo": "grupo",
    "Subgrupo": "subgrupo",
    "Função": "serie",
    "Subfunção": "subserie",
    "Atividade": "dossie_processo",
    "Documento": "item_documental",
    "Prazo corrente": "prazo_corrente",
    "Prazo intermediária": "prazo_intermediario",
    "Observação": "observacao",
    "Destinação resumida": "destinacao_final",
}

TEXT_FIELDS = [
    "natureza_documental", "grupo", "subgrupo", "serie", "subserie",
    "dossie_processo", "item_documental", "observacao", "destinacao_final",
    "codigo_classificacao",
]

def normalize_text(text: str) -> str:
    text = str(text or "").strip()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    return text

def split_code_description(value: str) -> tuple[str, str]:
    text = str(value or "").strip()
    match = re.match(r"^\s*([\d\.\-]+)\s*[-–]\s*(.+)$", text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return "", text

def safe_int_from_text(value: str) -> int:
    text = normalize_text(value)
    match = re.search(r"(\d+)", text)
    return int(match.group(1)) if match else 0

def build_classification_code(row: pd.Series) -> str:
    parts = []
    for field in ["grupo_codigo", "subgrupo_codigo", "serie_codigo", "subserie_codigo", "dossie_processo_codigo", "item_documental_codigo"]:
        val = str(row.get(field, "") or "").strip().strip(".")
        if val:
            parts.append(val)
    return ".".join(parts)

@lru_cache(maxsize=1)
def load_ttd() -> pd.DataFrame:
    if not Path(TTD_PATH).exists():
        raise FileNotFoundError(f"Arquivo TTD não encontrado em {TTD_PATH}")

    df = pd.read_excel(TTD_PATH, sheet_name="Todos")
    df = df.rename(columns=COL_MAP)
    df["natureza_documental"] = df["fonte_pdf"].map(
        {
            "Atividade-fim UDESC": "Atividade-fim",
            "Atividade-meio SEA": "Atividade-meio",
        }
    ).fillna(df["fonte_pdf"])

    df["marcado_eliminacao"] = df.get("Marcado eliminação?", "").astype(str).str.lower().eq("sim")
    df["marcado_guarda_permanente"] = df.get("Marcado guarda permanente?", "").astype(str).str.lower().eq("sim")

    for field in ["grupo", "subgrupo", "serie", "subserie", "dossie_processo", "item_documental"]:
        codes, descs = zip(*df[field].fillna("").map(split_code_description))
        df[f"{field}_codigo"] = list(codes)
        df[f"{field}_descricao"] = list(descs)

    df["codigo_classificacao"] = df.apply(build_classification_code, axis=1)
    df["source_priority"] = df["natureza_documental"].map({"Atividade-fim": 0, "Atividade-meio": 1}).fillna(9)
    df["prazo_corrente_anos"] = df["prazo_corrente"].map(safe_int_from_text)
    df["prazo_intermediario_anos"] = df["prazo_intermediario"].map(safe_int_from_text)
    df["total_prazo_anos"] = df["prazo_corrente_anos"] + df["prazo_intermediario_anos"]

    df["texto_busca"] = (
        df[TEXT_FIELDS].fillna("").astype(str).agg(" | ".join, axis=1).map(normalize_text)
    )

    return df.fillna("")

def get_filter_options(df: pd.DataFrame, field: str, filters: dict[str, str]) -> list[str]:
    filtered = apply_filters(df, filters)
    values = sorted(v for v in filtered[field].dropna().astype(str).unique().tolist() if v)
    return values

def apply_filters(df: pd.DataFrame, filters: dict[str, str]) -> pd.DataFrame:
    filtered = df.copy()
    for key, value in filters.items():
        if value:
            filtered = filtered[filtered[key] == value]
    return filtered
