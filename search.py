
from __future__ import annotations

from rapidfuzz import fuzz
import pandas as pd

from services.data_loader import normalize_text, apply_filters

SEARCH_FIELDS = [
    "codigo_classificacao",
    "item_documental",
    "item_documental_descricao",
    "dossie_processo_descricao",
    "subserie_descricao",
    "serie_descricao",
    "observacao",
]

def search_records(df: pd.DataFrame, query: str, filters: dict[str, str], limit: int = 50) -> pd.DataFrame:
    filtered = apply_filters(df, filters)
    if not query:
        return filtered.sort_values(["source_priority", "codigo_classificacao", "item_documental"]).head(limit)

    q = normalize_text(query)

    def score_row(row) -> int:
        haystack = " | ".join(str(row.get(col, "")) for col in SEARCH_FIELDS)
        haystack = normalize_text(haystack)
        partial = fuzz.partial_ratio(q, haystack)
        token = fuzz.token_set_ratio(q, haystack)
        exact_boost = 30 if q in haystack else 0
        code_boost = 45 if q and q in normalize_text(str(row.get("codigo_classificacao", ""))) else 0
        starts_boost = 20 if normalize_text(str(row.get("codigo_classificacao", ""))).startswith(q) else 0
        return max(partial, token) + exact_boost + code_boost + starts_boost

    scored = filtered.copy()
    scored["_score"] = scored.apply(score_row, axis=1)
    scored = scored[scored["_score"] >= 40].sort_values(
        ["_score", "source_priority", "codigo_classificacao", "serie"],
        ascending=[False, True, True, True],
    )
    return scored.head(limit)
