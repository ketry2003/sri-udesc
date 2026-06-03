from __future__ import annotations

import os
from contextlib import contextmanager

import psycopg2
from psycopg2.extras import RealDictCursor

try:
    import streamlit as st
except Exception:
    st = None


SCHEMA = """
CREATE TABLE IF NOT EXISTS inventory_items (
    id BIGSERIAL PRIMARY KEY,
    setor TEXT NOT NULL,
    tipo_documental TEXT,
    natureza_documental TEXT,
    grupo TEXT,
    subgrupo TEXT,
    serie TEXT,
    subserie TEXT,
    dossie_processo TEXT,
    item_documental TEXT,
    codigo_classificacao TEXT,
    prazo_corrente TEXT,
    prazo_intermediario TEXT,
    destinacao_final TEXT,
    datas_limite TEXT,
    quantidade INTEGER DEFAULT 1,
    caixa TEXT,
    observacoes TEXT,
    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS validation_records (
    id BIGSERIAL PRIMARY KEY,
    inventory_id BIGINT REFERENCES inventory_items(id) ON DELETE CASCADE,
    situacao_prazo TEXT,
    pode_eliminar BOOLEAN DEFAULT FALSE,
    justificativa TEXT,
    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
"""


def get_database_url() -> str:
    if st is not None:
        try:
            return st.secrets["SUPABASE_DB_URL"]
        except Exception:
            pass

    url = os.getenv("SUPABASE_DB_URL")

    if not url:
        raise RuntimeError(
            "SUPABASE_DB_URL não configurado. "
            "Configure em .streamlit/secrets.toml ou nas Secrets do Streamlit Cloud."
        )

    return url


@contextmanager
def get_conn():
    conn = psycopg2.connect(
        get_database_url(),
        cursor_factory=RealDictCursor,
    )

    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(SCHEMA)


def insert_inventory_item(payload: dict) -> int:
    fields = [
        "setor",
        "tipo_documental",
        "natureza_documental",
        "grupo",
        "subgrupo",
        "serie",
        "subserie",
        "dossie_processo",
        "item_documental",
        "codigo_classificacao",
        "prazo_corrente",
        "prazo_intermediario",
        "destinacao_final",
        "datas_limite",
        "quantidade",
        "caixa",
        "observacoes",
    ]

    values = [payload.get(field) for field in fields]
    placeholders = ", ".join(["%s"] * len(fields))

    sql = f"""
        INSERT INTO inventory_items ({", ".join(fields)})
        VALUES ({placeholders})
        RETURNING id
    """

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, values)
            row = cur.fetchone()
            return int(row["id"])


def list_inventory_items() -> list[dict]:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM inventory_items
                ORDER BY criado_em DESC, id DESC
                """
            )
            return cur.fetchall()


def list_setores_inventory() -> list[str]:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT TRIM(COALESCE(setor, '')) AS setor
                FROM inventory_items
                WHERE TRIM(COALESCE(setor, '')) <> ''
                ORDER BY setor
                """
            )
            rows = cur.fetchall()

    return [row["setor"] for row in rows]


def list_inventory_items_by_setor(setor: str) -> list[dict]:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT *
                FROM inventory_items
                WHERE TRIM(COALESCE(setor, '')) = TRIM(%s)
                ORDER BY criado_em DESC, id DESC
                """,
                (setor,),
            )
            return cur.fetchall()


def delete_inventory_item(item_id: int) -> None:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM validation_records WHERE inventory_id = %s",
                (item_id,),
            )
            cur.execute(
                "DELETE FROM inventory_items WHERE id = %s",
                (item_id,),
            )


def delete_inventory_items(item_ids: list[int]) -> int:
    if not item_ids:
        return 0

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM validation_records WHERE inventory_id = ANY(%s)",
                (item_ids,),
            )
            cur.execute(
                "DELETE FROM inventory_items WHERE id = ANY(%s)",
                (item_ids,),
            )

    return len(item_ids)


def delete_inventory_items_by_setor(setor: str) -> int:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id
                FROM inventory_items
                WHERE TRIM(COALESCE(setor, '')) = TRIM(%s)
                """,
                (setor,),
            )

            rows = cur.fetchall()
            item_ids = [row["id"] for row in rows]

            if not item_ids:
                return 0

            cur.execute(
                "DELETE FROM validation_records WHERE inventory_id = ANY(%s)",
                (item_ids,),
            )
            cur.execute(
                "DELETE FROM inventory_items WHERE id = ANY(%s)",
                (item_ids,),
            )

    return len(item_ids)


def replace_inventory_from_dataframe(df) -> int:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM validation_records")
            cur.execute("DELETE FROM inventory_items")

            for _, row in df.iterrows():
                cur.execute(
                    """
                    INSERT INTO inventory_items (
                        setor,
                        tipo_documental,
                        natureza_documental,
                        grupo,
                        subgrupo,
                        serie,
                        subserie,
                        dossie_processo,
                        item_documental,
                        codigo_classificacao,
                        prazo_corrente,
                        prazo_intermediario,
                        destinacao_final,
                        datas_limite,
                        quantidade,
                        caixa,
                        observacoes
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        row.get("setor", ""),
                        row.get("tipo_documental", ""),
                        row.get("natureza_documental", ""),
                        row.get("grupo", ""),
                        row.get("subgrupo", ""),
                        row.get("serie", ""),
                        row.get("subserie", ""),
                        row.get("dossie_processo", ""),
                        row.get("item_documental", ""),
                        row.get("codigo_classificacao", ""),
                        row.get("prazo_corrente", ""),
                        row.get("prazo_intermediario", ""),
                        row.get("destinacao_final", ""),
                        row.get("datas_limite", ""),
                        int(row.get("quantidade", 1) or 1),
                        row.get("caixa", ""),
                        row.get("observacoes", ""),
                    ),
                )

    return len(df)


def replace_inventory_from_dataframe_by_setor(df, setor: str) -> int:
    setor = str(setor or "").strip()

    if not setor:
        raise ValueError("Setor não informado.")

    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id
                FROM inventory_items
                WHERE TRIM(COALESCE(setor, '')) = TRIM(%s)
                """,
                (setor,),
            )

            rows = cur.fetchall()
            item_ids = [row["id"] for row in rows]

            if item_ids:
                cur.execute(
                    "DELETE FROM validation_records WHERE inventory_id = ANY(%s)",
                    (item_ids,),
                )
                cur.execute(
                    "DELETE FROM inventory_items WHERE id = ANY(%s)",
                    (item_ids,),
                )

            for _, row in df.iterrows():
                cur.execute(
                    """
                    INSERT INTO inventory_items (
                        setor,
                        tipo_documental,
                        natureza_documental,
                        grupo,
                        subgrupo,
                        serie,
                        subserie,
                        dossie_processo,
                        item_documental,
                        codigo_classificacao,
                        prazo_corrente,
                        prazo_intermediario,
                        destinacao_final,
                        datas_limite,
                        quantidade,
                        caixa,
                        observacoes
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        setor,
                        row.get("tipo_documental", ""),
                        row.get("natureza_documental", ""),
                        row.get("grupo", ""),
                        row.get("subgrupo", ""),
                        row.get("serie", ""),
                        row.get("subserie", ""),
                        row.get("dossie_processo", ""),
                        row.get("item_documental", ""),
                        row.get("codigo_classificacao", ""),
                        row.get("prazo_corrente", ""),
                        row.get("prazo_intermediario", ""),
                        row.get("destinacao_final", ""),
                        row.get("datas_limite", ""),
                        int(row.get("quantidade", 1) or 1),
                        row.get("caixa", ""),
                        row.get("observacoes", ""),
                    ),
                )

    return len(df)


def save_validation_record(
    inventory_id: int,
    situacao_prazo: str,
    pode_eliminar: bool,
    justificativa: str,
) -> None:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO validation_records (
                    inventory_id,
                    situacao_prazo,
                    pode_eliminar,
                    justificativa
                ) VALUES (%s, %s, %s, %s)
                """,
                (
                    inventory_id,
                    situacao_prazo,
                    bool(pode_eliminar),
                    justificativa,
                ),
            )


def list_elimination_candidates() -> list[dict]:
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT i.*, v.situacao_prazo, v.pode_eliminar, v.justificativa
                FROM inventory_items i
                INNER JOIN validation_records v ON v.inventory_id = i.id
                WHERE v.pode_eliminar = TRUE
                  AND LOWER(COALESCE(i.destinacao_final, '')) NOT LIKE '%permanente%'
                ORDER BY i.caixa, i.tipo_documental
                """
            )
            return cur.fetchall()


def update_inventory_item(item_id, payload):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE inventory_items
                SET
                    datas_limite = %s,
                    quantidade = %s,
                    caixa = %s,
                    observacoes = %s
                WHERE id = %s
                """,
                (
                    payload.get("datas_limite", ""),
                    payload.get("quantidade", 1),
                    payload.get("caixa", ""),
                    payload.get("observacoes", ""),
                    item_id,
                ),
            )

            total = cur.rowcount

    return total