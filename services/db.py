import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable

from config import DB_PATH

SCHEMA = """
CREATE TABLE IF NOT EXISTS inventory_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    setor TEXT NOT NULL,
    tipo_documental TEXT NOT NULL,
    natureza_documental TEXT,
    grupo TEXT,
    subgrupo TEXT,
    serie TEXT,
    subserie TEXT,
    dossie_processo TEXT,
    item_documental TEXT,
    prazo_corrente TEXT,
    prazo_intermediario TEXT,
    destinacao_final TEXT,
    datas_limite TEXT,
    quantidade INTEGER DEFAULT 1,
    caixa TEXT,
    observacoes TEXT,
    criado_em TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS validation_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    inventory_id INTEGER,
    situacao_prazo TEXT NOT NULL,
    pode_eliminar INTEGER NOT NULL DEFAULT 0,
    justificativa TEXT,
    criado_em TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (inventory_id) REFERENCES inventory_items(id)
);
"""


@contextmanager
def get_conn():
    Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with get_conn() as conn:
        conn.executescript(SCHEMA)


def insert_inventory_item(payload: dict) -> int:
    fields = [
        "setor", "tipo_documental", "natureza_documental", "grupo", "subgrupo", "serie",
        "subserie", "dossie_processo", "item_documental", "prazo_corrente",
        "prazo_intermediario", "destinacao_final", "datas_limite", "quantidade",
        "caixa", "observacoes",
    ]
    values = [payload.get(field) for field in fields]
    placeholders = ", ".join(["?"] * len(fields))
    sql = f"INSERT INTO inventory_items ({', '.join(fields)}) VALUES ({placeholders})"
    with get_conn() as conn:
        cur = conn.execute(sql, values)
        return int(cur.lastrowid)


def list_inventory_items() -> list[sqlite3.Row]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM inventory_items ORDER BY datetime(criado_em) DESC, id DESC"
        ).fetchall()
    return rows


def delete_inventory_item(item_id: int) -> None:
    with get_conn() as conn:
        conn.execute("DELETE FROM inventory_items WHERE id = ?", (item_id,))
        conn.execute("DELETE FROM validation_records WHERE inventory_id = ?", (item_id,))


def replace_inventory_from_dataframe(df) -> int:
    with get_conn() as conn:
        conn.execute("DELETE FROM validation_records")
        conn.execute("DELETE FROM inventory_items")
        for _, row in df.iterrows():
            conn.execute(
                """
                INSERT INTO inventory_items (
                    setor, tipo_documental, natureza_documental, grupo, subgrupo, serie,
                    subserie, dossie_processo, item_documental, prazo_corrente,
                    prazo_intermediario, destinacao_final, datas_limite, quantidade,
                    caixa, observacoes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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


def save_validation_record(inventory_id: int, situacao_prazo: str, pode_eliminar: bool, justificativa: str) -> None:
    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO validation_records (inventory_id, situacao_prazo, pode_eliminar, justificativa)
            VALUES (?, ?, ?, ?)
            """,
            (inventory_id, situacao_prazo, int(pode_eliminar), justificativa),
        )


def list_elimination_candidates() -> list[sqlite3.Row]:
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT i.*, v.situacao_prazo, v.pode_eliminar, v.justificativa
            FROM inventory_items i
            INNER JOIN validation_records v ON v.inventory_id = i.id
            WHERE v.pode_eliminar = 1 AND LOWER(COALESCE(i.destinacao_final, '')) NOT LIKE '%permanente%'
            ORDER BY i.caixa, i.tipo_documental
            """
        ).fetchall()
    return rows
