from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
REFERENCE_DIR = DATA_DIR / "reference"
RUNTIME_DIR = DATA_DIR / "runtime"

DB_PATH = RUNTIME_DIR / "sri_udesc.sqlite3"

FUNDO_PADRAO = "UDESC"

OFFICIAL_BOX_COVER_NAME = "capas_caixas_udesc.docx"
QUICK_FILL_WORKBOOK_NAME = "inventario_documental_preenchimento_rapido.xlsx"

OFFICIAL_INVENTORY_TEMPLATE_PATH = REFERENCE_DIR / "inventario_documental_modelo.xlsx"