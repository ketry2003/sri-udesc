from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
REFERENCE_DIR = DATA_DIR / "reference"
RUNTIME_DIR = DATA_DIR / "runtime"
EXPORTS_DIR = BASE_DIR / "exports"

DB_PATH = RUNTIME_DIR / "sri_udesc.sqlite3"

TTD_PATH = REFERENCE_DIR / "ttd.xlsx"
OFFICIAL_INVENTORY_TEMPLATE_PATH = REFERENCE_DIR / "inventario.xlsx"
OFFICIAL_BOX_LABEL_TEMPLATE_PATH = REFERENCE_DIR / "etiqueta_caixas.docx"
ATIVIDADE_FIM_TEMPLATE_PATH = REFERENCE_DIR / "ttd.xlsx"
ATIVIDADE_MEIO_TEMPLATE_PATH = REFERENCE_DIR / "ttd.xlsx"

FUNDO_PADRAO = "UDESC / CCT"
QUICK_FILL_WORKBOOK_NAME = "inventario_documental_preenchimento_assistido.xlsx"
OFFICIAL_BOX_COVER_NAME = "capas_caixa_cdoc.docx"

RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
EXPORTS_DIR.mkdir(parents=True, exist_ok=True)