
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
REFERENCE_DIR = DATA_DIR / "reference"
CDOC_REFERENCE_DIR = REFERENCE_DIR / "cdoc"
RUNTIME_DIR = DATA_DIR / "runtime"
EXPORTS_DIR = BASE_DIR / "exports"

DB_PATH = RUNTIME_DIR / "sri_udesc.sqlite3"

TTD_PATH = REFERENCE_DIR / "TTD_UDESC_SEA_filtravel.xlsx"
OFFICIAL_INVENTORY_TEMPLATE_PATH = CDOC_REFERENCE_DIR / "Anexo_IV_IN_008_2024_Inventario_AtvFim.xlsx"
OFFICIAL_BOX_LABEL_TEMPLATE_PATH = CDOC_REFERENCE_DIR / "Anexo_I_IN_008_2024_Etiqueta_da_Caixa.docx"
ATIVIDADE_FIM_TEMPLATE_PATH = CDOC_REFERENCE_DIR / "AnexoII_IN0082023_Tabela_de_Classificacao_AtvFim.xlsx"
ATIVIDADE_MEIO_TEMPLATE_PATH = CDOC_REFERENCE_DIR / "Anexo_V_IN_008_2024_Tabela_de_Classificacao_AtvMeio.xlsx"

FUNDO_PADRAO = "UDESC / CCT"
QUICK_FILL_WORKBOOK_NAME = "inventario_documental_preenchimento_assistido.xlsx"
OFFICIAL_BOX_COVER_NAME = "capas_caixa_cdoc.docx"

RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
