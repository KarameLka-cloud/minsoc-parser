from dotenv import dotenv_values

config = dotenv_values(".env")

# ═══════════════════════════════════════════════════════════════
#  КОНФИГУРАЦИЯ
# ═══════════════════════════════════════════════════════════════

ZIP_DIR = config.get("ZIP_DIR")
EXTRACT_DIR = config.get("EXTRACT_DIR")
PROCESSED_DIR = config.get("PROCESSED_DIR")
FILES_DIR = config.get("FILES_DIR")

EXCEL_SKIP_ROWS = int(config.get("EXCEL_SKIP_ROWS"))
FOLDER_COL_INDEX = int(config.get("FOLDER_COL_INDEX"))
FILE_COL_INDEX = int(config.get("FILE_COL_INDEX"))
