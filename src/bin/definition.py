import os
import platform
from datetime import datetime
from pathlib import Path

# ++++++++++++++++
DEBUG_MODE = True
REGISTRATION_EXCEL_FILE_NAME = "sports_registrations_2025-06-20_09-15-18.xlsx"
# ++++++++++++++++

PROJECT_PATH = Path(os.getcwd())
SRC_PATH = PROJECT_PATH / "src"
ASSETS_PATH = SRC_PATH / "assets"
BIN_PATH = SRC_PATH / "bin"
LIB_PATH = SRC_PATH / "lib"
LOG_PATH = SRC_PATH / "log"
OUT_PATH = SRC_PATH / "out"

INVOICE_MODELS_FOLDER_NAME = "invoice_models"

NUM_INVOICE_NAME = "1_N° facture.xlsx"
NUM_INVOICE_PATH = LIB_PATH / NUM_INVOICE_NAME

SPONSOR_DATABASE_NAME = "sponsor_database.xlsx"
SPONSOR_DATABASE_DEBUG_NAME = "sponsor_database_DEBUG.xlsx"
SHEET_NAME = "Sheet1"
if DEBUG_MODE:
    SPONSOR_DATABASE_PATH = LIB_PATH / SPONSOR_DATABASE_DEBUG_NAME
else:
    SPONSOR_DATABASE_PATH = LIB_PATH / SPONSOR_DATABASE_NAME


SPORTS_CATALOG_NAME = "sports_catalog.json"
SPORTS_CATALOG_PATH = LIB_PATH / SPORTS_CATALOG_NAME

CURRENT_TIME = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

if platform.system() == "Linux":
    # Ubuntu
    SOFFICE_BINARY_PATH = Path("/usr/lib/libreoffice/program/soffice")
else:
    # macOS
    SOFFICE_BINARY_PATH = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")

SPORTS_SHEET_NAME_LIST = ["Inscription Volley mixte", "Inscription Pétanque", "Inscription Tir à la corde"]
#SPORTS_LIST = ["Mixed Volleyball", "Pétanque", "Tug of War"]
SPORTS_LIST = ["Volley Mixte", "Pétanque", "Tir à la Corde"]
assert len(SPORTS_SHEET_NAME_LIST) == len(SPORTS_LIST), "SPORTS_SHEET_NAME_LIST and SPORTS_LIST must have the same length."
