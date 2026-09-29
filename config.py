"""
Central configuration — every machine-specific value comes from environment
variables (or a .env file next to this script). See .env.example.
"""

import os
import shutil
import sys

from dotenv import load_dotenv

BASE_DIR = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))


def _path(var, default):
    return os.path.abspath(os.path.expandvars(os.getenv(var) or default))


# PostgreSQL database holding the "sommier" table
DB_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "port": int(os.getenv("DB_PORT", "5432")),
    "database": os.getenv("DB_NAME", "sommier"),
    "user": os.getenv("DB_USER", "postgres"),
    "password": os.getenv("DB_PASSWORD", ""),
}

# OCR / PDF tools
TESSERACT_CMD = (
    os.getenv("TESSERACT_CMD")
    or shutil.which("tesseract")
    or r"C:\Program Files\Tesseract-OCR\tesseract.exe"
)
TESSERACT_LANG = os.getenv("TESSERACT_LANG", "fra")
POPPLER_PATH = os.getenv("POPPLER_PATH") or None  # folder containing pdftoppm; None = use PATH

# Output folder and Excel templates
RESULT_DIR = _path("RESULT_DIR", os.path.join(BASE_DIR, "resultat"))
ATTESTATION_TEMPLATE = _path("ATTESTATION_TEMPLATE", os.path.join(BASE_DIR, "templates", "attestation.xlsx"))
CALCUL_TEMPLATE = _path("CALCUL_TEMPLATE", os.path.join(BASE_DIR, "templates", "calcul.xlsx"))
CALCUL_LOG_FILE = _path("CALCUL_LOG_FILE", os.path.join(BASE_DIR, "excelfiles", "default.xlsx"))
