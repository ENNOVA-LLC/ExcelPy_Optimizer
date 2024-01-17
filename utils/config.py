"""`config` file

"""
from pathlib import Path

UTILS_DIR = Path(__file__).resolve().parent
ROOT_DIR = UTILS_DIR.parent
DATA_DIR = ROOT_DIR / 'data'
DOCS_DIR = ROOT_DIR / 'docs'