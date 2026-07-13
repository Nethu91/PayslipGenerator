from pathlib import Path

# Project root (this file is src/config.py, so parent.parent = project root)
BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
TEMPLATE_DIR = BASE_DIR / "templates"
ASSETS_DIR = BASE_DIR / "assets"

TEMPLATE_FILE = TEMPLATE_DIR / "Payslip_Template.xlsx"

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)