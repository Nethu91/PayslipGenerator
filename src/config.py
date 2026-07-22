import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent          # exe's folder (for input/output)
    BUNDLE_DIR = Path(sys._MEIPASS)                             # bundled read-only resources
else:
    BASE_DIR = Path(__file__).resolve().parent.parent
    BUNDLE_DIR = BASE_DIR

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
EXCEL_DIR = OUTPUT_DIR / "Excel"
PDF_DIR = OUTPUT_DIR / "PDF"
ASSETS_DIR = BUNDLE_DIR / "assets"

TEMPLATE_FILE = BASE_DIR / "templates" / "Payslip_Template.xlsx"   # external, editable - not bundled

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
EXCEL_DIR.mkdir(exist_ok=True)
PDF_DIR.mkdir(exist_ok=True)