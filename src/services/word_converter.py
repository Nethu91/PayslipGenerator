from pathlib import Path
from pdf2docx import Converter


def pdf_to_word(pdf_path, docx_path=None):
    pdf_path = str(pdf_path)
    docx_path = str(docx_path) if docx_path else pdf_path.replace(".pdf", ".docx")

    cv = Converter(pdf_path)
    try:
        cv.convert(docx_path)
    finally:
        cv.close()

    return Path(docx_path)