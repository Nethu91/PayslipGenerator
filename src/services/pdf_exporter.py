import win32com.client as win32


class PdfExporter:
    """Converts a generated payslip .xlsx into .pdf using installed Excel."""

    def __init__(self):
        self.excel = None

    def _get_excel(self):
        if self.excel is None:
            self.excel = win32.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
        return self.excel

    def export(self, xlsx_path, pdf_path=None):
        xlsx_path = str(xlsx_path)
        pdf_path = str(pdf_path) if pdf_path else xlsx_path.replace(".xlsx", ".pdf")

        excel = self._get_excel()
        wb = excel.Workbooks.Open(xlsx_path)
        try:
            ws = wb.Worksheets(1)
            ws.ExportAsFixedFormat(0, pdf_path)  # 0 = xlTypePDF
        finally:
            wb.Close(SaveChanges=False)

        return pdf_path

    def close(self):
        if self.excel is not None:
            try:
                if self.excel.Workbooks.Count == 0:
                    self.excel.Quit()
            except Exception:
                pass
            finally:
                self.excel = None