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

    def _force_a4_single_page(self, ws):
        """
        Force the worksheet to print as one A4 portrait page, left-aligned.

        IMPORTANT: Zoom must be set to False BEFORE FitToPagesWide/Tall -
        Excel COM ignores fit-to-page while Zoom has a numeric value
        (its default is 100), which is why the PDF was coming out as
        2 pages even though the xlsx "looked" fine.
        """
        ps = ws.PageSetup
        ps.Orientation = 1          # xlPortrait
        ps.PaperSize = 9            # xlPaperA4
        ps.Zoom = False             # must come before FitToPages*
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
        ps.LeftMargin = 14          # ~0.2in in points (72pt = 1in)
        ps.RightMargin = 14
        ps.TopMargin = 18
        ps.BottomMargin = 18
        ps.HeaderMargin = 7
        ps.FooterMargin = 7
        ps.CenterHorizontally = False   # left-aligned, not centered on the page
        ps.CenterVertically = False

    def export(self, xlsx_path, pdf_path=None):
        xlsx_path = str(xlsx_path)
        pdf_path = str(pdf_path) if pdf_path else xlsx_path.replace(".xlsx", ".pdf")
        excel = self._get_excel()
        wb = excel.Workbooks.Open(xlsx_path)
        try:
            ws = wb.Worksheets(1)
            self._force_a4_single_page(ws)
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