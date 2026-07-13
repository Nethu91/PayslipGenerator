from pathlib import Path
from openpyxl import load_workbook


class PayslipGenerator:

    def __init__(self):
        self.template = Path("templates/Payslip_Template.xlsx")
        self.output = Path("output")
        self.output.mkdir(exist_ok=True)

    def generate(self, emp):

        filename = self.output / f"{emp.epf}_{emp.name}.xlsx"

        wb = load_workbook(self.template)
        ws = wb.active

        # ---------------- HEADER ----------------

        ws["B2"] = emp.epf
        ws["B3"] = emp.department
        ws["B5"] = emp.name
        ws["B6"] = "June 2026"

        # ---------------- EARNINGS ----------------

        ws["B10"] = emp.basic_salary
        ws["B15"] = emp.total_for_epf
        ws["B35"] = emp.gross_salary

        # ---------------- DEDUCTIONS ----------------

        ws["B38"] = emp.epf8
        ws["B39"] = emp.salary_advance
        ws["B42"] = emp.meals
        ws["B47"] = emp.paye
        ws["B49"] = emp.total_deduction

        # ---------------- NET ----------------

        ws["B50"] = emp.net_salary

        # ---------------- EMPLOYER ----------------

        ws["B53"] = emp.epf12
        ws["B54"] = emp.etf3

        wb.save(filename)
        wb.close()