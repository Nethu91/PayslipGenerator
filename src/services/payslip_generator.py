from pathlib import Path
from openpyxl import load_workbook


class PayslipGenerator:

    def __init__(self, template_path=None, output_folder=None):
        from src.config import TEMPLATE_FILE, OUTPUT_DIR
        self.template = Path(template_path) if template_path else TEMPLATE_FILE
        self.output = Path(output_folder) if output_folder else OUTPUT_DIR
        self.output.mkdir(exist_ok=True)

    def generate(self, emp, pay_period=""):
        filename = self.output / f"{emp.epf}_{emp.name}.xlsx"

        wb = load_workbook(self.template)
        ws = wb.active

        # ---------------- HEADER ----------------
        ws["B2"] = emp.epf
        ws["B3"] = emp.department
        ws["B5"] = emp.name
        ws["B6"] = pay_period

        # ---------------- EARNINGS (input cells only) ----------------
        ws["B10"] = emp.basic_salary
        ws["B11"] = emp.holiday_pay
        ws["B12"] = emp.no_pay
        ws["B13"] = emp.nopay_revise
        ws["B14"] = emp.salary_adjustment
        # B15 Total for EPF -> left as formula =SUM(B10:B14)

        ws["B17"] = emp.night_shift
        ws["B18"] = emp.saturday_night
        ws["B19"] = emp.sunday_payment
        ws["B20"] = emp.poya_mercantile
        ws["B21"] = emp.mercantile_special_payment
        ws["B22"] = emp.refund_iceu_amount
        ws["B23"] = emp.ot_arrears
        ws["B24"] = emp.salary_arrears_deduction
        ws["B25"] = emp.over_time
        ws["B26"] = emp.double_time
        ws["B27"] = emp.powder_incentive
        ws["B28"] = emp.tailoring_fee
        ws["B29"] = emp.att_incentive
        ws["B30"] = emp.melt_incentive
        ws["B31"] = emp.team_leader_incentive
        ws["B32"] = emp.chemical_incentive
        ws["B33"] = emp.export_incentive
        ws["B34"] = emp.production_incentive
        # B35 Gross Salary -> left as formula =B15+SUM(B17:B34)

        # ---------------- DEDUCTIONS (input cells only) ----------------
        # B38 EPF 8% -> left as formula =B15*0.08
        ws["B39"] = emp.salary_advance
        ws["B40"] = emp.festival_advance
        ws["B41"] = emp.vision_care
        ws["B42"] = emp.meals
        ws["B43"] = emp.welfare_society
        ws["B44"] = emp.other_deduction_tp
        ws["B45"] = emp.festival_advance_guarantors
        ws["B46"] = emp.iceu_member_fee
        ws["B47"] = emp.paye_tax
        ws["B48"] = emp.stamp_duty
        # B49 Total Deductions -> formula =SUM(B38:B48)
        # B50 Net Salary -> formula =B35-B49
        # B51 To Be Paid Amount -> formula =B50

        # ---------------- EMPLOYER / INFO ----------------
        # B53 EPF 12% -> formula =B15*0.12
        # B54 ETF 3% -> formula =B15*0.03
        ws["B55"] = emp.casual_leave
        ws["B56"] = emp.rate_per_not
        ws["B57"] = emp.rate_per_dot

        wb.save(filename)
        wb.close()