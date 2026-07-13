from src.services.excel_reader import ExcelReader
from src.services.payslip_generator import PayslipGenerator

INPUT_FILE = "Payslips.xlsx"


def start():
    print("Reading Excel...")

    reader = ExcelReader()
    employees = reader.load_employees(INPUT_FILE)
    pay_period = reader.get_pay_period(INPUT_FILE)

    print(f"Employees : {len(employees)}")
    print(f"Pay Period: {pay_period}")

    generator = PayslipGenerator()

    for emp in employees:
        print(f"Generating {emp.epf} - {emp.name}")
        generator.generate(emp, pay_period)

    print("\nSUCCESS")