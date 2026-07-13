from src.services.excel_reader import ExcelReader
from src.services.payslip_generator import PayslipGenerator

def start():

    print("Reading Excel...")

    employees = ExcelReader().load_employees("Payslips.xlsx")

    print(f"Employees : {len(employees)}")

    generator = PayslipGenerator()

    for emp in employees:
        print(f"Generating {emp.epf} - {emp.name}")
        generator.generate(emp)

    print("\nSUCCESS")