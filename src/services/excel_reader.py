from pathlib import Path
import pandas as pd

from src.models.employee import Employee


class ExcelReader:

    def __init__(self):
        self.input_folder = Path("input")

    def load_employees(self, filename):

        file_path = self.input_folder / filename

        df = pd.read_excel(file_path, header=1)

        employees = []

        for _, row in df.iterrows():

            employee = Employee(

                epf=row["EPF"],

                name=row["Name"],

                department=row["Name.1"],

                basic_salary=row["BRA+ Anodizing+Salary"],

                gross_salary=row["gross salary 1."],

                total_for_epf=row["Tot. for EPF"],

                epf8=row["EPF 8%"],

                epf12=row["EPF 12"],

                etf3=row["ETF 3"],

                paye=row["Paye tax"],

                meals=row["Meals"],

                salary_advance=row["Salary Advance"],

                total_deduction=row["Total Deduction"],

                net_salary=row["Net Sal"]

            )

            employees.append(employee)

        return employees