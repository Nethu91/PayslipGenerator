from dataclasses import dataclass


@dataclass
class Employee:
    epf: int
    name: str
    department: str

    basic_salary: float
    gross_salary: float
    total_for_epf: float

    epf8: float
    epf12: float
    etf3: float

    paye: float

    meals: float
    salary_advance: float

    total_deduction: float
    net_salary: float