from dataclasses import dataclass


@dataclass
class Employee:
    """Represents one employee's payroll data for a single pay period."""

    epf: str
    name: str
    department: str

    # Earnings
    basic_salary: float = 0
    holiday_pay: float = 0
    no_pay: float = 0
    nopay_revise: float = 0
    salary_adjustment: float = 0

    night_shift: float = 0
    saturday_night: float = 0
    sunday_payment: float = 0
    poya_mercantile: float = 0
    mercantile_special_payment: float = 0
    refund_iceu_amount: float = 0
    ot_arrears: float = 0
    salary_arrears_deduction: float = 0
    over_time: float = 0
    double_time: float = 0
    powder_incentive: float = 0
    tailoring_fee: float = 0
    att_incentive: float = 0
    melt_incentive: float = 0
    team_leader_incentive: float = 0
    chemical_incentive: float = 0
    export_incentive: float = 0
    production_incentive: float = 0

    # Deductions
    salary_advance: float = 0
    festival_advance: float = 0
    vision_care: float = 0
    meals: float = 0
    welfare_society: float = 0
    other_deduction_tp: float = 0
    festival_advance_guarantors: float = 0
    iceu_member_fee: float = 0
    paye_tax: float = 0
    stamp_duty: float = 0

    # Info-only (not part of salary formulas)
    casual_leave: float = 0
    rate_per_not: float = 0
    rate_per_dot: float = 0