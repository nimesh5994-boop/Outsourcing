"""BrightPay report parsers - structural, not generic column-mapping, the
same reasoning as xero_reports.py: these are fixed exports from one named
platform with known column names, not something a preparer maps column-
by-column. BrightPay's CSV/PDF exports for Payroll Summary and Pensions
share one shape (a repeating "Month N (Ending <date>)" section per tax
month, employee rows underneath, a TOTAL row closing each section, then a
final "Months X to Y (Summary)" annual-aggregate section that's
deliberately NOT parsed here - it's exactly the sum of the months already
captured, so including it would double-count); P32 is flatter - one
"Tax Months X to Y (Summary)" title row, then one data row per tax month
directly, no employee breakdown at all (P32 is a company-wide HMRC filing,
never employee-level).

Detection mirrors document_detection.try_xero_native: each parser is
tried in turn, and a genuine structural match (the file's real BrightPay
column headers, not just similar-looking ones) is a near-certain signal -
much stronger than scoring column names the way the generic path does.
"""
from __future__ import annotations

import re

import pandas as pd

from app.parsers import DataSource

_MONTH_HEADER_RE = re.compile(r"^Month\s+\d+\s+\(Ending\s+(.+)\)$")
_ANNUAL_SUMMARY_RE = re.compile(r"^Months?\s+\d+\s+to\s+\d+\s+\(Summary\)$")
_UK_DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")

PAYROLL_SUMMARY_COLUMNS = [
    "Name", "Surname", "Gross pay", "Taxable gross", "Tax", "NIC-able gross", "Employee NICs",
    "Student + Postgrad Loan deduction", "Net pay", "Take-home pay", "Employer NICs",
    "Employer pension", "Cost to employer",
]
PENSIONS_COLUMNS = [
    "Name", "Surname", "Employee pensionable gross", "Employee pension", "Employee AVCs",
    "Employer pensionable gross", "Employer pension", "Employer AVCs", "Employee + employer pension",
]
P32_COLUMNS = [
    "Tax period ending", "Gross tax", "Tax refund received", "CIS deductions suffered", "Student loan",
    "Postgraduate loan", "Net tax", "Gross NICs", "Employment allowance claim", "Apprenticeship Levy",
    "Total deductions from NICs", "Net NICs", "Amount due",
]

BRIGHTPAY_NATIVE_REPORT_TYPES = {"paye_summary", "paye_p32", "paye_pensions"}


def _num(value) -> float:
    value = str(value).strip().replace(",", "")
    return float(value) if value not in ("", "-") else 0.0


def try_brightpay_native(source: DataSource) -> str | None:
    for report_type in BRIGHTPAY_NATIVE_REPORT_TYPES:
        try:
            if report_type == "paye_summary":
                parse_payroll_summary(source)
            elif report_type == "paye_p32":
                parse_p32(source)
            elif report_type == "paye_pensions":
                parse_pensions(source)
            return report_type
        except Exception:
            continue
    return None


def _require_columns(columns: list[str], expected: list[str], label: str) -> None:
    missing = [c for c in expected if c not in columns]
    if missing:
        raise ValueError(f"Not a BrightPay {label} export - missing column(s) {missing}")


def _parse_monthly_employee_report(source: DataSource, expected_columns: list[str], value_columns: list[str], label: str) -> pd.DataFrame:
    """Shared shape for Payroll Summary and Pensions: employee rows nested
    under "Month N (Ending <date>)" section headers, a TOTAL row closing
    each section, stopping at the trailing annual "Months X to Y
    (Summary)" section rather than parsing it as more months."""
    raw = source.raw_dataframe()
    _require_columns(list(raw.columns), expected_columns, label)

    rows = []
    period_end = None
    for _, row in raw.iterrows():
        name_cell = str(row["Name"]).strip()
        month_match = _MONTH_HEADER_RE.match(name_cell)
        if month_match:
            period_end = pd.to_datetime(month_match.group(1), dayfirst=True, errors="coerce")
            continue
        if _ANNUAL_SUMMARY_RE.match(name_cell):
            break
        if not name_cell or name_cell.upper().startswith("TOTAL") or period_end is None:
            continue
        surname = str(row["Surname"]).strip()
        employee = f"{name_cell} {surname}".strip()
        rows.append({
            "period_end": period_end, "employee": employee,
            **{col: _num(row[col]) for col in value_columns},
        })
    return pd.DataFrame(rows)


def parse_payroll_summary(source: DataSource) -> pd.DataFrame:
    """One row per employee per tax month. Columns: period_end, employee,
    gross_pay, taxable_gross, tax, nic_able_gross, employee_nics,
    student_loan, net_pay, take_home_pay, employer_nics, employer_pension,
    cost_to_employer."""
    value_columns = [
        "Gross pay", "Taxable gross", "Tax", "NIC-able gross", "Employee NICs",
        "Student + Postgrad Loan deduction", "Net pay", "Take-home pay",
        "Employer NICs", "Employer pension", "Cost to employer",
    ]
    df = _parse_monthly_employee_report(source, PAYROLL_SUMMARY_COLUMNS, value_columns, "Payroll Summary")
    return df.rename(columns={
        "Gross pay": "gross_pay", "Taxable gross": "taxable_gross", "Tax": "tax",
        "NIC-able gross": "nic_able_gross", "Employee NICs": "employee_nics",
        "Student + Postgrad Loan deduction": "student_loan", "Net pay": "net_pay",
        "Take-home pay": "take_home_pay", "Employer NICs": "employer_nics",
        "Employer pension": "employer_pension", "Cost to employer": "cost_to_employer",
    })


def parse_pensions(source: DataSource) -> pd.DataFrame:
    """One row per employee per tax month. Columns: period_end, employee,
    employee_pensionable_gross, employee_pension, employee_avcs,
    employer_pensionable_gross, employer_pension, employer_avcs,
    total_pension."""
    value_columns = [
        "Employee pensionable gross", "Employee pension", "Employee AVCs",
        "Employer pensionable gross", "Employer pension", "Employer AVCs", "Employee + employer pension",
    ]
    df = _parse_monthly_employee_report(source, PENSIONS_COLUMNS, value_columns, "Pensions")
    return df.rename(columns={
        "Employee pensionable gross": "employee_pensionable_gross", "Employee pension": "employee_pension",
        "Employee AVCs": "employee_avcs", "Employer pensionable gross": "employer_pensionable_gross",
        "Employer pension": "employer_pension", "Employer AVCs": "employer_avcs",
        "Employee + employer pension": "total_pension",
    })


def parse_p32(source: DataSource) -> pd.DataFrame:
    """One row per tax month - P32 is a company-wide HMRC filing, never
    broken down by employee. Columns: period_end, gross_tax, tax_refund,
    cis_deductions, student_loan, postgrad_loan, net_tax, gross_nics,
    employment_allowance, apprenticeship_levy, net_nics, amount_due."""
    raw = source.raw_dataframe()
    _require_columns(list(raw.columns), P32_COLUMNS, "P32")

    rows = []
    for _, row in raw.iterrows():
        period_cell = str(row["Tax period ending"]).strip()
        if not _UK_DATE_RE.match(period_cell):
            continue  # the "Tax Months X to Y (Summary)" title row, not a data row
        rows.append({
            "period_end": pd.to_datetime(period_cell, dayfirst=True, errors="coerce"),
            "gross_tax": _num(row["Gross tax"]), "tax_refund": _num(row["Tax refund received"]),
            "cis_deductions": _num(row["CIS deductions suffered"]), "student_loan": _num(row["Student loan"]),
            "postgrad_loan": _num(row["Postgraduate loan"]), "net_tax": _num(row["Net tax"]),
            "gross_nics": _num(row["Gross NICs"]), "employment_allowance": _num(row["Employment allowance claim"]),
            "apprenticeship_levy": _num(row["Apprenticeship Levy"]), "net_nics": _num(row["Net NICs"]),
            "amount_due": _num(row["Amount due"]),
        })
    if not rows:
        raise ValueError("Not a BrightPay P32 export - no tax-month data rows found")
    return pd.DataFrame(rows)
