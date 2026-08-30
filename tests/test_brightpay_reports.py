"""Unit tests for the BrightPay native report parsers (Payroll Summary,
P32, Pensions) - app/brightpay_reports.py. Structural parsing, same
reasoning as the Xero-native parsers: these are fixed exports from one
named platform with known column names, not something a preparer maps
column-by-column, so these tests build the real monthly-block CSV shape
directly (synthetic company/employee names - no real client data) rather
than going through the generic column-mapping path."""
import pandas as pd

from app.brightpay_reports import parse_p32, parse_payroll_summary, parse_pensions, try_brightpay_native
from app.parsers import FileDataSource

PAYROLL_SUMMARY_CSV = (
    "Name,Surname,Gross pay,Taxable gross,Tax,NIC-able gross,Employee NICs,"
    "Student + Postgrad Loan deduction,Net pay,Take-home pay,Employer NICs,Employer pension,Cost to employer\n"
    "Month 1 (Ending 30 April 2025)\n"
    "Jamie,Smith,2000.00,2000.00,200.00,2000.00,100.00,0.00,1700.00,1700.00,150.00,40.00,2190.00\n"
    "Alex,Doe,1000.00,1000.00,10.00,1000.00,5.00,0.00,985.00,985.00,0.00,0.00,1000.00\n"
    "TOTAL (2 employees),,3000.00,3000.00,210.00,3000.00,105.00,0.00,2685.00,2685.00,150.00,40.00,3190.00\n"
    "Month 2 (Ending 31 May 2025)\n"
    "Jamie,Smith,2000.00,2000.00,200.00,2000.00,100.00,0.00,1700.00,1700.00,150.00,40.00,2190.00\n"
    "TOTAL,,2000.00,2000.00,200.00,2000.00,100.00,0.00,1700.00,1700.00,150.00,40.00,2190.00\n"
    "Months 1 to 2 (Summary)\n"
    "Jamie,Smith,4000.00,4000.00,400.00,4000.00,200.00,0.00,3400.00,3400.00,300.00,80.00,4380.00\n"
    "Alex,Doe,1000.00,1000.00,10.00,1000.00,5.00,0.00,985.00,985.00,0.00,0.00,1000.00\n"
    "TOTAL (2 employees),,5000.00,5000.00,410.00,5000.00,205.00,0.00,4385.00,4385.00,300.00,80.00,5380.00\n"
)

P32_CSV = (
    "Tax period ending,Gross tax,Tax refund received,CIS deductions suffered,Student loan,Postgraduate loan,"
    "Net tax,Gross NICs,SMP recovered,NIC compensation on SMP,SPP recovered,NIC compensation on SPP,"
    "ShPP recovered,NIC compensation on ShPP,SAP recovered,NIC compensation on SAP,SPBP recovered,"
    "NIC compensation on SPBP,SNCP recovered,NIC compensation on SNCP,Employment allowance claim,"
    "Apprenticeship Levy,Total deductions from NICs,Net NICs,Amount due\n"
    "Tax Months 1 to 2 (Summary)\n"
    "05/05/2025,210.00,0.00,0.00,0.00,0.00,210.00,255.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,150.00,0.00,150.00,105.00,315.00\n"
    "05/06/2025,200.00,0.00,0.00,0.00,0.00,200.00,250.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,150.00,0.00,150.00,100.00,300.00\n"
)

PENSIONS_CSV = (
    "Name,Surname,Employee pensionable gross,Employee pension,Employee AVCs,Employer pensionable gross,"
    "Employer pension,Employer AVCs,Employee + employer pension\n"
    "Month 1 (Ending 30 April 2025)\n"
    "Jamie,Smith,2000.00,64.00,0.00,2000.00,40.00,0.00,104.00\n"
    "Alex,Doe,1000.00,0.00,0.00,1000.00,0.00,0.00,0.00\n"
    "TOTAL (2 employees),,3000.00,64.00,0.00,3000.00,40.00,0.00,104.00\n"
    "Months 1 to 1 (Summary)\n"
    "Jamie,Smith,2000.00,64.00,0.00,2000.00,40.00,0.00,104.00\n"
    "Alex,Doe,1000.00,0.00,0.00,1000.00,0.00,0.00,0.00\n"
    "TOTAL (2 employees),,3000.00,64.00,0.00,3000.00,40.00,0.00,104.00\n"
)


def _source(text: str, filename: str) -> FileDataSource:
    return FileDataSource(text.encode(), filename=filename)


def test_payroll_summary_parses_one_row_per_employee_per_month():
    df = parse_payroll_summary(_source(PAYROLL_SUMMARY_CSV, "payroll.csv"))
    assert len(df) == 3  # month 1 x2 employees + month 2 x1 employee - annual summary block excluded
    assert set(df["employee"]) == {"Jamie Smith", "Alex Doe"}
    assert df["period_end"].nunique() == 2
    jamie_m1 = df[(df["employee"] == "Jamie Smith") & (df["period_end"] == pd.Timestamp("2025-04-30"))].iloc[0]
    assert jamie_m1["net_pay"] == 1700.00
    assert jamie_m1["employer_nics"] == 150.00
    # TOTAL rows and the trailing "Months 1 to 2 (Summary)" block must not appear as data
    assert not (df["employee"].str.upper().str.startswith("TOTAL")).any()
    assert df["net_pay"].sum() == 1700.00 + 985.00 + 1700.00


def test_p32_parses_one_row_per_tax_month():
    df = parse_p32(_source(P32_CSV, "p32.csv"))
    assert len(df) == 2
    assert list(df["period_end"]) == [pd.Timestamp("2025-05-05"), pd.Timestamp("2025-06-05")]
    assert df.iloc[0]["amount_due"] == 315.00
    assert df.iloc[0]["employment_allowance"] == 150.00


def test_pensions_parses_one_row_per_employee_per_month_and_skips_annual_summary():
    df = parse_pensions(_source(PENSIONS_CSV, "pensions.csv"))
    assert len(df) == 2  # only month 1's two employees - the "Months 1 to 1 (Summary)" block is excluded
    jamie = df[df["employee"] == "Jamie Smith"].iloc[0]
    assert jamie["total_pension"] == 104.00
    assert jamie["employee_pension"] == 64.00
    assert jamie["employer_pension"] == 40.00


def test_try_brightpay_native_detects_each_report_type():
    assert try_brightpay_native(_source(PAYROLL_SUMMARY_CSV, "x.csv")) == "paye_summary"
    assert try_brightpay_native(_source(P32_CSV, "x.csv")) == "paye_p32"
    assert try_brightpay_native(_source(PENSIONS_CSV, "x.csv")) == "paye_pensions"


def test_try_brightpay_native_returns_none_for_unrelated_file():
    unrelated = "Date,Account Code,Description,Debit,Credit\n01/03/2025,4000,Sales,,1000\n"
    assert try_brightpay_native(_source(unrelated, "x.csv")) is None


def test_parse_payroll_summary_rejects_wrong_shaped_file():
    import pytest
    with pytest.raises(ValueError):
        parse_payroll_summary(_source(P32_CSV, "x.csv"))
