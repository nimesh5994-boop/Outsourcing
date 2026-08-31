"""Unit tests for auto-detecting what an uploaded file is (report type,
platform, period) and for PDF table extraction - the two pieces that
replace manually picking report type/platform/period before every upload.
No database needed; these work on in-memory DataFrames/bytes like the rest
of the non-storage test suite."""
import io
from pathlib import Path

import pytest

from app import document_detection as dd
from app import parsers

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


def test_try_xero_native_matches_real_xero_export():
    src = parsers.FileDataSource(SAMPLE_DIR / "trial_balance_current_xero.xlsx")
    assert dd.try_xero_native(src) == "trial_balance"


def test_try_xero_native_returns_none_for_generic_csv():
    src = parsers.FileDataSource(SAMPLE_DIR / "bank_statement_current.csv")
    assert dd.try_xero_native(src) is None


def test_try_xero_native_does_not_corrupt_columns_on_failed_attempts():
    """Regression test for the bug where a failed Xero-native parse
    permanently rewrote the source's cached column headers to integer
    positions, breaking the generic mapping fallback that runs right
    after it on the same DataSource object."""
    src = parsers.FileDataSource(SAMPLE_DIR / "bank_statement_current.csv")
    before = src.raw_columns()
    assert dd.try_xero_native(src) is None
    assert src.raw_columns() == before


@pytest.mark.parametrize("filename,expected_type", [
    ("bank_statement_current.csv", "bank_statement"),
    ("fixed_asset_register_prior_year.csv", "fixed_asset_register"),
    ("vat_return_current.csv", "vat_return"),
])
def test_classify_report_type_generic_exports(filename, expected_type):
    src = parsers.FileDataSource(SAMPLE_DIR / filename)
    report_type, confidence = dd.classify_report_type(src.raw_columns())
    assert report_type == expected_type
    assert confidence > 0.5


def test_classify_report_type_unrecognisable_columns_returns_none():
    report_type, confidence = dd.classify_report_type(["Foo", "Bar", "Baz"])
    assert report_type is None
    assert confidence == 0.0


def test_disambiguate_pl_vs_bs_from_category_values():
    import pandas as pd
    pl_df = pd.DataFrame({"cat": ["Turnover", "Cost of Sales", "Overheads"]})
    bs_df = pd.DataFrame({"cat": ["Fixed Assets", "Current Liabilities", "Equity"]})
    assert dd.disambiguate_pl_vs_bs(pl_df, "cat") == "profit_and_loss"
    assert dd.disambiguate_pl_vs_bs(bs_df, "cat") == "balance_sheet"


def test_classify_platform_defaults_to_other():
    assert dd.classify_platform(["Account", "Amount"], is_xero_native=False) == "other"


def test_classify_platform_xero_native_short_circuits():
    assert dd.classify_platform(["anything"], is_xero_native=True) == "xero"


def test_guess_period_second_upload_of_same_type_is_comparative():
    """No date column to go on (a TB has none) - the fallback heuristic:
    if a confirmed 'current' upload of this report type already exists on
    the job, a second one is very likely last year's comparative."""
    src = parsers.FileDataSource(SAMPLE_DIR / "vat_return_current.csv")
    job_with_existing_current = {
        "current_period_end": "2025-12-31", "comparative_period_end": "2024-12-31",
        "uploads": {"u1": {"report_type": "vat_return", "period": "current", "confirmed": True}},
    }
    assert dd.guess_period(src, "vat_return", job_with_existing_current, None) == "comparative"

    job_with_no_existing = {
        "current_period_end": "2025-12-31", "comparative_period_end": "2024-12-31",
        "uploads": {},
    }
    assert dd.guess_period(src, "vat_return", job_with_no_existing, None) == "current"


def _make_generic_workbook(headers: list[str], rows: list[list]) -> bytes:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_apply_mapping_does_not_swap_day_and_month_on_iso_dates():
    """Regression test for a real bug: pd.to_datetime(..., dayfirst=True)
    without format="mixed" infers ONE format from the first row and
    forces every other row through it - an unambiguous ISO date in row 1
    ("2025-06-01") made pandas commit to day-first parsing for the whole
    column, silently corrupting later ISO rows (found live: "2025-06-15"
    became NaT entirely, and "2025-06-01" itself became 6 January instead
    of 1 June). This only shows up with >=2 rows and only when the first
    row's day is <=12 (so it reads as ambiguous) - a single-row upload or
    one where the first date has day >12 would parse correctly by luck,
    which is exactly why this stayed hidden until a real multi-row
    generic-mapped upload exercised it."""
    import pandas as pd
    content = _make_generic_workbook(
        ["Date", "Account Code", "Account Name", "Reference", "Description", "Contact", "Source Type", "Debit", "Credit"],
        [
            ["2025-06-01", "1100", "DEBTORS CONTROL", "INV-1", "Sale", "Acme Ltd", "Invoice", 5000, 0],
            ["2025-06-15", "2100", "CREDITORS CONTROL", "BILL-1", "Purchase", "Gamma Supplies", "Bill", 0, 3000],
        ],
    )
    source = parsers.FileDataSource(content, filename="gl.xlsx")
    from app import mapping
    suggestion = mapping.suggest_mapping("nominal_activity", source.raw_columns())
    df = parsers.apply_mapping(source, "nominal_activity", suggestion)
    assert list(df["date"]) == [pd.Timestamp("2025-06-01"), pd.Timestamp("2025-06-15")]


def test_apply_mapping_still_reads_uk_day_first_dates_correctly():
    """The fix (format="mixed") must not lose the whole point of
    dayfirst=True - a genuinely ambiguous UK-style date still needs to
    read as day-first, not month-first."""
    import pandas as pd
    content = _make_generic_workbook(
        ["Date", "Account Code", "Account Name", "Reference", "Description", "Contact", "Source Type", "Debit", "Credit"],
        [["01/03/2025", "1100", "DEBTORS CONTROL", "INV-1", "Sale", "Acme Ltd", "Invoice", 5000, 0]],
    )
    source = parsers.FileDataSource(content, filename="gl.xlsx")
    from app import mapping
    suggestion = mapping.suggest_mapping("nominal_activity", source.raw_columns())
    df = parsers.apply_mapping(source, "nominal_activity", suggestion)
    assert df["date"].iloc[0] == pd.Timestamp("2025-03-01")  # 1 March, not 3 January


def test_guess_period_does_not_swap_day_and_month_on_iso_dates():
    """Same fix, same regression, in document_detection.guess_period's
    own separate pd.to_datetime(dayfirst=True) call (used to score which
    column looks like a date column and find the latest value in it)."""
    import pandas as pd
    content = _make_generic_workbook(
        ["Invoice Date", "Customer", "Net Amount"],
        [["2025-06-01", "Acme Ltd", 500], ["2025-06-15", "Beta Ltd", 300]],
    )
    source = parsers.FileDataSource(content, filename="filed_sales.xlsx")
    latest = dd._latest_date_in_columns(source, "vat_filed_sales")
    assert latest == pd.Timestamp("2025-06-15")


def _make_test_pdf(rows: list[list[str]]) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    table = Table(rows)
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    doc.build([table])
    return buf.getvalue()


def test_pdf_table_extraction_round_trip():
    reportlab = pytest.importorskip("reportlab", reason="reportlab is a dev-only dependency for building test PDFs")
    from app.pdf_extraction import extract_table_from_pdf

    pdf_bytes = _make_test_pdf([
        ["Account Code", "Account Name", "Debit", "Credit"],
        ["1000", "Bank Current Account", "12000.00", ""],
        ["4000", "Sales", "", "75000.00"],
    ])
    df = extract_table_from_pdf(pdf_bytes)
    assert list(df.columns) == ["Account Code", "Account Name", "Debit", "Credit"]
    assert len(df) == 2
    assert df.iloc[0]["Account Name"] == "Bank Current Account"


def test_pdf_flows_through_file_data_source_and_classifier():
    pytest.importorskip("reportlab", reason="reportlab is a dev-only dependency for building test PDFs")
    pdf_bytes = _make_test_pdf([
        ["Account Code", "Account Name", "Debit", "Credit"],
        ["1000", "Bank Current Account", "12000.00", ""],
        ["4000", "Sales", "", "75000.00"],
    ])
    src = parsers.FileDataSource(pdf_bytes, filename="trial_balance.pdf")
    assert dd.try_xero_native(src) is None
    report_type, confidence = dd.classify_report_type(src.raw_columns())
    assert report_type == "trial_balance"
    assert confidence > 0.5


def test_pdf_with_no_table_raises_clear_error():
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import Paragraph, SimpleDocTemplate
    from reportlab.lib.styles import getSampleStyleSheet

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    doc.build([Paragraph("Just some text, no table here.", getSampleStyleSheet()["Normal"])])

    from app.pdf_extraction import extract_table_from_pdf
    with pytest.raises(ValueError, match="No table could be found"):
        extract_table_from_pdf(buf.getvalue())
