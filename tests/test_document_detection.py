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


def test_vat_return_box_summary_native_detection_and_parsing():
    """Regression: Xero's exported VAT Return isn't a table - it's a
    vertical label/box-number/value listing (one row per HMRC box), with
    title/scheme-detail rows above it that carry no box number. A real
    client's VAT Return upload came through as a 3-sheet workbook (this
    report plus two structurally unrelated detail sheets), and the
    generic column-mapper found nothing to map on any of the three - the
    top-level VAT cross-check silently had nothing to work with. Now a
    Xero-native report type: try_xero_native must recognise the genuine
    box summary sheet and reject the two that aren't."""
    from app import xero_reports

    box_summary = _make_generic_workbook(
        ["", "", ""],
        [
            ["Acme Ltd", None, None],
            ["For the period 01 Apr 2025 - 30 Jun 2025", None, None],
            ["VAT Calculations", None, None],
            ["VAT due in the period on sales and other outputs", "1", 492.36],
            ["VAT due in the period on acquisitions", "2", 0],
            ["Total VAT due (the sum of boxes 1 and 2)", "3", 492.36],
            ["VAT reclaimed in the period on purchases", "4", 0],
            ["VAT to pay HMRC", "5", 492.36],
            ["Total value of sales excluding VAT", "6", 3938],
            ["Total value of purchases excluding VAT", "7", 0],
        ],
    )
    src = parsers.FileDataSource(box_summary, filename="vat_return.xlsx")
    assert dd.try_xero_native(src) == "vat_return"
    df = xero_reports.parse_vat_return_box_summary(src)
    row = df.iloc[0]
    assert row["box1"] == 492.36
    assert row["box5"] == 492.36
    assert row["box6"] == 3938
    assert row["box8"] == 0.0  # omitted from this export - defaults to 0

    unrelated_detail_sheet = _make_generic_workbook(
        ["Tax Rate", "Net", "VAT"],
        [["20% (VAT on Income)", "1000.00", "200.00"]],
    )
    src2 = parsers.FileDataSource(unrelated_detail_sheet, filename="vat_return.xlsx")
    assert dd.try_xero_native(src2) is None
    with pytest.raises(ValueError, match="couldn't find boxes"):
        xero_reports.parse_vat_return_box_summary(src2)


def test_vat_box_transactions_splits_one_upload_into_box1_and_box4():
    """Regression: Xero's 'Transactions by VAT Box' export has the actual
    transaction-level detail behind boxes 1 and 4 - exactly the
    vat_filed_sales/vat_filed_purchases data the VAT Reconciliation
    workspace needs, normally sourced from a separate filed-return
    detail file the client may not readily have. Found live: a real
    client's Box 4 section had FOUR separate tax-rate sub-groups (20% on
    Expenses, 20% on Expenses - Adjusted, Zero Rated Expenses,
    Adjustments with no accounting transactions), each with its own
    repeated 'Date | Account | Reference | Details | VAT | Net' header -
    the detail-row scan must re-enter on every repeated header, not
    assume one header per box. Boxes 6/7 repeat the same transactions
    net-only (no VAT column) and must be skipped as duplicates."""
    from app import xero_reports

    rows = [
        ["Box 1", "VAT due in the period on sales and other outputs", "", "", 100.0, ""],
        ["20% (VAT on Income)", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["01/03/2025", "Sales(200)", "INV-1", "Acme Ltd", 60.0, 300.0],
        ["05/03/2025", "Sales(200)", "INV-2", "Beta Ltd", 40.0, 200.0],
        [None, None, None, None, None, None],
        ["Box 4", "VAT reclaimed in the period on purchases and other inputs", "", "", 90.0, ""],
        ["20% (VAT on Expenses)", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["02/03/2025", "Motor Vehicle Expenses(449)", "", "Shell", 10.0, 50.0],
        [None, None, None, None, None, None],
        ["Zero Rated Expenses", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["03/03/2025", "General Expenses(429)", "", "Royal Mail", 0.0, 25.0],
        [None, None, None, None, None, None],
        ["Adjustments with no accounting transactions", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["28/03/2025", "N/A", "Created on 01/04/2025", "Van purchase VAT reclaim", 80.0, 0.0],
        [None, None, None, None, None, None],
        ["Box 6", "Total value of sales excluding VAT", "", "", "", 500.0],
        ["20% (VAT on Income)", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["01/03/2025", "Sales(200)", "INV-1", "Acme Ltd", "", 300.0],
    ]
    content = _make_generic_workbook(["", "", "", "", "", ""], rows)
    src = parsers.FileDataSource(content, filename="vat_box.xlsx")

    result = xero_reports.parse_vat_box_transactions(src)
    assert set(result.keys()) == {1, 4}  # Box 6 correctly skipped - same transactions as Box 1, net-only

    box1 = result[1]
    assert len(box1) == 2
    assert box1["vat_amount"].sum() == pytest.approx(100.0)
    assert box1["net_amount"].sum() == pytest.approx(500.0)
    assert set(box1["contact"]) == {"Acme Ltd", "Beta Ltd"}

    box4 = result[4]
    assert len(box4) == 3  # one row from each of the three tax-rate sub-groups
    assert box4["vat_amount"].sum() == pytest.approx(90.0)  # 10 + 0 + 80
    assert set(box4["contact"]) == {"Shell", "Royal Mail", "Van purchase VAT reclaim"}


def test_vat_box_transactions_handles_the_quarterly_shifted_column_shape():
    """Regression: found live across several real periods for the same
    client - a longer (quarterly) period's export inserts an extra
    leading column holding the *current* box number, repeated on every
    single row of that box's section (not just its header row) - e.g.
    column 0 says 'Box 1' for 160+ consecutive rows, then switches to
    'Box 4' for the next section, etc. This shifts Date/Account/../Net
    one column to the right of where a shorter (monthly) period's export
    puts them (column 0 IS the Date value there), and turned a fixed-
    column "is this cell exactly 'Box N'" check into a false match on
    every ordinary data row - since that leading column keeps repeating
    the box label, "start a new section" fired again on every single
    data row, immediately discarding it. This reproduces that shifted
    shape and confirms parsing still finds the real header/data by
    content, not by column position."""
    from app import xero_reports

    rows = [
        ["Box 1", "Box 1", "VAT due in the period on sales and other outputs", "", "", 60.0, ""],
        ["Box 1", "20% (VAT on Income)", None, None, None, None, None],
        ["Box 1", "Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["Box 1", "01/03/2025", "Sales(200)", "INV-1", "Acme Ltd", 60.0, 300.0],
        ["Box 4", "Box 4", "VAT reclaimed in the period on purchases and other inputs", "", "", 40.0, ""],
        ["Box 4", "20% (VAT on Expenses)", None, None, None, None, None],
        ["Box 4", "Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["Box 4", "02/03/2025", "Motor Vehicle Expenses(449)", "", "Shell", 40.0, 200.0],
    ]
    content = _make_generic_workbook(["", "", "", "", "", "", ""], rows)
    src = parsers.FileDataSource(content, filename="vat_box_quarterly.xlsx")

    result = xero_reports.parse_vat_box_transactions(src)
    assert set(result.keys()) == {1, 4}
    assert len(result[1]) == 1
    assert result[1].iloc[0]["contact"] == "Acme Ltd"
    assert result[1]["vat_amount"].sum() == pytest.approx(60.0)
    assert len(result[4]) == 1
    assert result[4].iloc[0]["contact"] == "Shell"
    assert result[4]["vat_amount"].sum() == pytest.approx(40.0)


def test_vat_box_transactions_does_not_double_count_ec_acquisition_reclaim():
    """Regression: a Northern Ireland/EU acquisition affects boxes 2 and 4
    (and 9) from the SAME underlying transaction, so a real client's
    export cross-references it into the Box 4 section as a duplicated
    pair of tax-rate sub-groups sharing the identical VAT figure -
    'EC Acquisitions (20%)' (the acquisition-due side, genuinely Box 2's
    own transaction) immediately followed by 'EC Acquisitions (20%)
    Reclaimed VAT' (the actual Box 4 input-VAT reclaim). Collecting both
    double-counted it into Box 4's total by exactly that transaction's
    VAT amount - found live, a real file's Box 4 sum came out £33.38
    over its own VAT Return summary because of this one pair. Only the
    'Reclaimed VAT' sub-group belongs to Box 4."""
    from app import xero_reports

    rows = [
        ["Box 4", "VAT reclaimed in the period on purchases and other inputs", "", "", 100.0, ""],
        ["20% (VAT on Expenses)", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["02/03/2025", "Motor Vehicle Expenses(449)", "", "Shell", 66.62, 333.10],
        [None, None, None, None, None, None],
        ["EC Acquisitions (20%)", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["01/05/2025", "Motor Vehicle Expenses(449)", "SI-1", "West Tanfield Garage Ltd", 33.38, 166.91],
        [None, None, None, None, None, None],
        ["EC Acquisitions (20%) Reclaimed VAT", None, None, None, None, None],
        ["Date", "Account", "Reference", "Details", "VAT", "Net"],
        ["01/05/2025", "Motor Vehicle Expenses(449)", "SI-1", "West Tanfield Garage Ltd", 33.38, 0.0],
    ]
    content = _make_generic_workbook(["", "", "", "", "", ""], rows)
    src = parsers.FileDataSource(content, filename="vat_box_ec.xlsx")

    result = xero_reports.parse_vat_box_transactions(src)
    box4 = result[4]
    assert len(box4) == 2  # the ordinary Shell row + exactly ONE EC acquisition row, not two
    assert box4["vat_amount"].sum() == pytest.approx(100.0)  # 66.62 + 33.38, not 133.38
    reclaim_rows = box4[box4["reference"] == "SI-1"]
    assert len(reclaim_rows) == 1


def test_vat_box_transactions_rejects_unrelated_workbook():
    from app import xero_reports

    unrelated = _make_generic_workbook(
        ["Tax Rate", "Net", "VAT"],
        [["20% (VAT on Income)", "1000.00", "200.00"]],
    )
    src = parsers.FileDataSource(unrelated, filename="unrelated.xlsx")
    with pytest.raises(ValueError):
        xero_reports.parse_vat_box_transactions(src)


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


def test_pdf_extraction_falls_back_to_text_strategy_and_finds_the_real_header():
    """Regression: a real client's Aged Payables Detail PDF had no visible
    gridlines around its data rows - only its Total/Percentage summary
    rows happened to sit on a drawn line - so pdfplumber's default line-
    based table detection found just those two 1-row fragments and this
    raised "no table found" on a perfectly real, non-scanned export. This
    reproduces the same shape: a borderless table (no GRID style) with
    title/client-name/period rows ahead of the real multi-column header -
    extract_table_from_pdf must fall back to the text-position strategy
    (which needs no drawn borders) and correctly pick the real header row
    (the one with the most filled cells), not the title row above it."""
    pytest.importorskip("reportlab", reason="reportlab is a dev-only dependency for building test PDFs")
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table

    rows = [
        ["Aged Payables Detail", "", "", ""],
        ["Acme Ltd", "", "", ""],
        ["As at 31 December 2025", "", "", ""],
        ["Invoice Date", "Reference", "Contact", "Total"],
        ["01 Jan 2025", "INV-1", "Acme Supplier", "100.00"],
        ["02 Jan 2025", "INV-2", "Beta Supplier", "200.00"],
    ]
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    doc.build([Table(rows)])  # deliberately no TableStyle/GRID - no drawn borders at all

    from app.pdf_extraction import extract_table_from_pdf
    df = extract_table_from_pdf(buf.getvalue())

    assert "Aged Payables Detail" not in df.columns  # the title row must not be mistaken for the header
    assert "Reference" in df.columns
    assert "Contact" in df.columns
    assert set(df["Reference"]) >= {"INV-1", "INV-2"}


def test_pdf_extraction_rejects_a_genuinely_empty_report_instead_of_a_one_column_table():
    """Regression: a real client's Aged Receivables Detail PDF for a
    client with zero outstanding receivables has no data table at all -
    just a title/client-name/period/footer block. The text-strategy
    fallback (added for the borderless-table case above) clustered that
    into a bogus 9-row, 1-column "table" (['Aged Receiv'], [''],
    ["Shpendi'sLtd"], ...), which passed the len(table) >= 2 sanity check
    and produced a nonsense single-column DataFrame instead of a clear
    error - this then got auto-classified as an unrecognisable upload
    with report_type left blank, stranding the file in Unclassified with
    no way to map it (there's nothing to map). extract_table_from_pdf
    must recognise this isn't real tabular data (its header row has only
    one filled cell, not one per column) and raise the "no table" error,
    which the upload route already turns into a clear parse_error instead
    of corrupting the batch."""
    pytest.importorskip("reportlab", reason="reportlab is a dev-only dependency for building test PDFs")
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table

    rows = [
        ["Aged Receivables Detail"],
        [""],
        ["Acme Ltd"],
        [""],
        ["As at 31 March 2026"],
        [""],
        ["Ageing by due date"],
    ]
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    doc.build([Table(rows)])  # no GRID - and only ever one cell per row, no real columns

    from app.pdf_extraction import extract_table_from_pdf
    with pytest.raises(ValueError, match="No table could be found"):
        extract_table_from_pdf(buf.getvalue())


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
