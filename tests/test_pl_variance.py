"""Unit tests for the P&L variance review's pure logic - see
app/pl_variance.py. No database needed."""
import pandas as pd

from app import pl_variance


def _pl(rows):
    return pd.DataFrame(rows, columns=["account_code", "account_name", "category", "amount"])


def _gl(rows):
    return pd.DataFrame(rows, columns=["account_code", "account_name", "date", "reference", "description", "contact", "debit", "credit"])


def test_flags_accounts_that_move_beyond_materiality_and_threshold():
    cur = _pl([
        ["4000", "Sales", "Turnover", -120000],
        ["6000", "Marketing", "Overheads", 25000],
        ["6100", "Rent", "Overheads", 12000],
    ])
    comp = _pl([
        ["4000", "Sales", "Turnover", -100000],
        ["6000", "Marketing", "Overheads", 8000],
        ["6100", "Rent", "Overheads", 11500],
    ])
    result = pl_variance.pl_variance_analysis(cur, comp, None, materiality=500, variance_pct_threshold=0.1)
    assert result.status == "review"
    flags = {row["account_name"]: row["flag"] for _, row in result.detail.iterrows()}
    assert flags["Sales"] is True
    assert flags["Marketing"] is True
    assert flags["Rent"] is False  # £500/4.3% - below the 10% threshold


def test_ok_status_when_nothing_moves_beyond_threshold():
    cur = _pl([["4000", "Sales", "Turnover", -100000]])
    comp = _pl([["4000", "Sales", "Turnover", -99000]])
    result = pl_variance.pl_variance_analysis(cur, comp, None, materiality=500, variance_pct_threshold=0.1)
    assert result.status == "ok"
    assert result.detail.iloc[0]["flag"] == False  # noqa: E712


def test_na_status_with_no_current_year_pl():
    result = pl_variance.pl_variance_analysis(None, None, None, materiality=500, variance_pct_threshold=0.1)
    assert result.status == "n/a"


def test_unmapped_category_shows_as_unclassified_not_blank():
    """Regression: an unmapped canonical field comes through from
    parsers.apply_mapping as an empty string, not NaN (it fills every
    unmapped non-numeric field with "" - see apply_mapping) - a bare
    fillna(NaN) alone silently leaves the category column blank in the
    real generated workbook instead of falling back to 'Unclassified'."""
    cur = _pl([["4000", "Sales", "", -120000]])
    comp = _pl([["4000", "Sales", "", -100000]])
    result = pl_variance.pl_variance_analysis(cur, comp, None, materiality=500, variance_pct_threshold=0.1)
    assert result.detail.iloc[0]["category"] == "Unclassified"


def test_missing_comparative_treated_as_zero_and_flags_the_whole_new_balance():
    cur = _pl([["4000", "Sales", "Turnover", -50000]])
    result = pl_variance.pl_variance_analysis(cur, None, None, materiality=500, variance_pct_threshold=0.1)
    assert result.status == "review"
    row = result.detail.iloc[0]
    assert row["comparative_year"] == 0.0
    assert row["variance_pct"] == 1.0
    assert row["flag"] == True  # noqa: E712


def test_split_across_codes_surfaces_a_contact_posting_to_more_than_one_account():
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal = _gl([
        ["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 15000, 0],
        ["6200", "Travel", "2025-04-01", "INV2", "travel", "Acme Agency", 4000, 0],
        ["6000", "Marketing", "2025-05-01", "INV3", "campaign", "Acme Agency", 10000, 0],
        ["6300", "Postage", "2025-06-01", "INV4", "stamps", "Solo Ltd", 50, 0],  # only ever posts to one code
    ])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    assert result.status == "review"
    assert not result.extra_detail.empty
    contacts = set(result.extra_detail["Contact"])
    assert contacts == {"Acme Agency"}  # Solo Ltd only posted to one code - nothing to compare
    accounts = set(result.extra_detail["Account"])
    assert accounts == {"6000 - Marketing", "6200 - Travel"}


def test_split_across_codes_empty_when_no_contact_posts_to_more_than_one_code():
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal = _gl([["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 25000, 0]])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    assert result.extra_detail.empty


def test_annotate_with_history_flags_recurring_vs_new_patterns():
    detail = pd.DataFrame([
        {"account_code": "4000", "account_name": "Sales", "flag": True},
        {"account_code": "6000", "account_name": "Marketing", "flag": True},
        {"account_code": "6100", "account_name": "Rent", "flag": False},
    ])
    history = {
        "2024-12-31": {
            "4000|Sales": {"flagged": True},
            "6000|Marketing": {"flagged": False},
            "6100|Rent": {"flagged": False},
        },
    }
    annotated = pl_variance.annotate_with_history(detail, history)
    patterns = dict(zip(annotated["account_name"], annotated["historical_pattern"]))
    assert "Recurring" in patterns["Sales"]
    assert "New this year" in patterns["Marketing"]
    assert "Not flagged this year" in patterns["Rent"]


def test_annotate_with_history_no_prior_years_at_all():
    detail = pd.DataFrame([{"account_code": "4000", "account_name": "Sales", "flag": True}])
    annotated = pl_variance.annotate_with_history(detail, {})
    assert annotated.iloc[0]["historical_pattern"] == "No prior-year history yet"


def test_snapshot_into_history_records_every_account_not_just_flagged():
    client = {}
    detail = pd.DataFrame([
        {"account_code": "4000", "account_name": "Sales", "current_year": -100000.0, "variance_amount": -5000.0, "variance_pct": -0.05, "flag": False},
        {"account_code": "6000", "account_name": "Marketing", "current_year": 25000.0, "variance_amount": 17000.0, "variance_pct": 2.125, "flag": True},
    ])
    pl_variance.snapshot_into_history(client, "2025-12-31", detail)
    assert set(client["pl_history"]["2025-12-31"].keys()) == {"4000|Sales", "6000|Marketing"}
    assert client["pl_history"]["2025-12-31"]["6000|Marketing"]["flagged"] is True
    assert client["pl_history"]["2025-12-31"]["4000|Sales"]["flagged"] is False


def test_snapshot_into_history_overwrites_same_year_end_rather_than_duplicating():
    client = {}
    detail1 = pd.DataFrame([{"account_code": "4000", "account_name": "Sales", "current_year": -100000.0, "variance_amount": 0.0, "variance_pct": 0.0, "flag": False}])
    detail2 = pd.DataFrame([{"account_code": "4000", "account_name": "Sales", "current_year": -110000.0, "variance_amount": -10000.0, "variance_pct": -0.1, "flag": True}])
    pl_variance.snapshot_into_history(client, "2025-12-31", detail1)
    pl_variance.snapshot_into_history(client, "2025-12-31", detail2)
    assert len(client["pl_history"]) == 1
    assert client["pl_history"]["2025-12-31"]["4000|Sales"]["amount"] == -110000.0


def test_build_client_report_produces_plain_english_comments_only_on_flagged_lines():
    import io

    import openpyxl

    cur = _pl([["4000", "Sales", "Turnover", -120000], ["6100", "Rent", "Overheads", 12000]])
    comp = _pl([["4000", "Sales", "Turnover", -100000], ["6100", "Rent", "Overheads", 11500]])
    result = pl_variance.pl_variance_analysis(cur, comp, None, materiality=500, variance_pct_threshold=0.1)
    result.detail = pl_variance.annotate_with_history(result.detail, {})

    xlsx_bytes = pl_variance.build_client_report("Acme Ltd", "Year ended 31 Dec 2025", result.detail)
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active
    rows = list(ws.iter_rows(min_row=5, max_row=6, values_only=True))
    by_account = {r[1]: r[6] for r in rows}
    assert by_account["Sales"] != ""
    assert "Decreased" in by_account["Sales"] or "decreased" in by_account["Sales"].lower()
    assert not by_account["Rent"]  # below threshold - no comment (openpyxl reads an empty-string cell back as None)
