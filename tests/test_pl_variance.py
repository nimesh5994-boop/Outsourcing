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


def test_pl_notes_detail_lists_every_contact_under_every_code_not_just_flagged():
    """The supplier/customer drill-down (result.matched_detail) is a full
    P&L note, not an exceptions list - a code that never breached the
    variance threshold still gets its own contact breakdown + TOTAL row."""
    cur = _pl([["6000", "Marketing", "Overheads", 25000], ["6300", "Postage", "Overheads", 1000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000], ["6300", "Postage", "Overheads", 980]])
    nominal = _gl([
        ["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 25000, 0],
        ["6300", "Postage", "2025-06-01", "INV4", "stamps", "Solo Ltd", 1000, 0],
    ])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    notes = result.matched_detail
    assert set(notes["Account"]) == {"6000 - Marketing", "6300 - Postage"}
    # Postage (6300) is unflagged (£20/2% - below threshold) but still appears.
    postage_rows = notes[notes["Account"] == "6300 - Postage"]
    assert set(postage_rows["Contact"]) == {"Solo Ltd", "TOTAL"}
    postage_total = postage_rows[postage_rows["Contact"] == "TOTAL"].iloc[0]
    assert postage_total["Main Variance Driver"] == ""  # not flagged - no driver text


def test_pl_notes_detail_identifies_main_variance_driver_on_flagged_code():
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal_cur = _gl([
        ["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 20000, 0],
        ["6000", "Marketing", "2025-04-01", "INV2", "socials", "Small Co", 5000, 0],
    ])
    nominal_comp = _gl([
        ["6000", "Marketing", "2024-03-01", "INV1", "ads", "Acme Agency", 6000, 0],
        ["6000", "Marketing", "2024-04-01", "INV2", "socials", "Small Co", 2000, 0],
    ])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal_cur, nominal_comp, materiality=500, variance_pct_threshold=0.1)
    total_row = result.matched_detail[result.matched_detail["Contact"] == "TOTAL"].iloc[0]
    # Acme Agency moved +14,000, Small Co only +3,000 - Acme is the main driver.
    assert "Acme Agency" in total_row["Main Variance Driver"]
    assert "+£14,000.00" in total_row["Main Variance Driver"]
    assert total_row["Main Variance Driver"].startswith("Higher")


def test_pl_notes_detail_main_variance_driver_can_be_a_decrease():
    cur = _pl([["6000", "Marketing", "Overheads", 2000]])
    comp = _pl([["6000", "Marketing", "Overheads", 20000]])
    nominal_cur = _gl([["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 2000, 0]])
    nominal_comp = _gl([["6000", "Marketing", "2024-03-01", "INV1", "ads", "Acme Agency", 20000, 0]])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal_cur, nominal_comp, materiality=500, variance_pct_threshold=0.1)
    total_row = result.matched_detail[result.matched_detail["Contact"] == "TOTAL"].iloc[0]
    assert total_row["Main Variance Driver"].startswith("Lower")
    assert "-£18,000.00" in total_row["Main Variance Driver"]


def test_pl_notes_detail_duplicate_name_flagged_across_any_two_codes_even_when_neither_is_flagged():
    """Unlike _split_across_codes (only checks contacts tied to an
    ALREADY-flagged account), the notes drill-down's duplicate-name check
    runs across every P&L code - a contact double-coded between two
    perfectly quiet accounts still gets caught."""
    cur = _pl([["6000", "Marketing", "Overheads", 5010], ["6100", "Rent", "Overheads", 5010]])
    comp = _pl([["6000", "Marketing", "Overheads", 5000], ["6100", "Rent", "Overheads", 5000]])
    nominal = _gl([
        ["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 5010, 0],
        ["6100", "Rent", "2025-04-01", "INV2", "office", "Acme Agency", 5010, 0],
        ["6000", "Marketing", "2025-05-01", "INV3", "stamps", "Solo Ltd", 0, 0],
    ])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    assert result.status == "ok"  # neither code breached materiality/threshold
    assert result.extra_detail.empty  # _split_across_codes has nothing flagged to key off
    notes = result.matched_detail
    acme_rows = notes[notes["Contact"] == "Acme Agency"]
    assert (acme_rows["Duplicate Name"] != "").all()
    assert set(acme_rows["Account"]) == {"6000 - Marketing", "6100 - Rent"}


def test_pl_notes_detail_no_duplicate_flag_for_contact_on_only_one_code():
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal = _gl([["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 25000, 0]])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    contact_row = result.matched_detail[result.matched_detail["Contact"] == "Acme Agency"].iloc[0]
    assert contact_row["Duplicate Name"] == ""


def test_pl_notes_detail_without_nominal_comparative_shows_zero_previous_year_per_contact():
    """No prior-year nominal upload means no per-contact comparative figure
    is possible - the TOTAL row still shows the real comparative_year
    figure from pl_comparative, but every contact's own 'Previous Year' is
    0, so the gap is visible rather than silently guessed at."""
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal = _gl([["6000", "Marketing", "2025-03-01", "INV1", "ads", "Acme Agency", 25000, 0]])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    contact_row = result.matched_detail[result.matched_detail["Contact"] == "Acme Agency"].iloc[0]
    total_row = result.matched_detail[result.matched_detail["Contact"] == "TOTAL"].iloc[0]
    assert contact_row["Previous Year"] == 0.0
    assert total_row["Previous Year"] == 8000.0


def test_extract_description_prefers_text_after_first_dash():
    assert pl_variance._extract_description("INV1023 - office repairs", "Acme Ltd") == "office repairs"


def test_extract_description_strips_contact_name_when_no_dash():
    assert pl_variance._extract_description("Acme Ltd: office repairs", "Acme Ltd") == "office repairs"


def test_extract_description_returns_as_is_when_no_dash_and_no_contact_match():
    assert pl_variance._extract_description("office repairs", "Someone Else Ltd") == "office repairs"


def test_extract_description_blank_input_returns_blank():
    assert pl_variance._extract_description("", "Acme Ltd") == ""
    assert pl_variance._extract_description("   ", "Acme Ltd") == ""


def test_pl_notes_detail_shows_deduplicated_current_and_previous_descriptions():
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal_cur = _gl([
        ["6000", "Marketing", "2025-03-01", "INV1", "INV1 - spring campaign", "Acme Agency", 15000, 0],
        ["6000", "Marketing", "2025-05-01", "INV3", "INV3 - spring campaign", "Acme Agency", 10000, 0],  # duplicate narrative, not repeated
    ])
    nominal_comp = _gl([
        ["6000", "Marketing", "2024-03-01", "INV0", "INV0 - winter campaign", "Acme Agency", 8000, 0],
    ])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal_cur, nominal_comp, materiality=500, variance_pct_threshold=0.1)
    contact_row = result.matched_detail[result.matched_detail["Contact"] == "Acme Agency"].iloc[0]
    assert contact_row["Current Period Description"] == "spring campaign"  # deduplicated, not "spring campaign; spring campaign"
    assert contact_row["Previous Period Description"] == "winter campaign"


def test_pl_notes_detail_joins_distinct_descriptions_with_semicolon():
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal = _gl([
        ["6000", "Marketing", "2025-03-01", "INV1", "INV1 - spring campaign", "Acme Agency", 15000, 0],
        ["6000", "Marketing", "2025-05-01", "INV3", "INV3 - print ads", "Acme Agency", 10000, 0],
    ])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    contact_row = result.matched_detail[result.matched_detail["Contact"] == "Acme Agency"].iloc[0]
    assert contact_row["Current Period Description"] == "spring campaign; print ads"


def test_pl_notes_detail_total_row_has_blank_descriptions():
    cur = _pl([["6000", "Marketing", "Overheads", 25000]])
    comp = _pl([["6000", "Marketing", "Overheads", 8000]])
    nominal = _gl([["6000", "Marketing", "2025-03-01", "INV1", "INV1 - spring campaign", "Acme Agency", 25000, 0]])
    result = pl_variance.pl_variance_analysis(cur, comp, nominal, materiality=500, variance_pct_threshold=0.1)
    total_row = result.matched_detail[result.matched_detail["Contact"] == "TOTAL"].iloc[0]
    assert total_row["Current Period Description"] == ""
    assert total_row["Previous Period Description"] == ""


def test_pl_notes_detail_empty_when_no_nominal_current():
    cur = _pl([["4000", "Sales", "Turnover", -120000]])
    comp = _pl([["4000", "Sales", "Turnover", -100000]])
    result = pl_variance.pl_variance_analysis(cur, comp, None, materiality=500, variance_pct_threshold=0.1)
    assert result.matched_detail.empty
    assert result.matched_detail_label == ""


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
