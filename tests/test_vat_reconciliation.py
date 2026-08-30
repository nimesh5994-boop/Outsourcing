"""Unit tests for the VAT Reconciliation matching engine (General Ledger vs
Filed VAT Return, Box 1/Box 4 reconciled independently) - see
app/vat_reconciliation.py for the three-pass cascade this exercises."""
import pandas as pd

from app.vat_reconciliation import VatReconSettings, reconcile, reconcile_box


def _gl(rows):
    df = pd.DataFrame(rows)
    for col in ("date", "reference", "contact", "description", "net_amount", "vat_amount"):
        if col not in df.columns:
            df[col] = "" if col not in ("net_amount", "vat_amount") else 0.0
    return df


def _filed(rows):
    df = _gl(rows)
    if "source_file" not in df.columns:
        df["source_file"] = "return.xlsx"
    return df


def test_no_filed_detail_is_not_applicable():
    result = reconcile_box("Box 1 (Sales)", _gl([]), None, VatReconSettings())
    assert result.status == "n/a"
    assert "No filed VAT return detail" in result.message


def test_reference_match_ignores_amount_tolerance_but_reports_the_variance():
    """The whole point of matching by invoice number: a real invoice with
    the wrong VAT amount posted against it should still be recognised as
    the same transaction, with the variance surfaced - not silently split
    into two unrelated 'unmatched' rows on either side."""
    gl = _gl([{"date": pd.Timestamp("2025-01-20"), "reference": "INV-101", "contact": "Beta Ltd", "net_amount": 500.0, "vat_amount": 100.0}])
    filed = _filed([{"date": pd.Timestamp("2025-01-20"), "reference": "INV-101", "contact": "Beta Ltd", "net_amount": 500.0, "vat_amount": 105.0}])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings(tolerance=0.0))
    assert result.status == "review"
    assert result.extra_detail_label
    assert len(result.extra_detail) == 1
    row = result.extra_detail.iloc[0]
    assert row["Exception type"] == "Matched but VAT amount differs"
    assert "5.00" in row["Note"]


def test_exact_match_by_reference_is_ok():
    gl = _gl([{"date": pd.Timestamp("2025-01-15"), "reference": "INV-100", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0}])
    filed = _filed([{"date": pd.Timestamp("2025-01-15"), "reference": "INV-100", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0}])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings())
    assert result.status == "ok"
    assert result.extra_detail.empty


def test_no_reference_falls_back_to_date_amount_contact():
    gl = _gl([{"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.0}])
    filed = _filed([{"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.0}])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings())
    assert result.status == "ok"


def test_amount_outside_tolerance_with_no_reference_is_unmatched():
    """reconcile_box() (a single box in isolation) only ever reports its
    own filed side's exceptions - unmatched GL activity is reconcile()'s
    job (see test_gl_activity_neither_box_accounts_for_is_reported_once),
    since a lone box has no way to know whether an unclaimed GL row
    belongs to some other box it was never told about."""
    gl = _gl([{"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.0}])
    filed = _filed([{"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 65.0}])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings(tolerance=0.0))
    assert result.status == "review"
    assert set(result.extra_detail["Exception type"]) == {"Filed return item - no GL match"}


def test_tolerance_absorbs_small_amount_differences():
    gl = _gl([{"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.0}])
    filed = _filed([{"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.5}])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings(tolerance=1.0))
    assert result.status == "ok"


def test_each_gl_row_matched_at_most_once():
    """Two filed items that would both plausibly match the same single GL
    row (loose amount+contact pass) must not both claim it - the second
    one is left as a genuine unmatched exception."""
    gl = _gl([{"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.0}])
    filed = _filed([
        {"date": pd.Timestamp("2025-01-25"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.0},
        {"date": pd.Timestamp("2025-02-01"), "reference": "", "contact": "Gamma Ltd", "net_amount": 300.0, "vat_amount": 60.0},
    ])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings())
    assert result.status == "review"
    assert (result.detail.loc[result.detail["Measure"] == "Matched to General Ledger", "Count"] == 1).all()
    assert (result.detail.loc[result.detail["Measure"] == "Filed items with no GL match", "Count"] == 1).all()


def test_cash_basis_uses_payment_date_with_fallback_to_transaction_date():
    """Two GL rows share the same contact/amount as their filed
    counterparts but on different dates, so the loosest pass (amount +
    contact) would match either way regardless of basis - what the basis
    setting actually changes is *how confidently* it matches: cash basis
    resolves it via the stronger date+amount+contact pass because the
    payment date lines up, where accrual basis has to fall back to the
    weakest "verify" pass because the transaction date doesn't."""
    gl = pd.DataFrame([
        {"date": pd.Timestamp("2025-01-15"), "payment_date": pd.Timestamp("2025-02-10"), "reference": "", "contact": "Acme Ltd", "description": "", "net_amount": 1000.0, "vat_amount": 200.0},
    ])
    filed = _filed([
        {"date": pd.Timestamp("2025-02-10"), "reference": "", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0},
    ])
    from app.vat_reconciliation import match_box, _prepare_pool

    accrual_settings = VatReconSettings(accounting_basis="accrual")
    matched, _, _ = match_box(_prepare_pool(gl, accrual_settings), filed, accrual_settings)
    assert matched.iloc[0]["Match basis"] == "amount + contact (verify)"

    cash_settings = VatReconSettings(accounting_basis="cash")
    matched, _, _ = match_box(_prepare_pool(gl, cash_settings), filed, cash_settings)
    assert matched.iloc[0]["Match basis"] == "date + amount + contact"


def test_box1_and_box4_share_one_gl_pool_without_cross_contamination():
    """A purchase transaction must never show up as a Box 1 exception
    just because Box 1's filed sales never mention it (and vice versa) -
    reconcile() runs both boxes against a shared pool so each box only
    ever reports its own filed items, with anything neither box claims
    surfaced once as its own General Ledger Coverage check."""
    gl = _gl([
        {"date": pd.Timestamp("2025-01-15"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0},
        {"date": pd.Timestamp("2025-02-01"), "reference": "BILL-1", "contact": "Supplier Co", "net_amount": 800.0, "vat_amount": 160.0},
    ])
    filed_sales = _filed([{"date": pd.Timestamp("2025-01-15"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0}])
    filed_purchases = _filed([{"date": pd.Timestamp("2025-02-01"), "reference": "BILL-1", "contact": "Supplier Co", "net_amount": 800.0, "vat_amount": 160.0}])

    results = reconcile({"vat_gl": gl, "vat_filed_sales": filed_sales, "vat_filed_purchases": filed_purchases}, VatReconSettings())
    assert len(results) == 4
    assert [r.name for r in results] == [
        "VAT Recon - Box 1 (Sales)", "VAT Recon - Box 4 (Purchases)", "VAT Recon - General Ledger Coverage",
        "VAT Recon - suggested box for unmatched General Ledger items",
    ]
    # every GL row claimed by exactly one box - nothing left for the coverage-gap suggestion to look at
    assert [r.status for r in results] == ["ok", "ok", "ok", "n/a"]


def test_gl_activity_neither_box_accounts_for_is_reported_once():
    gl = _gl([
        {"date": pd.Timestamp("2025-01-15"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0},
        {"date": pd.Timestamp("2025-03-01"), "reference": "MYSTERY-1", "contact": "Nobody Ltd", "net_amount": 50.0, "vat_amount": 10.0},
    ])
    filed_sales = _filed([{"date": pd.Timestamp("2025-01-15"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0}])
    results = reconcile({"vat_gl": gl, "vat_filed_sales": filed_sales, "vat_filed_purchases": None}, VatReconSettings())
    coverage = results[2]
    assert coverage.status == "review"
    assert len(coverage.extra_detail) == 1
    assert coverage.extra_detail.iloc[0]["Reference"] == "MYSTERY-1"


def test_matched_detail_lists_every_pair_with_implied_vat_rate():
    gl = _gl([
        {"date": pd.Timestamp("2025-04-14"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0},
        {"date": pd.Timestamp("2025-04-20"), "reference": "INV-2", "contact": "Widgets Co", "net_amount": 500.0, "vat_amount": 87.5},  # 17.5%
    ])
    filed = _filed([
        {"date": pd.Timestamp("2025-04-14"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0},
        {"date": pd.Timestamp("2025-04-20"), "reference": "INV-2", "contact": "Widgets Co", "net_amount": 500.0, "vat_amount": 87.5},
    ])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings())
    assert result.status == "ok"  # the non-standard rate is advisory only, never flips status
    assert len(result.matched_detail) == 2
    rates = dict(zip(result.matched_detail["Filed reference"], result.matched_detail["Implied VAT rate %"]))
    assert rates["INV-1"] == 20.0
    assert rates["INV-2"] == 17.5
    assert "1 matched item(s)" in result.message
    assert "standard UK rate" in result.message


def test_matched_detail_absent_when_nothing_matched():
    gl = _gl([])
    filed = _filed([{"date": pd.Timestamp("2025-04-14"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 1000.0, "vat_amount": 200.0}])
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings())
    assert result.matched_detail.empty
    assert "standard UK rate" not in result.message


def test_gl_coverage_gap_suggests_box_from_contact_history():
    gl = _gl([
        {"date": pd.Timestamp("2025-01-01"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},
        {"date": pd.Timestamp("2025-02-01"), "reference": "INV-2", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},
        {"date": pd.Timestamp("2025-03-01"), "reference": "INV-3", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},
        {"date": pd.Timestamp("2025-04-01"), "reference": "INV-4", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},  # unmatched
    ])
    filed_sales = _filed([
        {"date": pd.Timestamp("2025-01-01"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},
        {"date": pd.Timestamp("2025-02-01"), "reference": "INV-2", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},
        {"date": pd.Timestamp("2025-03-01"), "reference": "INV-3", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},
    ])
    results = reconcile({"vat_gl": gl, "vat_filed_sales": filed_sales}, VatReconSettings())
    suggestion = results[3]
    assert suggestion.name == "VAT Recon - suggested box for unmatched General Ledger items"
    assert suggestion.status == "review"
    assert len(suggestion.detail) == 1
    assert suggestion.detail.iloc[0]["Suggested box"] == "Box 1 (Sales)"
    assert suggestion.detail.iloc[0]["Reference"] == "INV-4"


def test_gl_coverage_gap_no_suggestion_without_enough_history():
    gl = _gl([
        {"date": pd.Timestamp("2025-01-01"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0},
        {"date": pd.Timestamp("2025-02-01"), "reference": "INV-2", "contact": "New Contact Ltd", "net_amount": 50.0, "vat_amount": 10.0},
    ])
    filed_sales = _filed([{"date": pd.Timestamp("2025-01-01"), "reference": "INV-1", "contact": "Acme Ltd", "net_amount": 100.0, "vat_amount": 20.0}])
    results = reconcile({"vat_gl": gl, "vat_filed_sales": filed_sales}, VatReconSettings())
    suggestion = results[3]
    assert suggestion.status == "ok"


def test_gl_coverage_gap_suggestion_na_with_no_general_ledger():
    results = reconcile({"vat_filed_sales": _filed([])}, VatReconSettings())
    assert results[3].status == "n/a"


def test_reconcile_with_no_data_returns_four_not_applicable_results():
    results = reconcile({}, VatReconSettings())
    assert [r.status for r in results] == ["n/a", "n/a", "n/a", "n/a"]


def test_multi_file_source_tagging_survives_into_unmatched_rows():
    gl = _gl([])
    filed = pd.concat([
        _filed([{"date": pd.Timestamp("2025-01-01"), "reference": "A", "contact": "X", "net_amount": 100.0, "vat_amount": 20.0, "source_file": "q1.xlsx"}]),
        _filed([{"date": pd.Timestamp("2025-04-01"), "reference": "B", "contact": "Y", "net_amount": 200.0, "vat_amount": 40.0, "source_file": "q2.xlsx"}]),
    ], ignore_index=True)
    result = reconcile_box("Box 1 (Sales)", gl, filed, VatReconSettings())
    assert set(result.extra_detail["Source File"]) == {"q1.xlsx", "q2.xlsx"}
