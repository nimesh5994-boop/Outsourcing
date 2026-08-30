"""Unit tests for the PAYE Reconciliation matching engine - see
app/paye_reconciliation.py. Net Pay is matched employee-by-employee-by-
month (fuzzy contact name match, since BrightPay's full legal name and
the General Ledger's contact don't always agree word-for-word); HMRC
PAYE & NI and Pension Contributions are matched as monthly company-wide
totals against a keyword-matched contact - real payroll bookkeeping
settles those as one lump payment a month, not per employee."""
import pandas as pd

from app.paye_reconciliation import PayeReconSettings, reconcile, reconcile_hmrc, reconcile_net_pay, reconcile_pension


def _gl(rows):
    df = pd.DataFrame(rows)
    for col in ("date", "reference", "contact", "description", "debit", "credit"):
        if col not in df.columns:
            df[col] = "" if col not in ("debit", "credit") else 0.0
    return df


def test_no_payroll_summary_is_not_applicable():
    result = reconcile_net_pay(None, _gl([]), PayeReconSettings())
    assert result.status == "n/a"


def test_net_pay_matches_by_fuzzy_employee_name():
    """BrightPay's "Jamie Francis Doe" (full legal name) vs the General
    Ledger's "Jamie Doe" (as a bookkeeper typed it) - a real mismatch
    found against an actual client export, not a hypothetical one - must
    still match on first+last name."""
    gl = _gl([{"date": pd.Timestamp("2025-04-30"), "contact": "Jamie Doe", "debit": 1084.84, "credit": 0.0}])
    payroll = pd.DataFrame([{"period_end": pd.Timestamp("2025-04-30"), "employee": "Jamie Francis Doe", "net_pay": 1084.84}])
    result = reconcile_net_pay(payroll, gl, PayeReconSettings())
    assert result.status == "ok"


def test_double_entry_legs_are_deduplicated_not_double_counted():
    """A Xero-style Account Transactions export lists one real payment
    twice - once per side of the double entry, same contact, equal and
    opposite debit/credit. Without deduplication this would either wrongly
    consume two GL rows for one BrightPay item, or leave a phantom second
    candidate sitting in the pool to falsely match some other month."""
    gl = _gl([
        {"date": pd.Timestamp("2025-04-30"), "contact": "Jamie Smith", "debit": 0.0, "credit": 1700.0},
        {"date": pd.Timestamp("2025-04-30"), "contact": "Jamie Smith", "debit": 1700.0, "credit": 0.0},
    ])
    payroll = pd.DataFrame([
        {"period_end": pd.Timestamp("2025-04-30"), "employee": "Jamie Smith", "net_pay": 1700.0},
        {"period_end": pd.Timestamp("2025-05-31"), "employee": "Jamie Smith", "net_pay": 1700.0},
    ])
    result = reconcile_net_pay(payroll, gl, PayeReconSettings())
    assert result.status == "review"  # only one real payment exists - the second month is a genuine miss
    assert (result.detail.loc[result.detail["Measure"] == "Matched to General Ledger", "Count"] == 1).all()
    assert (result.detail.loc[result.detail["Measure"] == "With no General Ledger match", "Count"] == 1).all()


def test_net_pay_skips_zero_pay_months():
    gl = _gl([])
    payroll = pd.DataFrame([{"period_end": pd.Timestamp("2025-12-31"), "employee": "Alex Doe", "net_pay": 0.0}])
    result = reconcile_net_pay(payroll, gl, PayeReconSettings())
    assert result.status == "n/a"


def test_unmatched_item_reports_closest_same_contact_candidate():
    gl = _gl([{"date": pd.Timestamp("2025-04-06"), "contact": "Jamie Doe", "debit": 980.0, "credit": 0.0}])
    payroll = pd.DataFrame([{"period_end": pd.Timestamp("2025-04-30"), "employee": "Jamie Francis Doe", "net_pay": 1084.84}])
    result = reconcile_net_pay(payroll, gl, PayeReconSettings(tolerance=0.0))
    assert result.status == "review"
    row = result.extra_detail.iloc[0]
    assert "980.00" in row["Closest GL candidate"]


def test_hmrc_matches_by_keyword_contact_and_lump_sum_amount():
    gl = _gl([{"date": pd.Timestamp("2025-05-20"), "contact": "HMRC", "debit": 449.52, "credit": 0.0}])
    p32 = pd.DataFrame([{"period_end": pd.Timestamp("2025-05-05"), "amount_due": 449.52}])
    result = reconcile_hmrc(p32, gl, PayeReconSettings())
    assert result.status == "ok"


def test_hmrc_respects_date_window():
    gl = _gl([{"date": pd.Timestamp("2025-08-01"), "contact": "HMRC", "debit": 449.52, "credit": 0.0}])
    p32 = pd.DataFrame([{"period_end": pd.Timestamp("2025-05-05"), "amount_due": 449.52}])
    result = reconcile_hmrc(p32, gl, PayeReconSettings(date_window_days=30))
    assert result.status == "review"


def test_pension_aggregates_across_employees_before_matching():
    """Pension is paid to the provider as one contribution covering every
    employee for the month - the per-employee BrightPay rows are summed
    before matching, not matched one at a time."""
    gl = _gl([{"date": pd.Timestamp("2025-04-21"), "contact": "Nest", "debit": 145.60, "credit": 0.0}])
    pensions = pd.DataFrame([
        {"period_end": pd.Timestamp("2025-04-30"), "employee": "Alex Doe", "total_pension": 145.60},
        {"period_end": pd.Timestamp("2025-04-30"), "employee": "Jamie Doe", "total_pension": 0.0},
    ])
    result = reconcile_pension(pensions, gl, PayeReconSettings())
    assert result.status == "ok"


def test_tolerance_absorbs_small_amount_differences():
    gl = _gl([{"date": pd.Timestamp("2025-04-21"), "contact": "Nest", "debit": 146.00, "credit": 0.0}])
    pensions = pd.DataFrame([{"period_end": pd.Timestamp("2025-04-30"), "employee": "Alex Doe", "total_pension": 145.60}])
    assert reconcile_pension(pensions, gl, PayeReconSettings(tolerance=0.0)).status == "review"
    assert reconcile_pension(pensions, gl, PayeReconSettings(tolerance=1.0)).status == "ok"


def test_reconcile_returns_three_results_in_order():
    results = reconcile({}, PayeReconSettings())
    assert [r.name for r in results] == [
        "PAYE Recon - Net Pay by Employee", "PAYE Recon - HMRC PAYE & NI", "PAYE Recon - Pension Contributions",
    ]
    assert [r.status for r in results] == ["n/a", "n/a", "n/a"]
