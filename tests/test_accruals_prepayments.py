"""Unit tests for the Accruals & Prepayments consolidated schedule - see
app/accruals_prepayments.py: a real working paper's single side-by-side
table of every prepayment and accrual line, brought forward + movement =
carried forward, checked against the TB - not scattered across a
separate tab per account (which control_accounts.py's per-account
rollforwards already provide individually)."""
import pandas as pd

from app.accruals_prepayments import build_schedule


def _tb(rows):
    return pd.DataFrame(rows)


def _nom(date, code, name, debit=0.0, credit=0.0, description="", contact=""):
    return {"date": pd.Timestamp(date), "account_code": code, "account_name": name,
            "reference": "", "description": description, "contact": contact, "debit": debit, "credit": credit}


def test_finds_prepayment_typed_and_accrual_named_accounts_and_ties_out():
    tb_current = _tb([
        {"account_code": "620", "account_name": "PREPAYMENTS - INSURANCE", "account_type": "Prepayment", "balance": 2400.0},
        {"account_code": "800", "account_name": "ACCRUED EXPENSES", "account_type": "Current Liability", "balance": -1500.0},
    ])
    tb_comparative = _tb([
        {"account_code": "620", "account_name": "PREPAYMENTS - INSURANCE", "account_type": "Prepayment", "balance": 1200.0},
        {"account_code": "800", "account_name": "ACCRUED EXPENSES", "account_type": "Current Liability", "balance": -900.0},
    ])
    nominal = pd.DataFrame([
        _nom("2025-06-01", "620", "PREPAYMENTS - INSURANCE", debit=2400.0, credit=1200.0, contact="AXA"),
        _nom("2025-12-31", "800", "ACCRUED EXPENSES", credit=600.0),
    ])
    result = build_schedule(tb_current, tb_comparative, nominal)
    assert result.status == "ok"
    assert len(result.detail) == 2
    assert set(result.detail["Type"]) == {"Prepayment", "Accrual"}
    assert "£2,400.00 prepaid" in result.message
    assert "£1,500.00 accrued" in result.message


def test_excludes_accounts_already_covered_by_their_own_dedicated_schedule():
    # "Accrued Corporation Tax" contains "accrued" but is already the CT
    # schedule's own line - including it here would double-count it
    tb_current = _tb([
        {"account_code": "810", "account_name": "ACCRUED CORPORATION TAX", "account_type": "Current Liability", "balance": -8000.0},
    ])
    result = build_schedule(tb_current, pd.DataFrame(), pd.DataFrame())
    assert result.status == "n/a"


def test_flags_a_line_that_does_not_tie_to_the_tb():
    tb_current = _tb([{"account_code": "620", "account_name": "PREPAYMENTS", "account_type": "Prepayment", "balance": 5000.0}])
    tb_comparative = _tb([{"account_code": "620", "account_name": "PREPAYMENTS", "account_type": "Prepayment", "balance": 1200.0}])
    # movement of only 1000 posted, but the TB balance implies 3800 moved - a genuine gap
    nominal = pd.DataFrame([_nom("2025-06-01", "620", "PREPAYMENTS", debit=1000.0)])
    result = build_schedule(tb_current, tb_comparative, nominal)
    assert result.status == "review"
    assert result.detail.iloc[0]["Diff"] != 0


def test_extra_detail_lists_the_postings_behind_each_line():
    tb_current = _tb([{"account_code": "620", "account_name": "PREPAYMENTS", "account_type": "Prepayment", "balance": 2400.0}])
    nominal = pd.DataFrame([_nom("2025-06-01", "620", "PREPAYMENTS", debit=2400.0, contact="AXA", description="Annual premium")])
    result = build_schedule(tb_current, pd.DataFrame(), nominal)
    assert len(result.extra_detail) == 1
    assert result.extra_detail.iloc[0]["Contact"] == "AXA"
    assert result.extra_detail.iloc[0]["Type"] == "Prepayment"


def test_na_when_no_matching_accounts_found():
    tb_current = _tb([{"account_code": "1", "account_name": "SALES", "account_type": "Sales", "balance": -1000.0}])
    result = build_schedule(tb_current, pd.DataFrame(), pd.DataFrame())
    assert result.status == "n/a"


def test_na_without_trial_balance():
    assert build_schedule(None, None, None).status == "n/a"
    assert build_schedule(pd.DataFrame(), pd.DataFrame(), pd.DataFrame()).status == "n/a"
