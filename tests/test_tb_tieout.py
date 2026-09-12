"""Unit tests for the Trial Balance Tie-Out - see app/tb_tieout.py.

Covers the two independent checks it runs: opening + this year's
nominal movement = reported closing (current year only - the
comparative year is already filed and finalised, so it isn't
re-verified), and a cross-check between two Trial Balance uploads'
opening figures for the same account.
"""
import pandas as pd

from app.tb_tieout import build_tieout


def _tb(rows):
    df = pd.DataFrame(rows)
    df["balance"] = df["debit"] - df["credit"]
    return df


def _nominal(rows):
    return pd.DataFrame(rows)


def test_na_without_a_current_trial_balance():
    result = build_tieout(None, None, None)
    assert result.status == "n/a"


def test_ok_when_every_account_ties_out():
    tb_comparative = _tb([
        {"account_code": "200", "account_name": "Sales", "account_type": "Revenue", "debit": 0, "credit": 1000},
        {"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500},
    ])
    tb_current = _tb([
        {"account_code": "200", "account_name": "Sales", "account_type": "Revenue", "debit": 0, "credit": 300},
        {"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 700},
    ])
    nominal_current = _nominal([
        {"account_code": "200", "debit": 0, "credit": 300},
        {"account_code": "610", "debit": 0, "credit": 200},
    ])
    result = build_tieout(tb_current, tb_comparative, nominal_current)
    assert result.status == "ok"
    assert (result.detail["Flag"] == "OK").all()
    # Sales is a P&L code - opening must be 0, not last year's 1000 closing
    sales_row = result.detail[result.detail["Account Code"] == "200"].iloc[0]
    assert sales_row["Opening (per comparative TB)"] == 0.0


def test_flags_an_account_that_does_not_tie_for_the_current_year():
    tb_comparative = _tb([
        {"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500},
    ])
    tb_current = _tb([
        {"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 700},
    ])
    # movement of 100 posted, but TB shows a 200 increase - a 100 gap
    nominal_current = _nominal([
        {"account_code": "610", "debit": 0, "credit": 100},
    ])
    result = build_tieout(tb_current, tb_comparative, nominal_current)
    assert result.status == "review"
    row = result.detail.iloc[0]
    assert row["Flag"] == "REVIEW"
    # opening -500 + movement -100 = derived -600; reported is -700 (700 credit) -
    # derived minus reported = 100
    assert row["Diff"] == 100.0


def test_dormant_account_with_no_postings_still_checked_not_skipped():
    tb_comparative = _tb([
        {"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500},
    ])
    tb_current = _tb([
        {"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500},
    ])
    # nominal activity uploaded for the year, but nothing posted to this code
    nominal_current = _nominal([
        {"account_code": "999", "debit": 10, "credit": 0},
    ])
    result = build_tieout(tb_current, tb_comparative, nominal_current)
    assert result.status == "ok"
    assert result.detail.iloc[0]["Flag"] == "OK"
    assert result.detail.iloc[0]["Movement (current year)"] == 0.0


def test_na_flag_when_no_nominal_activity_uploaded_at_all():
    tb_comparative = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500}])
    tb_current = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 700}])
    result = build_tieout(tb_current, tb_comparative, None)
    assert result.status == "ok"  # nothing to flag - can't verify, not a failure
    assert result.detail.iloc[0]["Flag"] == "n/a"


def test_opening_balance_disagreement_between_two_tb_uploads_is_flagged():
    tb_current = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 700}])
    tb_current_own_comparative = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500}])
    tb_comparative = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 480}])

    result = build_tieout(tb_current, tb_comparative, None, tb_current_own_comparative=tb_current_own_comparative)
    assert not result.matched_detail.empty
    row = result.matched_detail.iloc[0]
    assert row["Diff"] == -20.0
    assert result.status == "review"


def test_no_disagreement_reported_when_the_two_tb_uploads_agree():
    tb_current = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 700}])
    tb_current_own_comparative = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500}])
    tb_comparative = _tb([{"account_code": "610", "account_name": "Trade Creditors", "account_type": "Current Liability", "debit": 0, "credit": 500}])

    result = build_tieout(tb_current, tb_comparative, None, tb_current_own_comparative=tb_current_own_comparative)
    assert result.matched_detail.empty
