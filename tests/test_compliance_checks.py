"""Tests for the data-driven compliance checks distilled from a real
manual-job review checklist: DLA monthly withdrawals + S455, dividends vs
distributable reserves, petty cash running balance, and loan/BBL/HP
facility detection."""
import pandas as pd
import pytest

from app import compliance_checks as cc


def _tb_row(code, name, account_type, debit, credit):
    return {"account_code": code, "account_name": name, "account_type": account_type,
            "debit": debit, "credit": credit, "balance": debit - credit}


def _nom_row(date, code, name, debit=0.0, credit=0.0, source_type="Journal"):
    return {"date": pd.Timestamp(date), "account_code": code, "account_name": name,
            "reference": "", "description": "", "contact": "", "source_type": source_type,
            "debit": debit, "credit": credit}


# ---------- Directors' Loan Account ----------

def test_dla_flags_monthly_withdrawal_over_threshold_and_s455():
    tb_current = pd.DataFrame([_tb_row("8060", "DIRECTORS' CURRENT ACCOUNT", "Current Liability", 5000, 0)])
    tb_comparative = pd.DataFrame([_tb_row("8060", "DIRECTORS' CURRENT ACCOUNT", "Current Liability", 0, 2000)])
    nominal = pd.DataFrame([
        _nom_row("2025-03-05", "8060", "DIRECTORS' CURRENT ACCOUNT", debit=12000),
        _nom_row("2025-06-05", "8060", "DIRECTORS' CURRENT ACCOUNT", debit=3000),
    ])
    result = cc.directors_loan_account_review(tb_current, tb_comparative, nominal)

    assert result.status == "review"
    month_rows = result.detail[result.detail["Month"] == "2025-03"]
    assert len(month_rows) == 1
    assert month_rows.iloc[0]["Net withdrawal £"] == pytest.approx(12000.0)
    assert "S455" in result.message


def test_dla_ok_when_under_threshold_and_in_credit():
    tb_current = pd.DataFrame([_tb_row("8060", "DIRECTORS' CURRENT ACCOUNT", "Current Liability", 0, 500)])
    tb_comparative = pd.DataFrame([_tb_row("8060", "DIRECTORS' CURRENT ACCOUNT", "Current Liability", 0, 300)])
    nominal = pd.DataFrame([_nom_row("2025-03-05", "8060", "DIRECTORS' CURRENT ACCOUNT", debit=2000, credit=2200)])
    result = cc.directors_loan_account_review(tb_current, tb_comparative, nominal)
    assert result.status == "ok"


def test_dla_n_a_when_no_such_account():
    tb_current = pd.DataFrame([_tb_row("8010", "TRADE CREDITORS", "Current Liability", 0, 5000)])
    result = cc.directors_loan_account_review(tb_current, None, None)
    assert result.status == "n/a"


# ---------- Dividends vs reserves ----------

def test_dividend_flags_unlawful_dividend():
    tb_current = pd.DataFrame([_tb_row("9800", "DIVIDENDS PAID", "Equity", 50000, 0)])
    tb_comparative = pd.DataFrame([_tb_row("960", "RETAINED EARNINGS", "Equity", 0, 30000)])
    nominal = pd.DataFrame([_nom_row("2025-06-01", "9800", "DIVIDENDS PAID", debit=50000)])
    result = cc.dividend_reserves_review(tb_current, tb_comparative, nominal, current_year_profit=10000)

    assert result.status == "review"
    assert "unlawful dividend" in result.message
    row = result.detail.iloc[0]
    assert row["Available distributable reserves"] == pytest.approx(40000.0)
    assert row["Dividends declared this year"] == pytest.approx(50000.0)


def test_dividend_ok_when_covered_by_reserves():
    tb_current = pd.DataFrame([_tb_row("9800", "DIVIDENDS PAID", "Equity", 10000, 0)])
    tb_comparative = pd.DataFrame([_tb_row("960", "RETAINED EARNINGS", "Equity", 0, 30000)])
    nominal = pd.DataFrame([_nom_row("2025-06-01", "9800", "DIVIDENDS PAID", debit=10000)])
    result = cc.dividend_reserves_review(tb_current, tb_comparative, nominal, current_year_profit=10000)
    assert result.status == "ok"


def test_dividend_n_a_when_no_dividend_account():
    tb_current = pd.DataFrame([_tb_row("8010", "TRADE CREDITORS", "Current Liability", 0, 5000)])
    result = cc.dividend_reserves_review(tb_current, None, None, current_year_profit=1000)
    assert result.status == "n/a"


# ---------- Petty cash ----------

def test_petty_cash_flags_negative_running_balance():
    tb_comparative = pd.DataFrame([_tb_row("1000", "PETTY CASH", "Current Asset", 500, 0)])
    nominal = pd.DataFrame([
        _nom_row("2025-01-05", "1000", "PETTY CASH", credit=800, source_type="Bank"),
        _nom_row("2025-02-05", "1000", "PETTY CASH", debit=400, source_type="Bank"),
    ])
    result = cc.petty_cash_running_balance_review(tb_comparative, nominal)

    assert result.status == "review"
    assert len(result.detail) == 1
    assert result.detail.iloc[0]["Running balance after this transaction"] == pytest.approx(-300.0)


def test_petty_cash_ok_when_never_negative():
    tb_comparative = pd.DataFrame([_tb_row("1000", "PETTY CASH", "Current Asset", 500, 0)])
    nominal = pd.DataFrame([_nom_row("2025-01-05", "1000", "PETTY CASH", debit=100, source_type="Bank")])
    result = cc.petty_cash_running_balance_review(tb_comparative, nominal)
    assert result.status == "ok"


def test_petty_cash_n_a_when_no_petty_cash_account():
    tb_comparative = pd.DataFrame([_tb_row("8010", "TRADE CREDITORS", "Current Liability", 0, 5000)])
    result = cc.petty_cash_running_balance_review(tb_comparative, None)
    assert result.status == "n/a"


# ---------- Loan facility review ----------

def test_loan_facility_detects_bbl_and_hp_with_reminders():
    tb_current = pd.DataFrame([
        _tb_row("2200", "BOUNCE BACK LOAN", "Liability", 0, 25000),
        _tb_row("2300", "HIRE PURCHASE CREDITOR", "Liability", 0, 8000),
    ])
    result = cc.loan_facility_review(tb_current, None)

    assert result.status == "review"
    assert set(result.detail["Facility type"]) == {"Bounce Back Loan", "Hire Purchase"}
    bbl_row = result.detail[result.detail["Facility type"] == "Bounce Back Loan"].iloc[0]
    assert "12 months" in bbl_row["Reminder"]
    # no nominal activity supplied - can't say whether it ties, but the
    # rollforward block still shows b/fwd and c/fwd per the TB
    assert bbl_row["Rollforward ties to TB?"] == "No nominal activity detail supplied"
    assert "No nominal activity detail available" in result.extra_detail["Reference"].values


def test_loan_facility_rollforward_ties_when_bfwd_plus_movement_matches_tb():
    tb_current = pd.DataFrame([_tb_row("2250", "BANK LOAN", "Non-current Liability", 0, 18000)])
    tb_comparative = pd.DataFrame([_tb_row("2250", "BANK LOAN", "Non-current Liability", 0, 20000)])
    nominal = pd.DataFrame([_nom_row("2025-06-01", "2250", "BANK LOAN", debit=2000)])  # a repayment
    result = cc.loan_facility_review(tb_current, tb_comparative, nominal)

    row = result.detail.iloc[0]
    assert row["Rollforward ties to TB?"] == "Yes"
    assert "every one ties" in result.message
    assert not result.extra_detail.empty
    assert "BANK LOAN (2250)" in result.extra_detail["Item"].iloc[0]


def test_loan_facility_rollforward_flags_when_it_does_not_tie():
    tb_current = pd.DataFrame([_tb_row("2250", "BANK LOAN", "Non-current Liability", 0, 18000)])
    tb_comparative = pd.DataFrame([_tb_row("2250", "BANK LOAN", "Non-current Liability", 0, 20000)])
    # a repayment of 500 posted, but the TB moved by 2000 - a 1500 gap
    nominal = pd.DataFrame([_nom_row("2025-06-01", "2250", "BANK LOAN", debit=500)])
    result = cc.loan_facility_review(tb_current, tb_comparative, nominal)

    row = result.detail.iloc[0]
    assert row["Rollforward ties to TB?"] == "No - see rollforward below"
    assert "don't tie" in result.message
    assert not result.extra_detail.empty


def test_loan_facility_does_not_false_positive_on_unrelated_accounts():
    # regression check for the \bhp\b word-boundary regex: words containing
    # "hp"-adjacent letters shouldn't false-positive
    tb_current = pd.DataFrame([
        _tb_row("3800", "INSURANCES (NOT PREMISES)", "Overhead", 1200, 0),
        _tb_row("2860", "OTHER OFFICE SUPPLIES", "Overhead", 300, 0),
    ])
    result = cc.loan_facility_review(tb_current, None)
    assert result.status == "n/a"


def test_loan_facility_n_a_when_none_found():
    tb_current = pd.DataFrame([_tb_row("8010", "TRADE CREDITORS", "Current Liability", 0, 5000)])
    result = cc.loan_facility_review(tb_current, None)
    assert result.status == "n/a"


# ---------- Stock/inventory review ----------

def test_stock_review_detects_inventory_type_account_with_movement_and_reminder():
    tb_current = pd.DataFrame([_tb_row("1300", "FINISHED GOODS STOCK", "Inventory", 15000, 0)])
    tb_comparative = pd.DataFrame([_tb_row("1300", "FINISHED GOODS STOCK", "Inventory", 12000, 0)])
    result = cc.stock_review(tb_current, tb_comparative)

    assert result.status == "review"
    row = result.detail.iloc[0]
    assert row["Current year closing"] == 15000.0
    assert row["Comparative year closing"] == 12000.0
    assert row["Movement"] == 3000.0
    assert "physical stock count" in row["Reminder"]
    # no nominal activity for stock is the normal case, not a red flag
    assert row["Rollforward ties to TB?"] == "No nominal activity detail (normal for stock)"


def test_stock_review_detects_by_name_even_when_account_type_is_not_inventory():
    # not every chart of accounts uses Xero's "Inventory" type correctly
    tb_current = pd.DataFrame([_tb_row("1310", "WORK IN PROGRESS", "Current Asset", 4000, 0)])
    result = cc.stock_review(tb_current, None)
    assert result.status == "review"
    assert result.detail.iloc[0]["Account name"] == "WORK IN PROGRESS"


def test_stock_review_n_a_when_none_found():
    tb_current = pd.DataFrame([_tb_row("8010", "TRADE CREDITORS", "Current Liability", 0, 5000)])
    result = cc.stock_review(tb_current, None)
    assert result.status == "n/a"


def test_run_all_compliance_checks_returns_five_results_with_no_data():
    results = cc.run_all_compliance_checks(pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), None)
    assert len(results) == 5
    assert all(r.status == "n/a" for r in results)


def test_compliance_checks_run_cleanly_against_real_shaped_sample_data(canonical_data):
    # Brightwell (the fictional sample client) doesn't have DLA/dividend/
    # petty cash/loan accounts, so those four should gracefully report n/a
    # rather than erroring or false-positiving. It does carry a genuine
    # stock account, so that one is expected to come back "review" - by
    # design, stock can never reach "ok" on its own (see stock_review),
    # since a physical count is outside any data this system ingests.
    results = cc.run_all_compliance_checks(
        canonical_data["tb_current"], canonical_data["tb_comparative"], canonical_data["nominal_current"],
        current_year_profit=float(canonical_data["pl_current"]["amount"].sum()),
    )
    assert len(results) == 5
    other_checks = [r for r in results if r.name != "Stock/inventory review"]
    assert all(r.status in ("ok", "n/a") for r in other_checks)
    stock_result = next(r for r in results if r.name == "Stock/inventory review")
    assert stock_result.status == "review"
