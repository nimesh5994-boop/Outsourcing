"""Unit tests for the Going Concern indicators check - see
app/going_concern.py. Works from the already-derived Balance Sheet
statement (financial_statements.build_bs_statement), so these pin down
the sign-convention handling specifically: the statement's own
"Current liabilities" line is displayed flipped to a positive number
(same treatment as debtors/creditors elsewhere), NOT negative the way
"NET ASSETS" is genuinely signed - a real bug found and fixed while
building this feature, since getting this backwards silently inverts
the whole analysis (a working-capital deficit would read as a surplus)."""
import pandas as pd

from app.financial_statements import build_bs_statement
from app.going_concern import assess


def _bs(rows):
    return pd.DataFrame(rows)


def test_healthy_balance_sheet_is_ok():
    bs = _bs([
        {"account_code": "1", "account_name": "Fixed Assets", "category": "Fixed Asset", "amount": 10000.0},
        {"account_code": "2", "account_name": "Debtors", "category": "Current Asset", "amount": 20000.0},
        {"account_code": "3", "account_name": "Creditors", "category": "Current Liability", "amount": -8000.0},
        {"account_code": "4", "account_name": "Share Capital", "category": "Equity", "amount": -100.0},
        {"account_code": "5", "account_name": "Retained Earnings", "category": "Equity", "amount": -21900.0},
    ])
    statement = build_bs_statement(bs, net_profit=0.0).statement
    result = assess(statement)
    assert result.status == "ok"


def test_negative_net_current_assets_and_negative_net_assets_both_flagged():
    # current liabilities (20000) exceed current assets (1000) - a working
    # capital deficit - AND total liabilities exceed total assets overall
    bs = _bs([
        {"account_code": "1", "account_name": "Fixed Assets", "category": "Fixed Asset", "amount": 5000.0},
        {"account_code": "2", "account_name": "Bank", "category": "Bank", "amount": 1000.0},
        {"account_code": "3", "account_name": "Creditors", "category": "Current Liability", "amount": -20000.0},
        {"account_code": "4", "account_name": "Share Capital", "category": "Equity", "amount": -100.0},
        {"account_code": "5", "account_name": "Retained Earnings", "category": "Equity", "amount": 14100.0},
    ])
    statement = build_bs_statement(bs, net_profit=0.0).statement
    result = assess(statement)
    assert result.status == "review"
    assert "working-capital deficit" in result.message
    assert "balance sheet insolvency" in result.message
    row = result.detail.set_index("Indicator")
    assert row.loc["Net current assets / (liabilities)", "Amount"] == -19000.0
    assert row.loc["Net assets / (liabilities)", "Amount"] == -14000.0


def test_working_capital_deficit_alone_does_not_imply_negative_net_assets():
    # current liabilities exceed current assets, but there's enough fixed
    # asset value that net assets overall stay positive - only one flag
    bs = _bs([
        {"account_code": "1", "account_name": "Fixed Assets", "category": "Fixed Asset", "amount": 50000.0},
        {"account_code": "2", "account_name": "Bank", "category": "Bank", "amount": 1000.0},
        {"account_code": "3", "account_name": "Creditors", "category": "Current Liability", "amount": -20000.0},
        {"account_code": "4", "account_name": "Share Capital", "category": "Equity", "amount": -100.0},
        {"account_code": "5", "account_name": "Retained Earnings", "category": "Equity", "amount": -30900.0},
    ])
    statement = build_bs_statement(bs, net_profit=0.0).statement
    result = assess(statement)
    assert result.status == "review"
    assert "working-capital deficit" in result.message
    assert "balance sheet insolvency" not in result.message
    row = result.detail.set_index("Indicator")
    assert row.loc["Net assets / (liabilities)", "Amount"] > 0


def test_message_explicitly_disclaims_being_a_verdict():
    bs = _bs([
        {"account_code": "1", "account_name": "Bank", "category": "Bank", "amount": 1000.0},
        {"account_code": "3", "account_name": "Creditors", "category": "Current Liability", "amount": -20000.0},
        {"account_code": "4", "account_name": "Retained Earnings", "category": "Equity", "amount": 19000.0},
    ])
    statement = build_bs_statement(bs, net_profit=0.0).statement
    result = assess(statement)
    assert "not a verdict" in result.message
    assert "judgement" in result.message


def test_na_without_balance_sheet_data():
    assert assess(None).status == "n/a"
    assert assess(pd.DataFrame()).status == "n/a"
