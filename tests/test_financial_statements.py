"""Unit tests for P&L/Balance Sheet derivation and presentation - see
app/xero_reports.py's derive_pl_bs_from_tb and app/financial_statements.py.

Regression coverage for a real bug found via a live Xero export: Xero's own
standard Account Type label for a sales account is "Revenue" (the account
*name* is often "Sales", but the *type* is "Revenue") - derive_pl_bs_from_tb
didn't recognise it, so those accounts silently fell through to the B/S
side instead of the P&L, understating Turnover to zero and throwing the
B/S balance check off by the same amount.
"""
import pandas as pd

from app.financial_statements import build_pl_statement
from app.xero_reports import derive_pl_bs_from_tb


def _tb(rows):
    return pd.DataFrame(rows)


def test_derive_pl_bs_routes_revenue_account_type_to_pl_not_bs():
    tb = _tb([
        {"account_code": "200", "account_name": "Sales (Flat)", "account_type": "Revenue", "debit": 0, "credit": 18646.19},
        {"account_code": "400", "account_name": "Advertising", "account_type": "Overhead", "debit": 133.33, "credit": 0},
        {"account_code": "710", "account_name": "Office Equipment", "account_type": "Fixed Asset", "debit": 500.0, "credit": 0},
    ])
    pl, bs = derive_pl_bs_from_tb(tb)

    assert "Sales (Flat)" in pl["account_name"].values
    assert "Sales (Flat)" not in bs["account_name"].values
    assert pl.loc[pl["account_name"] == "Sales (Flat)", "category"].iloc[0] == "Revenue"


def test_pl_statement_counts_revenue_category_as_turnover():
    pl = _tb([
        {"account_code": "200", "account_name": "Sales (Flat)", "category": "Revenue", "amount": 18646.19},
        {"account_code": "400", "account_name": "Advertising", "category": "Overhead", "amount": -133.33},
    ])
    result = build_pl_statement(pl)
    turnover = result.statement.set_index("Line").loc["Turnover", "Amount"]
    assert turnover == 18646.19


def test_derive_pl_bs_is_case_insensitive_on_account_type():
    tb = _tb([
        {"account_code": "200", "account_name": "Sales (Flat)", "account_type": "revenue", "debit": 0, "credit": 100.0},
    ])
    pl, _bs = derive_pl_bs_from_tb(tb)
    assert "Sales (Flat)" in pl["account_name"].values
