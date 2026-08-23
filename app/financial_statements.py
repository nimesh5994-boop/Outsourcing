"""Structured P&L and Balance Sheet presentation, with an explicit balance
check on the B/S rather than leaving it to the reader to work out.

A B/S built directly from a TB's own account categorisation will NOT sum to
zero on its own: the current year's profit sits inside the P&L accounts
until a formal closing journal (Dr P&L accounts / Cr Retained Earnings) is
posted, which most bookkeeping software doesn't do mid-year. So "assets -
liabilities - equity" is only ever zero once the current year's profit is
bridged in - that's not a bug, it's what "the accounts aren't closed yet"
looks like. The check below does that bridge explicitly and flags whatever
residual is left, which is either genuinely zero or a real problem (an
account miscategorised between P&L and B/S, a movement missing from the
TB/nominal activity supplied, etc.) rather than an unexplained gap.
"""
from dataclasses import dataclass, field

import pandas as pd

MATERIALITY_AMOUNT = 500.0

_PL_TURNOVER = {"sales", "income", "other income", "turnover"}
_PL_DIRECT_COSTS = {"direct costs", "direct cost", "cost of sales"}
_BS_FIXED_ASSETS = {"fixed asset"}
_BS_CURRENT_ASSETS = {"current asset", "bank"}
_BS_CURRENT_LIABILITIES = {"current liability"}
_BS_LONG_TERM_LIABILITIES = {"liability"}
_BS_EQUITY = {"equity"}


@dataclass
class StatementResult:
    status: str  # "ok" | "review" | "n/a"
    message: str
    statement: pd.DataFrame = field(default_factory=pd.DataFrame)
    net_profit: float = 0.0


def build_pl_statement(pl: pd.DataFrame) -> StatementResult:
    if pl is None or pl.empty:
        return StatementResult("n/a", "No P&L data available.")

    cat = pl["category"].str.lower()
    turnover = float(pl.loc[cat.isin(_PL_TURNOVER), "amount"].sum())
    direct_costs = float(pl.loc[cat.isin(_PL_DIRECT_COSTS), "amount"].sum())
    gross_profit = turnover + direct_costs
    overheads = float(pl.loc[~cat.isin(_PL_TURNOVER | _PL_DIRECT_COSTS), "amount"].sum())
    net_profit = gross_profit + overheads

    rows = [
        {"Line": "Turnover", "Amount": round(turnover, 2)},
        {"Line": "Direct costs", "Amount": round(direct_costs, 2)},
        {"Line": "GROSS PROFIT", "Amount": round(gross_profit, 2)},
        {"Line": "Overheads & other expenses", "Amount": round(overheads, 2)},
        {"Line": "NET PROFIT / (LOSS) FOR THE YEAR", "Amount": round(net_profit, 2)},
    ]
    return StatementResult("ok", "Profit & Loss account, grouped by category.", pd.DataFrame(rows), net_profit)


def build_bs_statement(bs: pd.DataFrame, net_profit: float) -> StatementResult:
    if bs is None or bs.empty:
        return StatementResult("n/a", "No Balance Sheet data available.")

    cat = bs["category"].str.lower()
    fixed_assets = float(bs.loc[cat.isin(_BS_FIXED_ASSETS), "amount"].sum())
    current_assets = float(bs.loc[cat.isin(_BS_CURRENT_ASSETS), "amount"].sum())
    total_assets = fixed_assets + current_assets

    current_liabilities = float(bs.loc[cat.isin(_BS_CURRENT_LIABILITIES), "amount"].sum())
    long_term_liabilities = float(bs.loc[cat.isin(_BS_LONG_TERM_LIABILITIES), "amount"].sum())
    total_liabilities = current_liabilities + long_term_liabilities

    net_assets = total_assets + total_liabilities  # liabilities stored negative, so this subtracts

    equity_b_fwd = float(bs.loc[cat.isin(_BS_EQUITY), "amount"].sum())  # negative (credit-natured)
    equity_current_year = -net_profit  # a profit increases equity, which is negative in this convention
    total_equity = equity_b_fwd + equity_current_year

    check = round(net_assets + total_equity, 2)
    status = "ok" if abs(check) <= MATERIALITY_AMOUNT else "review"
    message = (
        "Balance sheet balances: net assets agree to total equity (including the current year's "
        "profit, which isn't yet closed to retained earnings in the TB itself)."
        if status == "ok"
        else f"Balance sheet does not balance by £{abs(check):,.2f} - after bridging in the current year's "
             f"profit/(loss), there's still a gap. Check for an account miscategorised between P&L and B/S "
             f"(wrong Account Type in the source data), or a B/S movement not captured in the TB supplied."
    )

    rows = [
        {"Line": "Fixed assets", "Amount": round(fixed_assets, 2)},
        {"Line": "Current assets", "Amount": round(current_assets, 2)},
        {"Line": "TOTAL ASSETS", "Amount": round(total_assets, 2)},
        {"Line": "Current liabilities", "Amount": round(-current_liabilities, 2)},
        {"Line": "Long-term liabilities", "Amount": round(-long_term_liabilities, 2)},
        {"Line": "TOTAL LIABILITIES", "Amount": round(-total_liabilities, 2)},
        {"Line": "NET ASSETS", "Amount": round(net_assets, 2)},
        {"Line": "Equity brought forward", "Amount": round(-equity_b_fwd, 2)},
        {"Line": "Profit/(loss) for the year (per P&L, not yet closed in the TB)", "Amount": round(net_profit, 2)},
        {"Line": "TOTAL EQUITY", "Amount": round(-total_equity, 2)},
        {"Line": "CHECK: Net Assets - Total Equity (should be £0)", "Amount": check},
    ]
    return StatementResult(status, message, pd.DataFrame(rows), net_profit)
