"""Going Concern indicators - the balance-sheet-shape red flags a real
working paper file always checks for, computed from the already-derived
Balance Sheet statement (financial_statements.build_bs_statement)
rather than re-deriving asset/liability categorisation a third time in
a third module.

Two purely arithmetic indicators, both standard going-concern red flags:

  - Negative net current assets (current liabilities > current assets):
    a working-capital deficit - on the numbers alone, the company may
    struggle to meet its short-term obligations as they fall due.
  - Negative net assets (total liabilities > total assets): balance
    sheet insolvency.

Never a verdict on whether the company IS a going concern - that needs
judgement and forward-looking information (cash flow forecasts, facility
renewal terms, director support, trading since the year end) this
system has no way to derive from historical accounting data alone. This
only ever states the arithmetic fact, explicitly labelled as such, so
it isn't missed rather than because the numbers alone decide the
question.
"""
import pandas as pd

from app.recon import ReconResult

MATERIALITY_AMOUNT = 500.0


def _line(statement: pd.DataFrame, label: str) -> float | None:
    match = statement[statement["Line"] == label]
    if match.empty:
        return None
    return float(match.iloc[0]["Amount"])


def assess(bs_statement: pd.DataFrame | None) -> ReconResult:
    name = "Going concern indicators"
    if bs_statement is None or bs_statement.empty:
        return ReconResult(name, "n/a", "No Balance Sheet data available to assess.")

    current_assets = _line(bs_statement, "Current assets")
    # financial_statements.build_bs_statement displays this line flipped
    # to a positive "amount of liability" so it reads naturally (same
    # treatment as debtors/creditors elsewhere) - NOT negative the way
    # "NET ASSETS" below is genuinely signed, so it's subtracted here,
    # not added.
    current_liabilities = _line(bs_statement, "Current liabilities")
    net_assets = _line(bs_statement, "NET ASSETS")
    if current_assets is None or current_liabilities is None or net_assets is None:
        return ReconResult(name, "n/a", "Balance Sheet statement is missing the lines needed to assess this.")

    net_current_assets = round(current_assets - current_liabilities, 2)

    detail = pd.DataFrame([
        {"Indicator": "Current assets", "Amount": round(current_assets, 2)},
        {"Indicator": "Current liabilities", "Amount": round(current_liabilities, 2)},
        {"Indicator": "Net current assets / (liabilities)", "Amount": net_current_assets},
        {"Indicator": "Net assets / (liabilities)", "Amount": round(net_assets, 2)},
    ])

    flags = []
    if net_current_assets < -MATERIALITY_AMOUNT:
        flags.append(
            f"net current LIABILITIES of £{abs(net_current_assets):,.2f} - current liabilities exceed current "
            f"assets, a working-capital deficit"
        )
    if net_assets < -MATERIALITY_AMOUNT:
        flags.append(
            f"net LIABILITIES of £{abs(net_assets):,.2f} - total liabilities exceed total assets "
            f"(balance sheet insolvency)"
        )

    status = "ok" if not flags else "review"
    if status == "ok":
        msg = "No balance-sheet-shape going concern indicators found - net current assets and net assets are both positive."
    else:
        msg = (
            "; ".join(flags).capitalize() + ". Purely an arithmetic fact about the balance sheet shape, not a "
            "verdict on going concern - that needs judgement and forward-looking information (cash flow forecasts, "
            "facility renewal terms, director support) this system has no way to derive from historical accounting "
            "data alone; flagged here so it isn't missed."
        )
    return ReconResult(name, status, msg, detail)
