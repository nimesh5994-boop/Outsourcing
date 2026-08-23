"""Data-driven compliance checks, distilled from a real manual-job review
checklist covering DLA/S455, dividends, petty cash, and loans (Bounce Back
Loan / Hire Purchase). Everything here is the automatable subset: a check
that can be answered from data this system already ingests (TB, nominal
activity). The rest of that checklist - "agreement received?", "CT
liability checked against HMRC login?", and similar items that need a
human to confirm something outside the data - is a separate manual
pro-forma tab (see excel_builder.build_compliance_checklist_sheet /
COMPLIANCE_CHECKLIST_ITEMS), not a data-driven flag.

Every check returns a recon.ReconResult, the same shape every other check
in this system uses, so these plug straight into the existing results
list/index/recon-sheet builder with no new plumbing - same pattern as
anomaly_detection.py.
"""
import re

import pandas as pd

from app.recon import ReconResult

DLA_MONTHLY_WITHDRAWAL_THRESHOLD = 10_000.0

_DLA_PATTERN = re.compile(r"director.*?(?:current|loan)\s*account|\bdla\b", re.IGNORECASE)
_DIVIDEND_PATTERN = re.compile(r"\bdividend", re.IGNORECASE)
_RETAINED_EARNINGS_PATTERN = re.compile(r"retained earnings|profit and loss reserve|p\s*&?\s*l\s*reserve", re.IGNORECASE)
_PETTY_CASH_PATTERN = re.compile(r"petty cash|cash in hand", re.IGNORECASE)
_LOAN_PATTERNS = {
    "Bounce Back Loan": re.compile(r"bounce\s*back|\bbbl\b", re.IGNORECASE),
    "Hire Purchase": re.compile(r"hire\s*purchase|\bhp\b", re.IGNORECASE),
    "Bank Loan": re.compile(r"\bbank\s*loan\b", re.IGNORECASE),
    "CBILS/Bounce Back-style Government-backed Loan": re.compile(r"\bcbils\b|\bbbls\b", re.IGNORECASE),
}


def _find_accounts(tb: pd.DataFrame, pattern: re.Pattern) -> pd.DataFrame:
    if tb is None or tb.empty:
        return pd.DataFrame()
    return tb[tb["account_name"].astype(str).str.contains(pattern, regex=True, na=False)]


def _balance(tb: pd.DataFrame, codes) -> float:
    if tb is None or tb.empty:
        return 0.0
    mask = tb["account_code"].astype(str).isin([str(c) for c in codes])
    return float(tb.loc[mask, "balance"].sum())


def directors_loan_account_review(
    tb_current: pd.DataFrame, tb_comparative: pd.DataFrame, nominal_activity: pd.DataFrame,
) -> ReconResult:
    """Two checks in one, both keyed off any account that looks like a
    Directors' Loan/Current Account: (1) any calendar month where net
    withdrawals exceed £10,000 - the point at which HMRC treats it as a
    loan needing a benefit-in-kind/interest review, not routine drawings -
    and (2) an S455 flag plus a drafted year-end balance note when the
    account is in debit (the director owes the company) at the year end."""
    name = "Directors' loan account review (S455 / £10,000 monthly threshold)"
    dla_accounts = _find_accounts(tb_current, _DLA_PATTERN)
    if dla_accounts.empty:
        return ReconResult(name, "n/a", "No Directors' Loan/Current Account found in the trial balance.")

    codes = dla_accounts["account_code"].astype(str).tolist()
    current_balance = _balance(tb_current, codes)
    comparative_balance = _balance(tb_comparative, codes)

    rows = []
    if nominal_activity is not None and not nominal_activity.empty:
        movement = nominal_activity[nominal_activity["account_code"].astype(str).isin(codes)].copy()
        if not movement.empty:
            movement["date"] = pd.to_datetime(movement["date"], errors="coerce")
            movement = movement[movement["date"].notna()]
            movement["net"] = movement["debit"].fillna(0) - movement["credit"].fillna(0)
            monthly = movement.groupby(movement["date"].dt.to_period("M"))["net"].sum()
            for period, net in monthly.items():
                if net > DLA_MONTHLY_WITHDRAWAL_THRESHOLD:
                    rows.append({
                        "Month": str(period), "Net withdrawal £": round(float(net), 2),
                        "Flag": f"Exceeds £{DLA_MONTHLY_WITHDRAWAL_THRESHOLD:,.0f} - review for benefit-in-kind/interest, and whether a dividend (if reserves allow) should clear it.",
                    })

    # the account is in debit (positive, in this debit-positive TB
    # convention) when the director owes the company money
    s455_applies = current_balance > 0.01
    note = (
        f"At the year end the director owed the company £{current_balance:,.2f} "
        f"(comparative: £{comparative_balance:,.2f}) by way of his/her director's current account."
        if s455_applies else
        f"The director's current account was not in debit at the year end (balance: £{-current_balance:,.2f} owed to the director; "
        f"comparative: £{-comparative_balance:,.2f})."
    )
    rows.append({"Month": "YEAR END POSITION", "Net withdrawal £": round(current_balance, 2), "Flag": note})
    if s455_applies:
        rows.append({"Month": "", "Net withdrawal £": "", "Flag": "S455 tax consideration applies (loan to a participator still outstanding 9 months after the year end) - confirm repayment/dividend timing before the CT600 is finalised."})

    detail = pd.DataFrame(rows)
    flagged_months = [r for r in rows if isinstance(r.get("Net withdrawal £"), (int, float)) and r["Month"] not in ("", "YEAR END POSITION") ]
    status = "review" if (flagged_months or s455_applies) else "ok"
    parts = []
    if flagged_months:
        parts.append(f"{len(flagged_months)} month(s) with net withdrawals over £{DLA_MONTHLY_WITHDRAWAL_THRESHOLD:,.0f}")
    if s455_applies:
        parts.append("account is in debit at the year end (S455 applies)")
    msg = "; ".join(parts) + "." if parts else "No month exceeded the £10,000 withdrawal threshold, and the account is not in debit at the year end."
    return ReconResult(name, status, msg, detail)


def dividend_reserves_review(
    tb_current: pd.DataFrame, tb_comparative: pd.DataFrame, nominal_activity: pd.DataFrame, current_year_profit: float | None,
) -> ReconResult:
    """Compares dividends declared this year against distributable
    reserves actually available (retained earnings b/fwd + this year's
    profit) - flags if dividends exceed what's available, since an
    illegal dividend is a real Companies Act problem, not just a
    presentation issue."""
    name = "Dividend vs distributable reserves review"
    dividend_accounts = _find_accounts(tb_current, _DIVIDEND_PATTERN)
    if dividend_accounts.empty:
        return ReconResult(name, "n/a", "No dividend account found in the trial balance/nominal activity - no dividend appears to have been declared this year.")

    codes = dividend_accounts["account_code"].astype(str).tolist()
    dividends_declared = 0.0
    if nominal_activity is not None and not nominal_activity.empty:
        movement = nominal_activity[nominal_activity["account_code"].astype(str).isin(codes)]
        dividends_declared = abs(float(movement["debit"].sum() - movement["credit"].sum())) if not movement.empty else 0.0
    if dividends_declared == 0.0:
        dividends_declared = abs(_balance(tb_current, codes))

    re_accounts = _find_accounts(tb_comparative, _RETAINED_EARNINGS_PATTERN)
    retained_earnings_b_fwd = -_balance(tb_comparative, re_accounts["account_code"].astype(str).tolist()) if not re_accounts.empty else 0.0
    profit_for_year = float(current_year_profit) if current_year_profit is not None else 0.0
    available_reserves = retained_earnings_b_fwd + profit_for_year

    variance = round(dividends_declared - available_reserves, 2)
    detail = pd.DataFrame([{
        "Retained earnings b/fwd": round(retained_earnings_b_fwd, 2),
        "Profit/(loss) for the year": round(profit_for_year, 2),
        "Available distributable reserves": round(available_reserves, 2),
        "Dividends declared this year": round(dividends_declared, 2),
        "Headroom / (shortfall)": -variance,
    }])
    status = "ok" if variance <= 0.01 else "review"
    msg = (
        f"Dividends declared (£{dividends_declared:,.2f}) are covered by available distributable reserves (£{available_reserves:,.2f})."
        if status == "ok" else
        f"Dividends declared (£{dividends_declared:,.2f}) exceed available distributable reserves (£{available_reserves:,.2f}) by "
        f"£{variance:,.2f} - a potential unlawful dividend under the Companies Act. Confirm the reserves position/dividend timing "
        f"before the accounts are finalised."
    )
    return ReconResult(name, status, msg, detail)


def petty_cash_running_balance_review(tb_comparative: pd.DataFrame, nominal_activity: pd.DataFrame) -> ReconResult:
    """Rolls a petty cash account's running balance forward transaction by
    transaction through the year - petty cash can never legitimately go
    negative (you can't spend physical cash you don't have), so any dip
    below zero means either a mis-dated/mis-posted transaction or cash
    that was fronted by the business/director and needs booking through
    the DLA instead."""
    name = "Petty cash running balance review"

    def find(tb):
        return _find_accounts(tb, _PETTY_CASH_PATTERN)

    petty_cash_accounts = find(tb_comparative) if tb_comparative is not None else pd.DataFrame()
    if petty_cash_accounts.empty and nominal_activity is not None and not nominal_activity.empty:
        names = nominal_activity["account_name"].astype(str)
        petty_cash_accounts = nominal_activity[names.str.contains(_PETTY_CASH_PATTERN, regex=True, na=False)][["account_code", "account_name"]].drop_duplicates()
    if petty_cash_accounts.empty:
        return ReconResult(name, "n/a", "No petty cash account found.")

    codes = petty_cash_accounts["account_code"].astype(str).unique().tolist()
    if nominal_activity is None or nominal_activity.empty:
        return ReconResult(name, "n/a", "Petty cash account found, but no nominal activity uploaded to roll it forward.")

    movement = nominal_activity[nominal_activity["account_code"].astype(str).isin(codes)].copy()
    if movement.empty:
        return ReconResult(name, "n/a", "Petty cash account found, but no transactions posted to it this year.")

    movement["date"] = pd.to_datetime(movement["date"], errors="coerce")
    movement = movement.sort_values("date", na_position="last")
    b_fwd = _balance(tb_comparative, codes) if tb_comparative is not None else 0.0

    running = b_fwd
    rows = []
    dips_below_zero = 0
    for _, t in movement.iterrows():
        running += float(t["debit"] or 0) - float(t["credit"] or 0)
        if running < -0.01:
            dips_below_zero += 1
            rows.append({
                "Date": t["date"], "Reference": t.get("reference"), "Description": t.get("description"),
                "Running balance after this transaction": round(running, 2),
            })

    status = "ok" if dips_below_zero == 0 else "review"
    msg = (
        "Petty cash running balance never goes negative during the year."
        if status == "ok" else
        f"Petty cash running balance goes negative on {dips_below_zero} occasion(s) - review for a mis-dated/mis-posted "
        f"transaction, or cash fronted by the business/director that should be booked through the DLA instead."
    )
    return ReconResult(name, status, msg, pd.DataFrame(rows))


def loan_facility_review(tb_current: pd.DataFrame, tb_comparative: pd.DataFrame) -> ReconResult:
    """Presence detection for Bounce Back Loan / Hire Purchase / Bank Loan
    -style accounts - not a computed check (there's no repayment schedule
    or agreement in the data this system has), but a reminder of the
    specific checklist points that apply whenever one of these is found:
    confirm the agreement/statement, confirm interest is calculated
    correctly (BBL: no interest in the first 12 months), and split the
    closing balance between amounts due within/after one year."""
    name = "Loan facility review (BBL / Hire Purchase / Bank Loan)"
    if tb_current is None or tb_current.empty:
        return ReconResult(name, "n/a", "No trial balance uploaded.")

    rows = []
    for label, pattern in _LOAN_PATTERNS.items():
        found = _find_accounts(tb_current, pattern)
        if found.empty:
            continue
        for _, r in found.iterrows():
            comp_balance = _balance(tb_comparative, [r["account_code"]]) if tb_comparative is not None else 0.0
            reminders = {
                "Bounce Back Loan": "Confirm no interest is charged in the first 12 months, and that repayment/interest afterwards is per the BBL calculator.",
                "Hire Purchase": "Confirm the agreement was received (interest rate, term, deposit, purchase fee), and split the closing balance between due within one year and due after one year.",
                "Bank Loan": "Confirm the statement was received for the year, and split the closing balance between due within one year and due after one year.",
                "CBILS/Bounce Back-style Government-backed Loan": "Confirm the facility terms (interest holiday period, repayment start date) and split the closing balance between due within one year and due after one year.",
            }[label]
            rows.append({
                "Facility type": label, "Nominal code": r["account_code"], "Account name": r["account_name"],
                "Current year balance": round(float(r["balance"]), 2), "Comparative year balance": round(comp_balance, 2),
                "Reminder": reminders,
            })

    detail = pd.DataFrame(rows)
    if detail.empty:
        return ReconResult(name, "n/a", "No Bounce Back Loan, Hire Purchase, or Bank Loan account found in the trial balance.")
    status = "review"
    msg = f"{len(detail)} loan/finance facility account(s) found - see the reminder against each for the specific checklist points to confirm before sign-off."
    return ReconResult(name, status, msg, detail)


def run_all_compliance_checks(
    tb_current: pd.DataFrame, tb_comparative: pd.DataFrame, nominal_activity: pd.DataFrame, current_year_profit: float | None = None,
) -> list[ReconResult]:
    return [
        directors_loan_account_review(tb_current, tb_comparative, nominal_activity),
        dividend_reserves_review(tb_current, tb_comparative, nominal_activity, current_year_profit),
        petty_cash_running_balance_review(tb_comparative, nominal_activity),
        loan_facility_review(tb_current, tb_comparative),
    ]
