"""Accruals & Prepayments schedule - the consolidated view a real
working paper always has: every prepayment and accrual line side by
side in one table (brought forward, movement, carried forward, checked
against the TB), rather than scattered across a separate tab per
account the way control_accounts.py's per-account rollforwards present
them. Built for the same reason fixed_assets.py has both a category-
level summary AND an asset-level detail - two different views of the
same balance sheet area, each useful on its own; a prepayment account
still also gets its own individual rollforward tab via control_
accounts.py (now that ROLLFORWARD_ACCOUNT_TYPES covers "prepayment" -
see that module), this is the one-table view alongside it.

Two categories, discovered differently since Xero doesn't have a
distinct "Accrual" account type the way it has a genuine "Prepayment"
one:

  - Prepayments: TB accounts of Xero's own "Prepayment" account type -
    no ambiguity, the platform already says what these are.
  - Accruals: Current Liability accounts whose name reads like an
    accrual (specific keyword phrases, same reasoning as control_
    accounts.py's DEBTOR_KEYWORDS/CREDITOR_KEYWORDS - a loose "accrue"
    substring would also catch unrelated accounts like "Accrued
    Corporation Tax", which already has its own dedicated CT schedule
    and shouldn't be double-counted here).
"""
import pandas as pd

from app.recon import ReconResult

MATERIALITY_AMOUNT = 500.0

ACCRUAL_KEYWORDS = ("accrued expense", "accrued expenses", "accrued cost", "accrued costs", "accrued charge", "accruals")
# deliberately excluded from the accrual keyword match even though they
# contain "accrued" - each already has its own dedicated, more specific
# check elsewhere, so including them here would double-count the same
# balance under two different schedules
_ACCRUAL_EXCLUDE_KEYWORDS = ("corporation tax", "vat", "paye", "national insurance", "pension")


def _find_accounts(tb_current: pd.DataFrame) -> list[tuple[str, str, str]]:
    """Returns (code, name, kind) - kind is "Prepayment" or "Accrual"."""
    if tb_current is None or tb_current.empty:
        return []
    type_l = tb_current["account_type"].astype(str).str.lower()
    name_l = tb_current["account_name"].astype(str).str.lower()

    rows = []
    prepayments = tb_current[type_l == "prepayment"]
    for _, r in prepayments.iterrows():
        rows.append((str(r["account_code"]), r["account_name"], "Prepayment"))

    is_accrual_name = name_l.apply(
        lambda n: any(k in n for k in ACCRUAL_KEYWORDS) and not any(x in n for x in _ACCRUAL_EXCLUDE_KEYWORDS)
    )
    accruals = tb_current[(type_l == "current liability") & is_accrual_name]
    for _, r in accruals.iterrows():
        rows.append((str(r["account_code"]), r["account_name"], "Accrual"))

    return rows


def build_schedule(
    tb_current: pd.DataFrame | None, tb_comparative: pd.DataFrame | None, nominal_activity: pd.DataFrame | None,
) -> ReconResult:
    name = "Accruals & Prepayments schedule"
    accounts = _find_accounts(tb_current)
    if not accounts:
        return ReconResult(name, "n/a", "No Prepayment-typed or accrual-named Current Liability accounts found in the trial balance.")

    def balance(tb, code):
        if tb is None or tb.empty:
            return 0.0
        mask = tb["account_code"].astype(str) == code
        return float(tb.loc[mask, "balance"].sum())

    def movement(code):
        if nominal_activity is None or nominal_activity.empty:
            return 0.0, 0.0
        m = nominal_activity[nominal_activity["account_code"].astype(str) == code]
        return float(m["debit"].sum()), float(m["credit"].sum())

    rows = []
    for code, acc_name, kind in accounts:
        b_fwd = balance(tb_comparative, code)
        c_fwd = balance(tb_current, code)
        debit, credit = movement(code)
        net_movement = debit - credit
        computed_c_fwd = b_fwd + net_movement
        diff = round(computed_c_fwd - c_fwd, 2)
        rows.append({
            "Type": kind, "Nominal code": code, "Description": acc_name,
            "Balance b/fwd": round(b_fwd, 2), "Movement": round(net_movement, 2),
            "Balance c/fwd (per TB)": round(c_fwd, 2), "Diff": diff,
        })

    detail = pd.DataFrame(rows)
    flagged = detail[detail["Diff"].abs() > MATERIALITY_AMOUNT]
    status = "ok" if flagged.empty else "review"

    total_prepayments = float(detail.loc[detail["Type"] == "Prepayment", "Balance c/fwd (per TB)"].sum())
    total_accruals = float(detail.loc[detail["Type"] == "Accrual", "Balance c/fwd (per TB)"].sum())
    msg = (
        f"{len(detail)} prepayment/accrual account(s): £{total_prepayments:,.2f} prepaid, "
        f"£{abs(total_accruals):,.2f} accrued. "
    )
    msg += (
        "Every line ties to the TB." if status == "ok"
        else f"{len(flagged)} line(s) don't tie to the TB - check for movements outside the nominal activity supplied (e.g. a manual journal)."
    )

    result = ReconResult(name, status, msg, detail)

    if nominal_activity is not None and not nominal_activity.empty:
        code_to_kind = {code: kind for code, _, kind in accounts}
        movement_detail = nominal_activity[nominal_activity["account_code"].astype(str).isin(code_to_kind)].copy()
        if not movement_detail.empty:
            movement_detail["Type"] = movement_detail["account_code"].astype(str).map(code_to_kind)
            result.extra_detail = movement_detail[[
                "Type", "date", "account_code", "account_name", "reference", "description", "contact", "debit", "credit",
            ]].rename(columns={
                "date": "Date", "account_code": "Nominal Code", "account_name": "Account",
                "reference": "Reference", "description": "Description", "contact": "Contact",
                "debit": "Debit £", "credit": "Credit £",
            }).sort_values(["Type", "Date"]).reset_index(drop=True)
            result.extra_detail_label = "Postings behind each line's movement (per nominal activity)"

    return result
