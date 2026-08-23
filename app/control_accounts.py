"""Control account rollforward schedules - the T-account style check used
throughout the real working paper file (e.g. 'Other Debtors Control
Account', 'Net Wages Control Account'): Balance b/fwd + movements during
the year should equal Balance c/fwd per the current year trial balance.
Any difference is exactly what needs journalling or investigating.

Where a closing balance belongs to debtors or creditors, the schedule also
carries the aged listing as a breakdown of that balance directly underneath
- so the reader sees what the balance is actually made up of, not just its
total, and any gap between the two is a single clearly-flagged number left
for the preparer to explain rather than silently absorbed.

Scoped to balance-sheet control accounts that have nominal-ledger detail
available (typically debtors/creditors/wages/VAT/other control accounts) -
Xero's own "Account Transactions" report doesn't include bank accounts
(those need a separate bank transactions export), so bank simply won't
produce a rollforward here; it's covered by the simpler statement-vs-TB
check in recon.py instead.
"""
from dataclasses import dataclass, field

import pandas as pd

MATERIALITY_AMOUNT = 500.0

ROLLFORWARD_ACCOUNT_TYPES = {"current asset", "current liability", "liability"}
EXCLUDE_NAME_KEYWORDS = ("retained earnings", "share capital", "profit for the year", "dividend")
# deliberately specific phrases - a loose "debtor"/"creditor" substring match
# would also catch unrelated accounts like "Sundry Creditors" or
# "Corporation Tax Payable", which should NOT get the trade debtors/
# creditors aged listing attached as their breakdown
DEBTOR_KEYWORDS = ("debtors control", "trade debtors", "accounts receivable", "sales ledger control")
CREDITOR_KEYWORDS = ("creditors control", "trade creditors", "accounts payable", "purchase ledger control")


@dataclass
class ControlAccountResult:
    account_code: str
    account_name: str
    status: str  # "ok" | "review" | "n/a"
    message: str
    schedule: pd.DataFrame = field(default_factory=pd.DataFrame)
    breakdown: pd.DataFrame = field(default_factory=pd.DataFrame)
    breakdown_label: str = ""


def find_control_accounts(tb_current: pd.DataFrame, nominal_activity: pd.DataFrame) -> list[tuple[str, str]]:
    """Balance-sheet accounts present in the current TB that also have
    nominal-activity detail available, so a rollforward is actually possible."""
    if tb_current is None or tb_current.empty or nominal_activity is None or nominal_activity.empty:
        return []
    codes_with_activity = set(nominal_activity["account_code"].astype(str))
    candidates = tb_current[
        tb_current["account_type"].str.lower().isin(ROLLFORWARD_ACCOUNT_TYPES)
        & ~tb_current["account_name"].str.lower().apply(lambda n: any(k in n for k in EXCLUDE_NAME_KEYWORDS))
        & tb_current["account_code"].astype(str).isin(codes_with_activity)
    ]
    return list(zip(candidates["account_code"].astype(str), candidates["account_name"]))


def _breakdown_for(account_name: str, aged_debtors: pd.DataFrame, aged_creditors: pd.DataFrame) -> tuple[pd.DataFrame, str, str]:
    name_l = account_name.lower()
    if any(k in name_l for k in DEBTOR_KEYWORDS) and aged_debtors is not None and not aged_debtors.empty:
        return aged_debtors, "customer", "Aged debtors listing"
    if any(k in name_l for k in CREDITOR_KEYWORDS) and aged_creditors is not None and not aged_creditors.empty:
        return aged_creditors, "supplier", "Aged creditors listing"
    return pd.DataFrame(), "", ""


def build_rollforward(
    account_code: str,
    account_name: str,
    tb_current: pd.DataFrame,
    tb_comparative: pd.DataFrame,
    nominal_activity: pd.DataFrame,
    aged_debtors: pd.DataFrame | None = None,
    aged_creditors: pd.DataFrame | None = None,
) -> ControlAccountResult:
    b_fwd = 0.0
    if tb_comparative is not None and not tb_comparative.empty:
        mask = tb_comparative["account_code"].astype(str) == account_code
        b_fwd = float(tb_comparative.loc[mask, "balance"].sum())

    c_fwd_per_tb = 0.0
    if tb_current is not None and not tb_current.empty:
        mask = tb_current["account_code"].astype(str) == account_code
        c_fwd_per_tb = float(tb_current.loc[mask, "balance"].sum())

    has_activity = nominal_activity is not None and not nominal_activity.empty
    movement = nominal_activity[nominal_activity["account_code"].astype(str) == account_code] if has_activity else pd.DataFrame()
    total_debit = float(movement["debit"].sum()) if not movement.empty else 0.0
    total_credit = float(movement["credit"].sum()) if not movement.empty else 0.0
    net_movement = total_debit - total_credit

    computed_c_fwd = b_fwd + net_movement
    difference = round(computed_c_fwd - c_fwd_per_tb, 2)
    rollforward_ties = movement.empty is False and abs(difference) <= MATERIALITY_AMOUNT

    def dr_cr(amount: float) -> tuple[float, float]:
        return (round(amount, 2), 0.0) if amount >= 0 else (0.0, round(-amount, 2))

    rows = []
    dr, cr = dr_cr(b_fwd)
    rows.append({"Item": "BALANCE B/FWD", "Reference": "Comparative year TB", "Debit £": dr, "Credit £": cr})
    if movement.empty:
        rows.append({"Item": "MOVEMENTS DURING YEAR", "Reference": "No nominal activity detail available", "Debit £": "", "Credit £": ""})
    else:
        rows.append({"Item": "MOVEMENTS DURING YEAR", "Reference": "Nominal activity detail", "Debit £": round(total_debit, 2), "Credit £": round(total_credit, 2)})
    dr, cr = dr_cr(c_fwd_per_tb)
    rows.append({"Item": "BALANCE C/FWD (per TB)", "Reference": "Current year TB", "Debit £": dr, "Credit £": cr})
    if not movement.empty:
        dr, cr = dr_cr(-difference)
        rows.append({"Item": "DIFFERENCE (computed c/fwd vs per TB)", "Reference": "", "Debit £": dr, "Credit £": cr})
    schedule = pd.DataFrame(rows)

    breakdown_df, party_col, breakdown_label = _breakdown_for(account_name, aged_debtors, aged_creditors)
    breakdown_diff = None
    breakdown_out = pd.DataFrame()
    if not breakdown_df.empty:
        breakdown_out = breakdown_df[[party_col, "total"]].rename(columns={party_col: "Party", "total": "Amount £"}).copy()
        breakdown_total = float(breakdown_out["Amount £"].sum())
        breakdown_diff = round(breakdown_total - abs(c_fwd_per_tb), 2)
        breakdown_out = pd.concat([
            breakdown_out,
            pd.DataFrame([
                {"Party": "TOTAL PER BREAKDOWN", "Amount £": round(breakdown_total, 2)},
                {"Party": "BALANCE PER TB", "Amount £": round(abs(c_fwd_per_tb), 2)},
                {"Party": "UNEXPLAINED DIFFERENCE (for preparer to review)", "Amount £": breakdown_diff},
            ]),
        ], ignore_index=True)

    breakdown_ok = breakdown_diff is None or abs(breakdown_diff) <= MATERIALITY_AMOUNT
    status = "ok" if rollforward_ties and breakdown_ok else ("n/a" if movement.empty and breakdown_diff is None else "review")

    parts = []
    if not movement.empty and not rollforward_ties:
        parts.append(f"rollforward does not agree to the trial balance by £{abs(difference):,.2f}")
    if movement.empty:
        parts.append("no nominal activity detail supplied for this account, so the movement can't be rolled forward")
    if breakdown_diff is not None and not breakdown_ok:
        parts.append(f"£{abs(breakdown_diff):,.2f} of the closing balance isn't explained by the {breakdown_label.lower()}")
    if not parts:
        msg = "Rollforward and breakdown both agree to the trial balance."
    else:
        msg = "; ".join(parts).capitalize() + " - review and correct as needed."

    return ControlAccountResult(account_code, account_name, status, msg, schedule, breakdown_out, breakdown_label)


def build_all_rollforwards(
    tb_current: pd.DataFrame,
    tb_comparative: pd.DataFrame,
    nominal_activity: pd.DataFrame,
    aged_debtors: pd.DataFrame | None = None,
    aged_creditors: pd.DataFrame | None = None,
) -> list[ControlAccountResult]:
    accounts = dict(find_control_accounts(tb_current, nominal_activity))

    # debtors/creditors control accounts still deserve a breakdown-only
    # schedule even without nominal activity detail, since the aged listing
    # alone is enough to show what makes up the balance
    if tb_current is not None and not tb_current.empty:
        for _, row in tb_current.iterrows():
            name_l = str(row["account_name"]).lower()
            has_breakdown_source = (
                (any(k in name_l for k in DEBTOR_KEYWORDS) and aged_debtors is not None and not aged_debtors.empty)
                or (any(k in name_l for k in CREDITOR_KEYWORDS) and aged_creditors is not None and not aged_creditors.empty)
            )
            code = str(row["account_code"])
            if has_breakdown_source and code not in accounts:
                accounts[code] = row["account_name"]

    return [
        build_rollforward(code, name, tb_current, tb_comparative, nominal_activity, aged_debtors, aged_creditors)
        for code, name in accounts.items()
    ]
