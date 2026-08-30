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
available (typically debtors/creditors/wages/VAT/other control accounts,
plus - since these are just as much "any balance-sheet account with
nominal detail" - prepayments, stock/inventory, and non-current assets/
liabilities like a long-term loan) - Xero's own "Account Transactions"
report doesn't include bank accounts (those need a separate bank
transactions export), so bank simply won't produce a rollforward here;
it's covered by the simpler statement-vs-TB check in recon.py instead.
Fixed assets are deliberately excluded here too - they get their own,
more sophisticated treatment in fixed_assets.py (category/asset-level
rollforward, depreciation, capex-miscoding suggestions), so a fixed
asset account showing up as a second, plainer rollforward here would
just be a worse duplicate of that.
"""
from dataclasses import dataclass, field

import pandas as pd

from app.recon import ReconResult

MATERIALITY_AMOUNT = 500.0

# "current asset"/"current liability"/"liability" were the original set;
# "prepayment", "inventory", "non-current asset" and "non-current
# liability" are equally real, standard Xero account types (Prepayment,
# Inventory, Non-current Asset, Non-current Liability in Xero's own
# Account Type dropdown) that were simply missing - a prepayment or a
# long-term loan account got zero rollforward treatment at all before
# this, not because it isn't a control account, but because its type
# string wasn't in this set.
ROLLFORWARD_ACCOUNT_TYPES = {
    "current asset", "current liability", "liability",
    "prepayment", "inventory", "non-current asset", "non-current liability",
}
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
    # The actual nominal-activity postings behind the schedule's single
    # "MOVEMENTS DURING YEAR" total - same purpose as fixed_assets.py's
    # far_additions_detail: a preparer can check real transactions, not
    # just trust an aggregate.
    extra_detail: pd.DataFrame = field(default_factory=pd.DataFrame)
    extra_detail_label: str = ""
    # Only populated when `breakdown` isn't (i.e. no aged debtors/creditors
    # listing applies to this account) - the year's net movement grouped
    # by contact, so every control account gets *some* view of what's
    # driving it, not just debtors/creditors. Deliberately kept separate
    # from `breakdown`: this is the year's movement only, not the closing
    # balance make-up (no opening-balance-by-contact data exists to add to
    # it), so it's never checked against the TB balance the way the aged-
    # listing breakdown is - stated as such in movement_breakdown_label.
    movement_breakdown: pd.DataFrame = field(default_factory=pd.DataFrame)
    movement_breakdown_label: str = ""


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

    extra_detail = pd.DataFrame()
    extra_detail_label = ""
    if not movement.empty:
        extra_detail = movement[["date", "reference", "description", "contact", "debit", "credit"]].rename(columns={
            "date": "Date", "reference": "Reference", "description": "Description", "contact": "Contact",
            "debit": "Debit £", "credit": "Credit £",
        }).sort_values("Date").reset_index(drop=True)
        extra_detail_label = "Postings behind the year's movement (per nominal activity)"

    # Only every control account that has NO aged-listing breakdown gets
    # this - debtors/creditors already have an authoritative party-level
    # view (the aged listing), so adding a second, differently-scoped one
    # there would just be confusing.
    movement_breakdown = pd.DataFrame()
    movement_breakdown_label = ""
    if breakdown_out.empty and not movement.empty and movement["contact"].astype(str).str.strip().ne("").any():
        net_by_contact = movement.groupby("contact")["debit"].sum() - movement.groupby("contact")["credit"].sum()
        by_contact = net_by_contact.round(2).reset_index()
        by_contact.columns = ["Party", "Net movement £"]
        by_contact = by_contact[by_contact["Party"].astype(str).str.strip() != ""]
        if not by_contact.empty:
            movement_breakdown = by_contact.sort_values("Net movement £", key=abs, ascending=False).reset_index(drop=True)
            movement_breakdown_label = (
                "Net movement during the year by contact - NOT the closing balance make-up (no opening-"
                "balance-by-contact data is available), shown so a preparer can see what's driving this "
                "account without an aged listing to check it against."
            )

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

    return ControlAccountResult(
        account_code, account_name, status, msg, schedule, breakdown_out, breakdown_label,
        extra_detail, extra_detail_label, movement_breakdown, movement_breakdown_label,
    )


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


# --- Possible postings coded to the wrong control account ----------------
#
# Same idea as fixed_assets.suggest_capital_expenditure_reclassification:
# read the GL in depth and surface likely mis-postings using the client's
# own data, rather than trust the TB's totals in isolation. The naive
# version of this - "flag any other account this contact also posts to" -
# is far too noisy to be useful: in a double-entry export every genuine
# transaction naturally shows the contact under at least one OTHER
# account too (an invoice's contra is Sales/a P&L account, a receipt's
# contra is Bank) - that's completely normal, not a miscoding. So this
# only ever looks at a contact's postings to a DIFFERENT balance-sheet
# control-account-shaped account (ROLLFORWARD_ACCOUNT_TYPES) - never
# Bank, never a P&L account - since a customer/supplier/director's name
# turning up on an unrelated balance-sheet account, never their own
# control account, is the shape of a posting that bypassed the control
# account it should have cleared through.

def suggest_control_account_miscoding(
    tb_current: pd.DataFrame | None,
    nominal_activity: pd.DataFrame | None,
    control_accounts: list[tuple[str, str]],
    aged_debtors: pd.DataFrame | None = None,
    aged_creditors: pd.DataFrame | None = None,
    threshold: float = MATERIALITY_AMOUNT,
) -> ReconResult:
    name = "Control accounts - possible postings coded to the wrong control account"
    if tb_current is None or tb_current.empty or nominal_activity is None or nominal_activity.empty or not control_accounts:
        return ReconResult(name, "n/a", "No trial balance / nominal activity / control accounts available to scan.")

    tb = tb_current.copy()
    tb["account_code"] = tb["account_code"].astype(str)
    type_by_code = dict(zip(tb["account_code"], tb["account_type"].astype(str).str.lower().str.strip()))

    activity = nominal_activity.copy()
    activity["account_code"] = activity["account_code"].astype(str)
    activity["_contact_key"] = activity["contact"].astype(str).str.strip().str.lower()
    activity = activity[activity["_contact_key"] != ""]

    # Which control account a contact "belongs to": wherever they already
    # post within it this year, plus - for debtors/creditors specifically
    # - the aged listing's own party names, an independent ground truth
    # that doesn't depend on this year's postings at all.
    contact_home: dict[str, tuple[str, str]] = {}
    for code, acc_name in control_accounts:
        contacts_here = set(activity.loc[activity["account_code"] == code, "_contact_key"])
        name_l = acc_name.lower()
        if any(k in name_l for k in DEBTOR_KEYWORDS) and aged_debtors is not None and not aged_debtors.empty:
            contacts_here |= set(aged_debtors["customer"].astype(str).str.strip().str.lower())
        if any(k in name_l for k in CREDITOR_KEYWORDS) and aged_creditors is not None and not aged_creditors.empty:
            contacts_here |= set(aged_creditors["supplier"].astype(str).str.strip().str.lower())
        for contact_key in contacts_here:
            if contact_key and contact_key not in contact_home:
                contact_home[contact_key] = (code, acc_name)

    ok_message = f"No postings above £{threshold:,.2f} found on another balance-sheet control account for a contact known to a different one."
    if not contact_home:
        return ReconResult(name, "n/a", "No contacts found on any control account to check against.")

    candidates = activity[
        activity["_contact_key"].isin(contact_home)
        & activity["account_code"].map(type_by_code).fillna("").isin(ROLLFORWARD_ACCOUNT_TYPES)
    ].copy()
    if candidates.empty:
        return ReconResult(name, "ok", ok_message)

    candidates["_home"] = candidates["_contact_key"].map(contact_home)
    candidates = candidates[candidates["account_code"] != candidates["_home"].map(lambda h: h[0])]
    candidates = candidates[(candidates["debit"].astype(float) > threshold) | (candidates["credit"].astype(float) > threshold)]
    if candidates.empty:
        return ReconResult(name, "ok", ok_message)

    candidates["Normally clears through"] = candidates["_home"].map(lambda h: h[1])
    candidates["Amount"] = candidates["debit"].astype(float) - candidates["credit"].astype(float)
    detail = candidates[["date", "account_code", "account_name", "reference", "description", "contact", "Amount", "Normally clears through"]].rename(columns={
        "date": "Date", "account_code": "Nominal Code", "account_name": "Coded to",
        "reference": "Reference", "description": "Description", "contact": "Contact",
    }).sort_values("Date").reset_index(drop=True)

    msg = (
        f"{len(detail)} posting(s) above £{threshold:,.2f} were coded to a different balance-sheet control "
        f"account for a contact who normally clears through another one - review and check whether these "
        f"should be coded through that control account instead. Nothing here is moved automatically; this "
        f"is a candidate to check, not proof of a coding error."
    )
    return ReconResult(name, "review", msg, detail)
