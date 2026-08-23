"""Reconciliation and cross-check engine.

Every function takes canonical DataFrames (see models.py / parsers.py) and
returns a small, uniform result: a DataFrame of line-level detail plus a
summary dict with a status of "ok" / "review" / "error" and a human message.
The Excel builder consumes these results directly to prefill each schedule
and flag what needs preparer/reviewer attention.
"""
from dataclasses import dataclass, field

import pandas as pd

MATERIALITY_AMOUNT = 500.0   # £ - absolute variance below this is never flagged
VARIANCE_PCT_THRESHOLD = 0.10  # 10% movement year-on-year triggers a flag
ROUND_SUM_THRESHOLD = 1000.0   # journals at/above this that are a round number get flagged


@dataclass
class ReconResult:
    name: str
    status: str  # "ok" | "review" | "error" | "n/a"
    message: str
    detail: pd.DataFrame = field(default_factory=pd.DataFrame)


def check_tb_self_balances(tb_current: pd.DataFrame, tb_comparative: pd.DataFrame) -> ReconResult:
    rows = []
    for label, tb in (("Current year", tb_current), ("Comparative year", tb_comparative)):
        if tb is None or tb.empty:
            continue
        net = round(float(tb["debit"].sum() - tb["credit"].sum()), 2)
        rows.append({"period": label, "total_debit": tb["debit"].sum(), "total_credit": tb["credit"].sum(), "out_of_balance": net})
    detail = pd.DataFrame(rows)
    if detail.empty:
        return ReconResult("TB self-balance check", "n/a", "No trial balance uploaded.", detail)
    bad = detail[detail["out_of_balance"].abs() > 0.01]
    if not bad.empty:
        return ReconResult("TB self-balance check", "error",
                            f"{len(bad)} period(s) do not balance to zero - check for a mapping or export error.", detail)
    return ReconResult("TB self-balance check", "ok", "Trial balance debits equal credits for all periods uploaded.", detail)


def tb_to_pl_bs_tie(tb: pd.DataFrame, pl: pd.DataFrame, bs: pd.DataFrame, period_label: str) -> ReconResult:
    if tb is None or tb.empty or (pl is None or pl.empty) and (bs is None or bs.empty):
        return ReconResult(f"TB to P&L/B&S tie-out ({period_label})", "n/a", "Insufficient data uploaded for this check.")

    pl_total = float(pl["amount"].sum()) if pl is not None and not pl.empty else 0.0
    bs_total = float(bs["amount"].sum()) if bs is not None and not bs.empty else 0.0
    tb_total = float(tb["balance"].sum())  # should be ~0 if TB includes both P&L and B/S codes

    variance = round(tb_total - (pl_total + bs_total), 2)
    detail = pd.DataFrame([{
        "TB net balance": tb_total,
        "P&L total": pl_total,
        "B/S total": bs_total,
        "P&L + B/S": pl_total + bs_total,
        "Variance": variance,
    }])
    status = "ok" if abs(variance) <= 1.0 else "review"
    msg = "TB ties to P&L + B/S." if status == "ok" else f"TB does not tie to P&L + B/S by {variance:,.2f} - check for accounts missing from one report or a sign convention mismatch."
    return ReconResult(f"TB to P&L/B&S tie-out ({period_label})", status, msg, detail)


def variance_analysis(tb_current: pd.DataFrame, tb_comparative: pd.DataFrame) -> ReconResult:
    if tb_current is None or tb_current.empty:
        return ReconResult("Current vs comparative variance analysis", "n/a", "No current-year trial balance uploaded.")
    cur = tb_current.groupby(["account_code", "account_name"], as_index=False)["balance"].sum().rename(columns={"balance": "current_year"})
    if tb_comparative is not None and not tb_comparative.empty:
        comp = tb_comparative.groupby(["account_code", "account_name"], as_index=False)["balance"].sum().rename(columns={"balance": "comparative_year"})
    else:
        comp = pd.DataFrame(columns=["account_code", "account_name", "comparative_year"])

    merged = pd.merge(cur, comp, on=["account_code", "account_name"], how="outer").fillna(0.0)
    merged["variance_amount"] = merged["current_year"] - merged["comparative_year"]
    merged["variance_pct"] = merged.apply(
        lambda r: (r["variance_amount"] / abs(r["comparative_year"])) if r["comparative_year"] not in (0, 0.0) else (1.0 if r["variance_amount"] != 0 else 0.0),
        axis=1,
    )
    merged["flag"] = (
        (merged["variance_amount"].abs() >= MATERIALITY_AMOUNT) &
        (merged["variance_pct"].abs() >= VARIANCE_PCT_THRESHOLD)
    )
    merged = merged.sort_values("variance_amount", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)
    flagged = int(merged["flag"].sum())
    status = "ok" if flagged == 0 else "review"
    msg = "No nominal codes moved beyond materiality." if flagged == 0 else f"{flagged} nominal code(s) moved >{VARIANCE_PCT_THRESHOLD:.0%} and >£{MATERIALITY_AMOUNT:,.0f} year-on-year - review for reclassification or genuine business change."
    return ReconResult("Current vs comparative variance analysis", status, msg, merged)


def nominal_activity_review(nominal_current: pd.DataFrame) -> ReconResult:
    if nominal_current is None or nominal_current.empty:
        return ReconResult("Nominal activity review (reallocation candidates)", "n/a", "No nominal activity uploaded.")

    df = nominal_current.copy()
    df["amount"] = df["debit"] - df["credit"]
    df["abs_amount"] = df["amount"].abs()

    reasons = []
    for _, r in df.iterrows():
        row_reasons = []
        name_l = (r.get("account_name", "") or "").lower()
        desc_l = (r.get("description", "") or "").lower()
        if any(k in name_l for k in ("suspense", "misc", "miscellaneous", "uncategorised", "uncategorized", "unknown")):
            row_reasons.append("Posted to a suspense/miscellaneous code")
        if r["abs_amount"] >= ROUND_SUM_THRESHOLD and r["abs_amount"] % 100 == 0:
            row_reasons.append("Round-sum manual journal above threshold")
        if str(r.get("source_type", "")).lower() in ("journal", "manual journal", "je") and r["abs_amount"] >= MATERIALITY_AMOUNT:
            row_reasons.append("Material manual journal")
        if any(k in desc_l for k in ("reallocat", "correction", "reclass", "to be reviewed", "tbc", "query")):
            row_reasons.append("Description flags it as a pending correction/reallocation")
        reasons.append("; ".join(row_reasons))

    df["flag_reason"] = reasons
    flagged = df[df["flag_reason"] != ""].copy()
    flagged["proposed_reallocation_code"] = ""
    flagged["reviewer_comment"] = ""

    status = "ok" if flagged.empty else "review"
    msg = "No transactions flagged for reallocation review." if flagged.empty else f"{len(flagged)} transaction(s) flagged for reallocation/coding review."
    cols = ["date", "account_code", "account_name", "reference", "description", "contact", "source_type", "amount", "flag_reason", "proposed_reallocation_code", "reviewer_comment"]
    return ReconResult("Nominal activity review (reallocation candidates)", status, msg, flagged[cols] if not flagged.empty else pd.DataFrame(columns=cols))


def debtors_creditors_control_recon(aged: pd.DataFrame, tb: pd.DataFrame, control_account_names: list[str], party_col: str, label: str) -> ReconResult:
    if aged is None or aged.empty:
        return ReconResult(label, "n/a", "No aged report uploaded.")
    aged_total = float(aged["total"].sum())

    tb_total = 0.0
    if tb is not None and not tb.empty:
        mask = tb["account_name"].str.lower().apply(lambda n: any(k in n for k in control_account_names))
        tb_total = float(tb.loc[mask, "balance"].sum())

    # aged listings always report a positive amount owed; the TB control
    # account balance is credit-natured (negative, debit-positive convention)
    # for creditors, so compare on the same footing
    variance = round(aged_total - abs(tb_total), 2)
    detail = pd.DataFrame([{
        "Aged report total": aged_total,
        "TB control account balance": abs(tb_total),
        "Variance": variance,
    }])
    status = "ok" if abs(variance) <= MATERIALITY_AMOUNT else "review"
    msg = "Aged listing agrees to the control account." if status == "ok" else f"Aged listing does not agree to the TB control account by £{variance:,.2f} - investigate unposted invoices/credits or a control account misstatement."
    return ReconResult(label, status, msg, detail)


def bank_reconciliation(bank_statement: pd.DataFrame, tb: pd.DataFrame, bank_account_names: list[str] | None = None) -> ReconResult:
    if bank_statement is None or bank_statement.empty:
        return ReconResult("Bank reconciliation", "n/a", "No bank closing statement uploaded.")

    rows = []
    for _, b in bank_statement.iterrows():
        acct_name = b["account_name"]
        tb_balance = 0.0
        if tb is not None and not tb.empty:
            mask = tb["account_name"].str.lower().str.strip() == str(acct_name).lower().strip()
            if not mask.any() and bank_account_names:
                mask = tb["account_name"].str.lower().apply(lambda n: acct_name.lower() in n or n in acct_name.lower())
            tb_balance = float(tb.loc[mask, "balance"].sum())
        variance = round(float(b["closing_balance"]) - tb_balance, 2)
        rows.append({
            "Bank account": acct_name,
            "Statement date": b.get("statement_date", ""),
            "Statement closing balance": b["closing_balance"],
            "TB book balance": tb_balance,
            "Unreconciled variance": variance,
            "Reconciling items (to complete)": "",
        })
    detail = pd.DataFrame(rows)
    flagged = detail[detail["Unreconciled variance"].abs() > MATERIALITY_AMOUNT]
    status = "ok" if flagged.empty else "review"
    msg = "All bank accounts agree to TB within materiality." if flagged.empty else f"{len(flagged)} bank account(s) show an unreconciled variance - list reconciling items (unpresented cheques, uncleared deposits, bank charges) on the Bank Recon tab."
    return ReconResult("Bank reconciliation", status, msg, detail)


def vat_cross_check(vat_return: pd.DataFrame, pl: pd.DataFrame, tb: pd.DataFrame, vat_control_names: list[str]) -> ReconResult:
    if vat_return is None or vat_return.empty:
        return ReconResult("VAT return cross-check", "n/a", "No VAT return uploaded.")
    v = vat_return.iloc[0]

    turnover = 0.0
    if pl is not None and not pl.empty:
        mask = pl["category"].str.lower().str.contains("turnover|sales|income|revenue", regex=True, na=False)
        turnover = float(pl.loc[mask, "amount"].sum()) if mask.any() else float(pl["amount"].sum())

    vat_control_balance = 0.0
    if tb is not None and not tb.empty:
        mask = tb["account_name"].str.lower().apply(lambda n: any(k in n for k in vat_control_names))
        vat_control_balance = float(tb.loc[mask, "balance"].sum())

    sales_variance = round(float(v["box6"]) - turnover, 2)
    # VAT control is credit-natured (negative, debit-positive convention);
    # Box 5 net VAT due is reported as a positive amount owed
    vat_variance = round(float(v["box5"]) - abs(vat_control_balance), 2)

    detail = pd.DataFrame([
        {"Check": "Box 6 (net sales) vs P&L turnover", "VAT return": v["box6"], "Nominal ledger": turnover, "Variance": sales_variance},
        {"Check": "Box 5 (net VAT due) vs VAT control account (TB)", "VAT return": v["box5"], "Nominal ledger": abs(vat_control_balance), "Variance": vat_variance},
    ])
    flagged = detail[detail["Variance"].abs() > MATERIALITY_AMOUNT]
    status = "ok" if flagged.empty else "review"
    msg = "VAT return agrees to the nominal ledger within materiality." if flagged.empty else "VAT return does not agree to the nominal ledger - check for timing differences (cash vs accrual VAT scheme) or unposted VAT adjustments."
    return ReconResult("VAT return cross-check", status, msg, detail)


def run_all_recons(data: dict) -> list[ReconResult]:
    """data keys: tb_current, tb_comparative, nominal_current, aged_debtors,
    aged_creditors, vat_return, bank_statement, pl_current, bs_current."""
    results = [
        check_tb_self_balances(data.get("tb_current"), data.get("tb_comparative")),
        tb_to_pl_bs_tie(data.get("tb_current"), data.get("pl_current"), data.get("bs_current"), "current year"),
        variance_analysis(data.get("tb_current"), data.get("tb_comparative")),
        nominal_activity_review(data.get("nominal_current")),
        debtors_creditors_control_recon(
            data.get("aged_debtors"), data.get("tb_current"),
            ["debtors control", "trade debtors", "accounts receivable", "sales ledger control"],
            "customer", "Debtors control account reconciliation",
        ),
        debtors_creditors_control_recon(
            data.get("aged_creditors"), data.get("tb_current"),
            ["creditors control", "trade creditors", "accounts payable", "purchase ledger control"],
            "supplier", "Creditors control account reconciliation",
        ),
        bank_reconciliation(data.get("bank_statement"), data.get("tb_current")),
        vat_cross_check(
            data.get("vat_return"), data.get("pl_current"), data.get("tb_current"),
            ["vat control", "vat liability", "sales tax payable", "vat payable"],
        ),
    ]
    return results
