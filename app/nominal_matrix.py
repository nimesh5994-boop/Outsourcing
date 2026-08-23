"""Nominal activity analysis matrix - the pattern used in schedules like
'5A Bank Receipt Analysis' / '5B Bank Payments' / '12 Purchases & Expenses'
in the real working paper: every transaction against a chosen account is
allocated to its contra nominal code as a matrix column, so a reviewer can
see at a glance where the money went (or came from) and what still needs
allocating/reclassifying.

Relies on nominal_activity's contra_code/contra_name/contra_needs_review
columns, populated from Xero's 'Related account' field by xero_reports.py.
Generic-mapped (non-Xero) uploads won't have those columns, so this schedule
simply won't be produced for them - it's a Xero-specific capability today.
"""
from dataclasses import dataclass, field

import pandas as pd

MAX_CONTRA_COLUMNS = 10


@dataclass
class MatrixResult:
    account_code: str
    account_name: str
    status: str  # "ok" | "review" | "n/a"
    message: str
    matrix: pd.DataFrame = field(default_factory=pd.DataFrame)


def _label(code: str, name: str) -> str:
    code, name = (code or "").strip(), (name or "").strip()
    if code and name:
        return f"{code} - {name}"
    return code or name or "(unspecified)"


def build_matrix(account_code: str, account_name: str, nominal_activity: pd.DataFrame) -> MatrixResult:
    if nominal_activity is None or nominal_activity.empty or "contra_code" not in nominal_activity.columns:
        return MatrixResult(account_code, account_name, "n/a", "No contra-account detail available for this account (requires a Xero Account Transactions export).")

    sub = nominal_activity[nominal_activity["account_code"].astype(str) == account_code].copy()
    if sub.empty:
        return MatrixResult(account_code, account_name, "n/a", "No transactions found for this account.")

    sub["net"] = sub["debit"] - sub["credit"]
    sub["contra_label"] = sub.apply(
        lambda r: "UNALLOCATED - needs manual review" if (not r["contra_code"] and not r["contra_name"]) else _label(r["contra_code"], r["contra_name"]),
        axis=1,
    )
    sub.loc[sub["contra_needs_review"], "contra_label"] = sub.loc[sub["contra_needs_review"], "contra_label"] + " (multi-code split - review)"

    totals_by_label = sub.groupby("contra_label")["net"].sum().abs().sort_values(ascending=False)
    top_labels = list(totals_by_label.head(MAX_CONTRA_COLUMNS).index)
    sub["matrix_col"] = sub["contra_label"].where(sub["contra_label"].isin(top_labels), "OTHER (see nominal activity detail)")

    pivot = sub.pivot_table(index=["date", "reference", "description", "contact"], columns="matrix_col", values="net", aggfunc="sum", fill_value=0.0)
    pivot = pivot.reset_index()

    col_names = [c for c in pivot.columns if c not in ("date", "reference", "description", "contact")]
    pivot["TOTAL"] = pivot[col_names].sum(axis=1)
    row_total = sub.groupby(["date", "reference", "description", "contact"])["net"].sum().reset_index().rename(columns={"net": "_actual"})
    pivot = pivot.merge(row_total, on=["date", "reference", "description", "contact"], how="left")
    pivot["DIFF"] = (pivot["TOTAL"] - pivot["_actual"]).round(2)
    pivot = pivot.drop(columns=["_actual"])

    needs_review_count = int(sub["contra_needs_review"].sum())
    unallocated = sub[sub["contra_label"].str.startswith("UNALLOCATED")]["net"].sum()

    status = "ok" if needs_review_count == 0 and abs(unallocated) < 0.01 else "review"
    parts = []
    if needs_review_count:
        parts.append(f"{needs_review_count} transaction(s) had a multi-code split in the source system and need manual allocation")
    if abs(unallocated) >= 0.01:
        parts.append(f"£{unallocated:,.2f} is unallocated to any contra code")
    msg = "; ".join(parts) + "." if parts else "All transactions allocated to a contra nominal code."

    return MatrixResult(account_code, account_name, status, msg, pivot)


def build_all_matrices(tb_current: pd.DataFrame, nominal_activity: pd.DataFrame, account_codes: list[str] | None = None) -> list[MatrixResult]:
    if tb_current is None or tb_current.empty or nominal_activity is None or nominal_activity.empty:
        return []
    if account_codes is None:
        # default to the accounts with the most transaction volume, capped
        # to a sensible number of schedules
        counts = nominal_activity["account_code"].astype(str).value_counts()
        account_codes = list(counts.head(6).index)

    results = []
    for code in account_codes:
        name_rows = tb_current.loc[tb_current["account_code"].astype(str) == code, "account_name"]
        name = name_rows.iloc[0] if not name_rows.empty else code
        results.append(build_matrix(code, name, nominal_activity))
    return results
