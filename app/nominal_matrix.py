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

from app.recon import ReconResult

MAX_CONTRA_COLUMNS = 10
MATERIALITY_AMOUNT = 500.0
OTHER_LABEL = "OTHER (see nominal activity detail)"
UNALLOCATED_PREFIX = "UNALLOCATED"


@dataclass
class MatrixResult:
    account_code: str
    account_name: str
    status: str  # "ok" | "review" | "n/a"
    message: str
    matrix: pd.DataFrame = field(default_factory=pd.DataFrame)
    # What's actually inside the "OTHER" column - same purpose as
    # fixed_assets.far_additions_detail: a preparer can check the real
    # contra accounts folded together, not just trust an opaque bucket.
    extra_detail: pd.DataFrame = field(default_factory=pd.DataFrame)
    extra_detail_label: str = ""


def _label(code: str, name: str) -> str:
    code, name = (code or "").strip(), (name or "").strip()
    if code and name:
        return f"{code} - {name}"
    return code or name or "(unspecified)"


def _label_rows(account_code: str, nominal_activity: pd.DataFrame) -> pd.DataFrame | None:
    """Filters nominal_activity to this account and buckets each row into
    its contra_label / matrix_col (top N contra accounts by value, the rest
    folded into OTHER) - the shared basis for both build_matrix's pivot
    values and build_matrix_row_groups' row-id groupings (used by the
    formula-linked Excel builder), so the two can never drift apart on
    which transaction landed in which bucket."""
    sub = nominal_activity[nominal_activity["account_code"].astype(str) == account_code].copy()
    if sub.empty:
        return None

    sub["net"] = sub["debit"] - sub["credit"]
    sub["contra_label"] = sub.apply(
        lambda r: f"{UNALLOCATED_PREFIX} - needs manual review" if (not r["contra_code"] and not r["contra_name"]) else _label(r["contra_code"], r["contra_name"]),
        axis=1,
    )
    sub.loc[sub["contra_needs_review"], "contra_label"] = sub.loc[sub["contra_needs_review"], "contra_label"] + " (multi-code split - review)"

    totals_by_label = sub.groupby("contra_label")["net"].sum().abs().sort_values(ascending=False)
    top_labels = list(totals_by_label.head(MAX_CONTRA_COLUMNS).index)
    sub["matrix_col"] = sub["contra_label"].where(sub["contra_label"].isin(top_labels), OTHER_LABEL)
    return sub


def _other_bucket_breakdown(sub: pd.DataFrame) -> pd.DataFrame:
    """The contra accounts actually folded together into the OTHER column
    - each with its own net amount and transaction count - so "OTHER" is
    never a dead end for a reviewer, and a materially large contra account
    that only missed the top-N cut isn't left looking immaterial.

    Excludes UNALLOCATED: with enough distinct contra accounts it can
    itself get folded into OTHER by the same column cap, but its amount
    is already called out on its own in the result message - repeating it
    here as if it were just another hidden contra account would conflate
    two different things (a posting with no contra code at all, vs. one
    that has a contra code but didn't make the top-N cut)."""
    other = sub[(sub["matrix_col"] == OTHER_LABEL) & ~sub["contra_label"].str.startswith(UNALLOCATED_PREFIX)]
    if other.empty:
        return pd.DataFrame()
    breakdown = other.groupby("contra_label")["net"].agg(["sum", "count"]).reset_index()
    breakdown.columns = ["Contra account", "Net amount £", "Transaction count"]
    breakdown["Net amount £"] = breakdown["Net amount £"].round(2)
    return breakdown.sort_values("Net amount £", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)


def build_matrix(
    account_code: str, account_name: str, nominal_activity: pd.DataFrame, materiality: float = MATERIALITY_AMOUNT,
) -> MatrixResult:
    if nominal_activity is None or nominal_activity.empty or "contra_code" not in nominal_activity.columns:
        return MatrixResult(account_code, account_name, "n/a", "No contra-account detail available for this account (requires a Xero Account Transactions export).")

    sub = _label_rows(account_code, nominal_activity)
    if sub is None:
        return MatrixResult(account_code, account_name, "n/a", "No transactions found for this account.")

    pivot = sub.pivot_table(index=["date", "reference", "description", "contact"], columns="matrix_col", values="net", aggfunc="sum", fill_value=0.0)
    pivot = pivot.reset_index()

    col_names = [c for c in pivot.columns if c not in ("date", "reference", "description", "contact")]
    pivot["TOTAL"] = pivot[col_names].sum(axis=1)
    row_total = sub.groupby(["date", "reference", "description", "contact"])["net"].sum().reset_index().rename(columns={"net": "_actual"})
    pivot = pivot.merge(row_total, on=["date", "reference", "description", "contact"], how="left")
    pivot["DIFF"] = (pivot["TOTAL"] - pivot["_actual"]).round(2)
    pivot = pivot.drop(columns=["_actual"])

    needs_review_count = int(sub["contra_needs_review"].sum())
    unallocated = sub[sub["contra_label"].str.startswith(UNALLOCATED_PREFIX)]["net"].sum()

    other_breakdown = _other_bucket_breakdown(sub)
    material_other = other_breakdown[other_breakdown["Net amount £"].abs() > materiality] if not other_breakdown.empty else pd.DataFrame()

    status = "ok" if needs_review_count == 0 and abs(unallocated) < 0.01 else "review"
    parts = []
    if needs_review_count:
        parts.append(f"{needs_review_count} transaction(s) had a multi-code split in the source system and need manual allocation")
    if abs(unallocated) >= 0.01:
        parts.append(f"£{unallocated:,.2f} is unallocated to any contra code")
    msg = "; ".join(parts) + "." if parts else "All transactions allocated to a contra nominal code."
    if not material_other.empty:
        # Advisory only, same as fixed_assets' system-estimated depreciation
        # - never a reason to flag "review" by itself, since folding into
        # OTHER purely by the display cap isn't itself wrong, just worth
        # knowing about; stated explicitly rather than left silent.
        msg += (
            f" £{material_other['Net amount £'].abs().sum():,.2f} of contra activity across "
            f"{len(material_other)} account(s) was folded into '{OTHER_LABEL}' only because it didn't rank "
            f"in the top {MAX_CONTRA_COLUMNS} columns by value - see the breakdown below."
        )

    result = MatrixResult(account_code, account_name, status, msg, pivot)
    if not other_breakdown.empty:
        result.extra_detail = other_breakdown
        result.extra_detail_label = f"'{OTHER_LABEL}' column breakdown"
    return result


def build_matrix_row_groups(account_code: str, nominal_activity_with_row_ids: pd.DataFrame) -> tuple[list[dict], list[str]]:
    """Same bucketing as build_matrix, but for the formula-linked Excel
    builder: instead of computing each cell's value in Python, returns
    each pivot row's grouping key plus - per matrix column - the row_ids
    (from the DATA_Nominal sheet's synthetic RowID column, see
    data_sheets.py) of the transactions bucketed into it. The Excel sheet
    then re-sums the raw debit/credit figures for exactly those row_ids via
    formula, rather than trusting a Python-computed literal - so recompute
    matches the pivot's contra-account bucketing, but the number itself
    stays live against the raw data. Column order matches build_matrix's
    pivot_table (sorted, pandas' default)."""
    sub = _label_rows(account_code, nominal_activity_with_row_ids)
    if sub is None:
        return [], []

    col_names = sorted(sub["matrix_col"].unique())
    rows = []
    for (date, reference, description, contact), g in sub.groupby(["date", "reference", "description", "contact"], dropna=False):
        row_ids_by_column = {
            col: g.loc[g["matrix_col"] == col, "row_id"].tolist()
            for col in col_names if (g["matrix_col"] == col).any()
        }
        rows.append({
            "date": date, "reference": reference, "description": description, "contact": contact,
            "row_ids_by_column": row_ids_by_column,
        })
    return rows, col_names


def build_all_matrices(
    tb_current: pd.DataFrame, nominal_activity: pd.DataFrame, account_codes: list[str] | None = None,
    materiality: float = MATERIALITY_AMOUNT,
) -> list[MatrixResult]:
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
        results.append(build_matrix(code, name, nominal_activity, materiality))
    return results


# --- Suggested allocations for unallocated transactions -------------------
#
# Same idea as fixed_assets.suggest_capital_expenditure_reclassification
# and control_accounts.suggest_control_account_miscoding: use the
# client's OWN data as the vocabulary, not a generic guess. Here, that's
# a contact's other, already-allocated transactions on the same account -
# if a customer/supplier's spend on this account is consistently coded to
# one contra account, an unallocated transaction for that same contact is
# probably meant to go there too. Only suggested when that pattern
# clearly dominates (a real majority, over a real sample) - otherwise
# left as a genuine unallocated item with no guess attached. Never
# allocates anything itself.

MIN_HISTORY_FOR_SUGGESTION = 2
DOMINANT_SHARE_THRESHOLD = 0.6


def suggest_contra_for_unallocated(sub: pd.DataFrame) -> pd.DataFrame:
    unallocated = sub[sub["contra_label"].str.startswith(UNALLOCATED_PREFIX)]
    if unallocated.empty:
        return pd.DataFrame()
    allocated = sub[
        ~sub["contra_label"].str.startswith(UNALLOCATED_PREFIX) & ~sub["contra_needs_review"]
    ]
    if allocated.empty:
        return pd.DataFrame()

    rows = []
    for _, txn in unallocated.iterrows():
        contact = str(txn.get("contact", "") or "").strip()
        if not contact:
            continue
        history = allocated[allocated["contact"].astype(str).str.strip() == contact]
        if len(history) < MIN_HISTORY_FOR_SUGGESTION:
            continue
        counts = history["contra_label"].value_counts()
        top_label, top_count = counts.index[0], int(counts.iloc[0])
        if top_count / len(history) < DOMINANT_SHARE_THRESHOLD:
            continue
        rows.append({
            "Date": txn["date"], "Reference": txn["reference"], "Description": txn["description"],
            "Contact": contact, "Amount": round(float(txn["net"]), 2),
            "Suggested contra account": top_label,
            "Based on": f"{top_count} of {len(history)} of this contact's other transactions on this account",
        })
    return pd.DataFrame(rows)


def suggest_unallocated_reallocations(nominal_activity: pd.DataFrame | None, account_codes: list[str] | None = None) -> ReconResult:
    name = "Nominal activity - suggested allocations for unallocated transactions"
    if nominal_activity is None or nominal_activity.empty or "contra_code" not in nominal_activity.columns:
        return ReconResult(name, "n/a", "No contra-account detail available (requires a Xero Account Transactions export).")

    if account_codes is None:
        counts = nominal_activity["account_code"].astype(str).value_counts()
        account_codes = list(counts.head(6).index)

    all_suggestions = []
    for code in account_codes:
        sub = _label_rows(code, nominal_activity)
        if sub is None:
            continue
        suggestions = suggest_contra_for_unallocated(sub)
        if not suggestions.empty:
            suggestions.insert(0, "Account", code)
            all_suggestions.append(suggestions)

    if not all_suggestions:
        return ReconResult(name, "ok", "No unallocated transaction had a clear enough pattern in that contact's other transactions to suggest an allocation.")

    detail = pd.concat(all_suggestions, ignore_index=True).sort_values("Date").reset_index(drop=True)
    msg = (
        f"{len(detail)} unallocated transaction(s) have a likely contra account based on that contact's own "
        f"other transactions on the same account - a suggestion to check, not an automatic allocation."
    )
    return ReconResult(name, "review", msg, detail)
