"""Trial Balance Tie-Out: the accounting identity every current-year
account should satisfy - opening balance + this year's nominal-ledger
movement = the balance the Trial Balance actually reports as closing -
checked for every account (not just the balance-sheet ones
control_accounts.py already gives a full rollforward schedule to), plus
a cross-check that the two Trial Balance uploads agree on where the
current year opened.

Deliberately current-year only: the comparative year is already filed
and finalised, so it isn't re-verified here - it's only ever used as the
reference point for this year's opening balances and for variance/notes
elsewhere (pl_variance.py, the TB Lead Schedule). Re-tying it out would
just be re-auditing work that's already done.

This is the "does everything the client's bookkeeping software says tie
together" layer of an Extended Trial Balance: two independent, purely
arithmetic checks, each surfacing a candidate for the preparer's own
professional judgement rather than resolving anything automatically -
posting the actual fix (in Xero, then re-generating) stays a manual step.

Deliberately kept separate from control_accounts.py: that module builds
a full T-account rollforward schedule (with an aged-listing tie-out,
contact-level movement, etc.) for the specific accounts where that depth
of drill-down earns its keep. This module runs the same underlying
identity across *every* account in the TB - including P&L codes, where
"opening balance" is simply 0 - as one compact exception-based table,
not a schedule per account.
"""
import pandas as pd

from app.recon import ReconResult
from app.xero_reports import PL_ACCOUNT_TYPES

NAME = "Trial balance tie-out (opening + movement = closing)"
TOLERANCE = 0.01  # a rounding allowance, not a materiality judgement - this identity is exact


def _movement_by_code(nominal: pd.DataFrame | None) -> dict[str, float]:
    if nominal is None or nominal.empty:
        return {}
    grouped = nominal.groupby("account_code").agg(debit=("debit", "sum"), credit=("credit", "sum"))
    return (grouped["debit"] - grouped["credit"]).to_dict()


def _tieout_table(tb_current: pd.DataFrame, tb_opening: pd.DataFrame | None, nominal_current: pd.DataFrame | None) -> pd.DataFrame:
    """One row per current-year account code: opening (always 0 for a P&L
    code, since income-statement accounts don't carry an opening balance
    forward the way a balance-sheet one does - what a TB shows mid-year
    is a running year-to-date actual, not a b/fwd; using tb_opening's own
    prior-year closing P&L figure here would compare this year's
    movement against last year's total, which never ties and isn't even
    the right question), this year's movement per the nominal ledger,
    the two added together, and how that compares to what the TB itself
    reports as the closing balance."""
    opening_by_code = tb_opening.groupby("account_code")["balance"].sum().to_dict() if tb_opening is not None and not tb_opening.empty else {}
    type_by_code = dict(zip(tb_current["account_code"].astype(str), tb_current["account_type"].astype(str).str.lower()))
    movement = _movement_by_code(nominal_current)
    codes_with_movement_data = set(movement.keys())
    has_nominal_upload = nominal_current is not None and not nominal_current.empty

    rows = []
    for _, r in tb_current.iterrows():
        code = str(r["account_code"])
        account_type = type_by_code.get(code, "")
        is_pl_account = account_type in PL_ACCOUNT_TYPES
        opening = 0.0 if is_pl_account else float(opening_by_code.get(code, 0.0))
        reported_closing = float(r["balance"])
        if account_type == "bank":
            # Xero's own "Account Transactions" report structurally never
            # includes bank accounts (a bank feed needs a separate bank
            # statement/transactions export - see control_accounts.py's
            # same exclusion, and recon.py's simpler statement-vs-TB check
            # instead) - found live: without this, a bank account with a
            # real £3k+ year of movement but zero rows in the nominal
            # upload read as "dormant, opening should equal closing",
            # flagging its entire year of legitimate activity as an
            # unexplained difference.
            movement_value, derived_closing, diff, flag_text = None, None, None, "n/a"
        elif not has_nominal_upload:
            movement_value, derived_closing, diff, flag_text = None, None, None, "n/a"
        elif code not in codes_with_movement_data:
            # No postings at all this year for this code is a legitimate,
            # common case (a dormant account) - opening should simply equal
            # closing, so it's still checkable, not "no data".
            movement_value = 0.0
            derived_closing = opening
            diff = round(derived_closing - reported_closing, 2)
            flag_text = "REVIEW" if abs(diff) > TOLERANCE else "OK"
        else:
            movement_value = movement[code]
            derived_closing = opening + movement_value
            diff = round(derived_closing - reported_closing, 2)
            flag_text = "REVIEW" if abs(diff) > TOLERANCE else "OK"
        rows.append({
            "Account Code": code,
            "Account Name": r["account_name"],
            "Account Type": r["account_type"],
            "Opening (per comparative TB)": round(opening, 2),
            "Movement (current year)": round(movement_value, 2) if movement_value is not None else "",
            "Derived Closing": round(derived_closing, 2) if derived_closing is not None else "",
            "Reported Closing (per current TB)": round(reported_closing, 2),
            "Diff": diff if diff is not None else "",
            "Flag": flag_text,
        })
    return pd.DataFrame(rows)


def _opening_balance_disagreements(tb_current_own_comparative: pd.DataFrame | None, tb_comparative: pd.DataFrame | None) -> pd.DataFrame:
    """Cross-checks the current Trial Balance's own embedded comparative
    column against a *separately* uploaded comparative-period Trial
    Balance for the same account - these should always agree (both claim
    to say what the same account closed at last year end, the starting
    point for everything in this year's tie-out above), so a
    disagreement usually means a restated comparative, a wrong file
    uploaded, or a genuine post-close adjustment made in one but not the
    other. Only ever produces an exception list (rows that actually
    disagree), not a full account-by-account table."""
    if tb_current_own_comparative is None or tb_current_own_comparative.empty or tb_comparative is None or tb_comparative.empty:
        return pd.DataFrame()

    embedded = tb_current_own_comparative.groupby("account_code")["balance"].sum()
    uploaded = tb_comparative.groupby("account_code")["balance"].sum()
    names = pd.concat([
        tb_current_own_comparative.drop_duplicates("account_code").set_index("account_code")["account_name"],
        tb_comparative.drop_duplicates("account_code").set_index("account_code")["account_name"],
    ])
    names = names[~names.index.duplicated(keep="first")]

    rows = []
    for code in sorted(set(embedded.index) | set(uploaded.index)):
        embedded_value = float(embedded.get(code, 0.0))
        uploaded_value = float(uploaded.get(code, 0.0))
        diff = round(embedded_value - uploaded_value, 2)
        if abs(diff) <= TOLERANCE:
            continue
        rows.append({
            "Account Code": code,
            "Account Name": names.get(code, ""),
            "Per current TB's own comparative column": round(embedded_value, 2),
            "Per separately-uploaded comparative TB": round(uploaded_value, 2),
            "Diff": diff,
        })
    return pd.DataFrame(rows)


def build_tieout(
    tb_current: pd.DataFrame | None,
    tb_comparative: pd.DataFrame | None,
    nominal_current: pd.DataFrame | None,
    tb_current_own_comparative: pd.DataFrame | None = None,
) -> ReconResult:
    if tb_current is None or tb_current.empty:
        return ReconResult(NAME, "n/a", "No current-year Trial Balance uploaded.")

    current_table = _tieout_table(tb_current, tb_comparative, nominal_current)
    opening_disagreements = _opening_balance_disagreements(tb_current_own_comparative, tb_comparative)

    flagged_current = current_table[current_table["Flag"] == "REVIEW"] if not current_table.empty else pd.DataFrame()

    parts = []
    if not flagged_current.empty:
        parts.append(f"{len(flagged_current)} account(s) don't tie out for the current year")
    if not opening_disagreements.empty:
        parts.append(f"{len(opening_disagreements)} account(s) have an opening balance that disagrees between the two Trial Balance uploads")

    status = "review" if parts else "ok"
    message = (
        "; ".join(parts) + " - each is opening balance + the year's nominal-ledger movement failing to reach "
        "the balance the Trial Balance itself reports (or, for the second kind, two Trial Balance uploads "
        "disagreeing on the same account's opening figure). Review each for an unposted entry, timing "
        "difference, or miscoding, and make a professional judgement on whether an adjustment is needed - "
        "nothing here is changed automatically. The comparative year isn't re-checked here - it's already "
        "filed and finalised, and is used only as this year's opening reference."
        if parts else
        "Every account's opening balance plus the current year's nominal-ledger movement ties to the balance "
        "the Trial Balance itself reports as closing."
    )

    result = ReconResult(NAME, status, message, current_table)
    if not opening_disagreements.empty:
        result.matched_detail = opening_disagreements
        result.matched_detail_label = "Opening balance disagreements between the two Trial Balance uploads"
    return result
