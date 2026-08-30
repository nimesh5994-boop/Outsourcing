"""Related Party Transactions - surfaces postings to a contact already
known to be a related party (identified from Directors' Loan Account
activity), landing on some OTHER account - the FRS 102 Section 33
disclosure area this system had no support for at all.

Reuses compliance_checks.py's own DLA-account detection (single source
of truth for "which account is the DLA" - never re-derived a second
way) purely as the source of related-party identity: this client's own
DLA contact names ARE the related-party vocabulary, the same "client's
own data as vocabulary" technique as fixed_assets' capex suggestion,
control_accounts' miscoding suggestion, and nominal_matrix's
unallocated-transaction suggestion.

Explicitly NOT a completeness check: a director who never drew from or
repaid the DLA this year, or a related party who isn't a director at
all (a close family member, a company under common control), won't be
caught - this is the only source of related-party identity this system
has, not a director register or company data. Stated plainly in the
message every time, since silently implying completeness would be
worse than not having the check at all. Never decides what should or
shouldn't be disclosed - only ever a candidate list for the preparer to
consider.
"""
import pandas as pd

from app.compliance_checks import _DLA_PATTERN, _find_accounts
from app.recon import ReconResult

MATERIALITY_AMOUNT = 500.0


def find_related_party_transactions(
    tb_current: pd.DataFrame | None, nominal_activity: pd.DataFrame | None, threshold: float = MATERIALITY_AMOUNT,
) -> ReconResult:
    name = "Related party transactions (per Directors' Loan Account activity)"
    dla_accounts = _find_accounts(tb_current, _DLA_PATTERN)
    if dla_accounts.empty or nominal_activity is None or nominal_activity.empty:
        return ReconResult(
            name, "n/a",
            "No Directors' Loan Account found, or no nominal activity available, to identify related parties from.",
        )
    dla_codes = set(dla_accounts["account_code"].astype(str))

    activity = nominal_activity.copy()
    activity["account_code"] = activity["account_code"].astype(str)
    activity["_contact_key"] = activity["contact"].astype(str).str.strip()

    dla_activity = activity[activity["account_code"].isin(dla_codes)]
    related_parties = set(dla_activity["_contact_key"]) - {""}
    if not related_parties:
        return ReconResult(name, "n/a", "No named contacts found on the Directors' Loan Account to identify related parties from.")

    other_activity = activity[~activity["account_code"].isin(dla_codes)]
    candidates = other_activity[
        other_activity["_contact_key"].isin(related_parties)
        & ((other_activity["debit"].astype(float) > threshold) | (other_activity["credit"].astype(float) > threshold))
    ].copy()

    ok_message = (
        f"No postings above £{threshold:,.2f} outside the Directors' Loan Account found for the "
        f"{len(related_parties)} contact(s) identified as related parties from DLA activity."
    )
    if candidates.empty:
        return ReconResult(name, "ok", ok_message)

    # Same contact/date/amount can appear twice (once per side of the
    # double entry) - collapsed here so one real transaction isn't
    # counted, and shown, as two.
    candidates["_amount_key"] = (candidates["debit"].astype(float) - candidates["credit"].astype(float)).abs().round(2)
    candidates = candidates.drop_duplicates(subset=["date", "_contact_key", "_amount_key"], keep="first")
    if candidates.empty:
        return ReconResult(name, "ok", ok_message)

    candidates["Amount"] = candidates["debit"].astype(float) - candidates["credit"].astype(float)
    detail = candidates[["date", "account_code", "account_name", "reference", "description", "contact", "Amount"]].rename(columns={
        "date": "Date", "account_code": "Nominal Code", "account_name": "Coded to",
        "reference": "Reference", "description": "Description", "contact": "Contact",
    }).sort_values("Date").reset_index(drop=True)

    msg = (
        f"{len(detail)} posting(s) above £{threshold:,.2f} found for {len(related_parties)} contact(s) identified "
        f"as related parties from Directors' Loan Account activity - consider for related party disclosure "
        f"(FRS 102 Section 33). Not a completeness check: only ever catches a related party who ALSO transacted "
        f"through the DLA this year, since that's the only source of related-party identity this system has."
    )
    return ReconResult(name, "review", msg, detail)
