"""PAYE Reconciliation - BrightPay payroll data vs the General Ledger.

Three independent checks, because real small-business bookkeeping settles
these three differently (confirmed against a real BrightPay + Xero export
pair, not assumed):

  - Net Pay: paid to each employee individually, so it's reconciled
    employee-by-employee-by-month against General Ledger postings whose
    contact matches that employee's name - wherever in the ledger that
    posting actually landed (a dedicated wages control account, or coded
    straight to another account entirely), not restricted to one assumed
    account. That's the point of matching by contact across the whole
    ledger rather than casting one control account's balance: a
    bookkeeper who codes things differently than expected is exactly the
    scenario this needs to survive.
  - HMRC PAYE & NI: HMRC is paid one lump sum a month (there is no
    per-employee trace of this in the ledger), so it's reconciled at
    company-wide monthly total against General Ledger postings whose
    contact looks like HMRC.
  - Pension contributions: paid to the pension provider in one lump sum
    a month per employee+employer contributions combined, same shape as
    HMRC - reconciled monthly against postings whose contact looks like
    a pension provider.

No reference/invoice numbers exist anywhere in this domain (BrightPay
doesn't generate them and neither does a bank payment run), so matching
is a single pass per item: contact (fuzzy name match) + amount (within
tolerance), picking the nearest-dated candidate within a configurable
window when more than one qualifies, and never letting the same General
Ledger row settle two different months. An item that doesn't find a
within-tolerance match still gets the *nearest same-contact posting*
reported as a diagnostic candidate (ignoring amount) - "no match" is
rarely as useful to a preparer as "here's the closest thing, look at the
difference yourself".

Returns app.recon.ReconResult objects - the same shape every other check
returns - so this plugs into the results list, the job summary, and the
Excel builder without any of them needing to know this module exists.
"""
import re
from dataclasses import dataclass

import pandas as pd

from app.recon import ReconResult

DEFAULT_TOLERANCE = 0.0  # £
DEFAULT_DATE_WINDOW_DAYS = 60  # HMRC/pension payments trail the period end by days to weeks, never land same-day

HMRC_CONTACT_KEYWORDS = ["hmrc", "hm revenue"]
PENSION_CONTACT_KEYWORDS = [
    "nest", "now:pensions", "now pensions", "peoples pension", "people's pension",
    "smart pension", "aviva pension", "royal london", "scottish widows", "standard life", "the pensions trust",
]

GL_DISPLAY = ["date", "reference", "contact", "description", "debit", "credit"]


@dataclass
class PayeReconSettings:
    tolerance: float = DEFAULT_TOLERANCE
    date_window_days: int = DEFAULT_DATE_WINDOW_DAYS


def _normalise_words(name) -> list[str]:
    return re.sub(r"[^a-z0-9\s]", "", str(name).lower()).split()


def _employee_contact_match(gl_contact, employee_name) -> bool:
    """True if the GL contact's first and last word both appear in the
    employee's name - handles "Jamie Doe" (GL) matching "Jamie Francis
    Doe" (BrightPay's full legal name), a real mismatch found against an
    actual client export, without loosening to a single-word match that
    would false-positive on a common first name alone."""
    gl_words, employee_words = _normalise_words(gl_contact), _normalise_words(employee_name)
    if not gl_words or not employee_words:
        return False
    return gl_words[0] == employee_words[0] and gl_words[-1] == employee_words[-1]


def _keyword_contact_match(gl_contact, keywords: list[str]) -> bool:
    contact = str(gl_contact).lower()
    return any(k in contact for k in keywords)


def _gl_amount(row) -> float:
    """Net cash movement for a GL row, debit-positive convention (matches
    how a payment OUT - to an employee, HMRC, or a pension provider -
    already appears in the sample data: a positive Debit)."""
    return float(row["debit"]) - float(row["credit"])


CASH_SOURCE_KEYWORDS = ["spend money", "receive money", "receivable payment", "payment", "bank transfer"]


def _looks_like_cash_movement(source_type) -> bool:
    return any(k in str(source_type).lower() for k in CASH_SOURCE_KEYWORDS)


def _prepare_pool(nominal_activity: pd.DataFrame | None) -> tuple[pd.DataFrame, str]:
    """Two distinct things can put a posting in the General Ledger for
    the same contact and amount, and BrightPay's figures (money actually
    paid) should only ever be checked against the first:

      1. A cash settlement - Xero's Source column reads "Payment",
         "Receivable Payment", "Spend Money" or "Receive Money". This is
         what actually left the bank.
      2. An accrual booking with no settlement yet - "Payable Invoice",
         "Sales Invoice", a manual journal - money recorded as owed, not
         necessarily paid. Confirmed against a real export: an invoice
         raised in March and paid in April shows up as two GL rows on
         two different dates, and a second invoice with no payment row
         at all by year end is exactly "recorded but not yet paid" -
         neither should be treated as a completed payment.

    Filtered to (1) here, with the exclusion count folded into the
    result each check returns - never a silent assumption, since which
    postings count as "paid" is exactly the kind of judgement call a
    preparer needs to be able to check, not just trust.

    Within the cash-settlement rows, a single real payment still lists
    twice - one row per side of the double entry the bank leg and the
    account it was coded against, same contact, same date, equal-and-
    opposite debit/credit. Collapsed here by magnitude (only the sign
    convention differs between which leg a row represents), so it isn't
    treated as two independent candidates, or left as a phantom second
    candidate for some other month once one leg is matched."""
    columns = ["date", "reference", "contact", "description", "debit", "credit", "_amount", "_pool_id"]
    if nominal_activity is None or nominal_activity.empty:
        return pd.DataFrame(columns=columns), ""

    pool = nominal_activity.copy()
    pool["_amount"] = pool.apply(_gl_amount, axis=1).abs()

    has_source_type = "source_type" in pool.columns and pool["source_type"].astype(str).str.strip().ne("").any()
    if has_source_type:
        cash_mask = pool["source_type"].apply(_looks_like_cash_movement)
        excluded = int((~cash_mask).sum())
        pool = pool[cash_mask]
        note = (
            f"General Ledger narrowed to {len(pool)} cash-settlement posting(s) (Payment/Spend Money/Receive "
            f"Money) - {excluded} accrual-only posting(s) (an invoice or bill with no matching payment row) "
            f"excluded, since these figures are money actually paid, not amounts invoiced."
        )
    else:
        note = (
            "This General Ledger upload doesn't identify each posting's source type, so cash-settlement "
            "postings couldn't be told apart from accrual-only bookings (an invoice with no payment yet) - "
            "every posting to a matching contact was treated as a candidate."
        )

    pool = pool.drop_duplicates(subset=["date", "contact", "_amount"], keep="first").reset_index(drop=True)
    pool["_pool_id"] = range(len(pool))
    return pool, note


def _closest_candidate(pool: pd.DataFrame, contact_match) -> dict | None:
    """Diagnostic only: the nearest-dated posting to this contact,
    ignoring amount entirely - shown on an unmatched item so a preparer
    has something concrete to look at instead of just "not found"."""
    candidates = pool[pool["contact"].apply(contact_match)]
    if candidates.empty:
        return None
    row = candidates.iloc[0]
    return {"date": row["date"], "reference": row.get("reference", ""), "amount": row["_amount"]}


def _match_one(pool: pd.DataFrame, contact_match, target_amount: float, target_date, tolerance: float, window_days: int) -> tuple[pd.Series | None, pd.DataFrame]:
    candidates = pool[
        pool["contact"].apply(contact_match)
        & ((pool["_amount"] - target_amount).abs() <= (tolerance + 1e-9))
    ]
    if pd.notna(target_date):
        # The date window is a hard requirement, not a preference: a
        # same-contact, same-amount posting six months away is a
        # coincidence, not this payment - excluded rather than silently
        # accepted just because nothing closer happened to qualify.
        in_window = candidates["date"].apply(
            lambda d: pd.notna(d) and abs((d - target_date).days) <= window_days
        )
        candidates = candidates[in_window]
    if candidates.empty:
        return None, pool
    if pd.notna(target_date) and candidates["date"].notna().any():
        candidates = candidates.assign(_gap=(candidates["date"] - target_date).abs())
        best = candidates.sort_values("_gap").iloc[0]
    else:
        best = candidates.iloc[0]
    return best, pool[pool["_pool_id"] != best["_pool_id"]]


def reconcile_net_pay(payroll_summary: pd.DataFrame | None, nominal_activity: pd.DataFrame | None, settings: PayeReconSettings) -> ReconResult:
    name = "PAYE Recon - Net Pay by Employee"
    if payroll_summary is None or payroll_summary.empty:
        return ReconResult(name, "n/a", "No BrightPay Payroll Summary uploaded.")

    pool, note = _prepare_pool(nominal_activity)
    items = payroll_summary[payroll_summary["net_pay"] > 0]
    if items.empty:
        return ReconResult(name, "n/a", "No months with a positive net pay figure to reconcile.")

    matched_rows, unmatched_rows = [], []
    for _, item in items.iterrows():
        contact_match = lambda c, emp=item["employee"]: _employee_contact_match(c, emp)
        best, pool = _match_one(pool, contact_match, item["net_pay"], item["period_end"], settings.tolerance, settings.date_window_days)
        if best is not None:
            matched_rows.append({
                "Employee": item["employee"], "Month ending": item["period_end"], "Net pay (BrightPay)": item["net_pay"],
                "GL date": best["date"], "GL amount": best["_amount"], "Variance": round(item["net_pay"] - best["_amount"], 2),
            })
        else:
            hint = _closest_candidate(pool, contact_match)
            unmatched_rows.append({
                "Employee": item["employee"], "Month ending": item["period_end"], "Net pay (BrightPay)": item["net_pay"],
                "Closest GL candidate": (f"£{hint['amount']:.2f} on {hint['date'].date() if pd.notna(hint['date']) else '?'}" if hint else "none found for this contact"),
            })

    return _build_result(name, matched_rows, unmatched_rows, "Employee", settings, "employee-month(s)", note)


def reconcile_hmrc(p32: pd.DataFrame | None, nominal_activity: pd.DataFrame | None, settings: PayeReconSettings) -> ReconResult:
    name = "PAYE Recon - HMRC PAYE & NI"
    if p32 is None or p32.empty:
        return ReconResult(name, "n/a", "No BrightPay P32 uploaded.")

    pool, note = _prepare_pool(nominal_activity)
    items = p32[p32["amount_due"] > 0]
    if items.empty:
        return ReconResult(name, "n/a", "No tax months with an amount due to reconcile.")

    contact_match = lambda c: _keyword_contact_match(c, HMRC_CONTACT_KEYWORDS)
    matched_rows, unmatched_rows = [], []
    for _, item in items.iterrows():
        best, pool = _match_one(pool, contact_match, item["amount_due"], item["period_end"], settings.tolerance, settings.date_window_days)
        if best is not None:
            matched_rows.append({
                "Month ending": item["period_end"], "Amount due (P32)": item["amount_due"],
                "GL date": best["date"], "GL amount": best["_amount"], "Variance": round(item["amount_due"] - best["_amount"], 2),
            })
        else:
            hint = _closest_candidate(pool, contact_match)
            unmatched_rows.append({
                "Month ending": item["period_end"], "Amount due (P32)": item["amount_due"],
                "Closest GL candidate": (f"£{hint['amount']:.2f} on {hint['date'].date() if pd.notna(hint['date']) else '?'}" if hint else "no HMRC-like contact found in the ledger"),
            })

    return _build_result(name, matched_rows, unmatched_rows, "Month ending", settings, "tax month(s)", note)


def reconcile_pension(pensions: pd.DataFrame | None, nominal_activity: pd.DataFrame | None, settings: PayeReconSettings) -> ReconResult:
    name = "PAYE Recon - Pension Contributions"
    if pensions is None or pensions.empty:
        return ReconResult(name, "n/a", "No BrightPay Pensions report uploaded.")

    monthly = pensions.groupby("period_end", as_index=False)["total_pension"].sum()
    monthly = monthly[monthly["total_pension"] > 0]
    if monthly.empty:
        return ReconResult(name, "n/a", "No months with a pension contribution to reconcile.")

    pool, note = _prepare_pool(nominal_activity)
    contact_match = lambda c: _keyword_contact_match(c, PENSION_CONTACT_KEYWORDS)
    matched_rows, unmatched_rows = [], []
    for _, item in monthly.iterrows():
        best, pool = _match_one(pool, contact_match, item["total_pension"], item["period_end"], settings.tolerance, settings.date_window_days)
        if best is not None:
            matched_rows.append({
                "Month ending": item["period_end"], "Pension contributions (BrightPay)": item["total_pension"],
                "GL date": best["date"], "GL amount": best["_amount"], "Variance": round(item["total_pension"] - best["_amount"], 2),
            })
        else:
            hint = _closest_candidate(pool, contact_match)
            unmatched_rows.append({
                "Month ending": item["period_end"], "Pension contributions (BrightPay)": item["total_pension"],
                "Closest GL candidate": (f"£{hint['amount']:.2f} on {hint['date'].date() if pd.notna(hint['date']) else '?'}" if hint else "no pension-provider-like contact found in the ledger"),
            })

    return _build_result(name, matched_rows, unmatched_rows, "Month ending", settings, "month(s)", note)


def _build_result(name: str, matched_rows: list[dict], unmatched_rows: list[dict], key_label: str, settings: PayeReconSettings, unit: str, note: str = "") -> ReconResult:
    """Unlike VAT's reference-based pass (which deliberately ignores
    amount so it can catch "same invoice, wrong amount"), every match
    here is only ever found by requiring the amount within tolerance in
    the first place - there is no pass that could produce a "matched but
    the amount differs outside tolerance" case, so status is just
    matched vs. unmatched. The Variance column on a matched row still
    shows any small in-tolerance drift, for transparency, without it
    being treated as its own exception.

    `note` (from _prepare_pool) states, in the message every time, which
    General Ledger postings were even in scope to be matched against
    (cash-settlement only, and how many accrual-only rows that excluded)
    - the assumption behind every match/no-match conclusion below, not
    left implicit."""
    matched = pd.DataFrame(matched_rows)
    unmatched = pd.DataFrame(unmatched_rows)
    total_items = len(matched) + len(unmatched)

    detail = pd.DataFrame([
        {"Measure": "Matched to General Ledger", "Count": len(matched)},
        {"Measure": "With no General Ledger match", "Count": len(unmatched)},
    ])

    status = "ok" if len(unmatched) == 0 else "review"
    tol_note = "exact match" if settings.tolerance == 0 else f"£{settings.tolerance:.2f} tolerance"
    msg = (
        f"All {total_items} {unit} matched to the General Ledger ({tol_note}, within {settings.date_window_days} days)."
        if status == "ok" else
        f"{len(unmatched)} of {total_items} {unit} have no General Ledger match "
        f"({tol_note}, within {settings.date_window_days} days) - review the exceptions below."
    )
    if note:
        msg = f"{msg} {note}"

    result = ReconResult(name, status, msg, detail)
    if not unmatched.empty:
        result.extra_detail = unmatched
        result.extra_detail_label = f"{name} - items requiring review"
    return result


def reconcile(data: dict, settings: PayeReconSettings) -> list[ReconResult]:
    """data keys: paye_summary, paye_p32, paye_pensions (the three
    BrightPay uploads), nominal_current (the job's existing General
    Ledger upload - reused as-is, no separate GL upload for this
    section). Always returns three results in this order."""
    nominal_activity = data.get("nominal_current")
    return [
        reconcile_net_pay(data.get("paye_summary"), nominal_activity, settings),
        reconcile_hmrc(data.get("paye_p32"), nominal_activity, settings),
        reconcile_pension(data.get("paye_pensions"), nominal_activity, settings),
    ]
