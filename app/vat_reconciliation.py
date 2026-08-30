"""VAT Reconciliation - General Ledger vs Filed VAT Return matching.

Box 1 (VAT due on sales - output tax) and Box 4 (VAT reclaimed on
purchases - input tax) are reconciled independently against the same
General Ledger transaction pool, Box 1 first then Box 4, through three
cascading passes per box - strongest match first:

  1. reference (invoice/bill number) alone - not gated on amount, since
     the point of matching by invoice number is to catch a genuinely
     wrong VAT amount posted against a real invoice; that variance is
     then reported (see _review_rows), not hidden by requiring the
     amounts to already agree before the two sides count as the same
     transaction.
  2. date + VAT amount within tolerance + contact
  3. VAT amount within tolerance + contact (weakest - flagged "verify")

Each GL row is claimed by at most one filed-return row across BOTH
boxes: Box 1 matches first and removes what it claims from the shared
pool before Box 4 runs, so a sales transaction can't also get claimed
as a purchase. Whatever GL activity neither box's filed return
accounts for is reported once, as its own "General Ledger Coverage"
check, rather than being reported as an "exception" under both boxes
independently - a purchase transaction is not a Box 1 exception just
because Box 1's filed sales never mention it, and vice versa.

"VAT Accounting Basis" (accrual vs cash) picks which General Ledger
date is used for the date-based pass: accrual uses the transaction/
invoice date (always present); cash uses a payment date when the GL
export actually carries one, falling back to the transaction date
row-by-row when it doesn't (a GL without payment dates can't be cash-
matched more precisely than that, so it degrades rather than drops
rows).

Deliberately free of any FastAPI/storage/upload concerns - takes plain
canonical DataFrames (see models.py: vat_gl / vat_filed_sales /
vat_filed_purchases) and returns app.recon.ReconResult objects, the
exact same shape every other check in the pipeline returns, so this
plugs into the results list, the job summary table, the AI notes
agent, and the Excel builder without any of them needing to know this
module exists.
"""
import re
from dataclasses import dataclass

import pandas as pd

from app.recon import ReconResult

DEFAULT_TOLERANCE = 0.0  # £ - "Exact Match"
FLOAT_NOISE_FLOOR = 0.005  # £ - always absorbed regardless of tolerance, to avoid penny/float rounding false positives

GL_COLUMNS = ["date", "reference", "contact", "description", "net_amount", "vat_amount"]
FILED_COLUMNS = GL_COLUMNS + ["source_file"]
GL_DISPLAY_COLUMNS = {"date": "Date", "reference": "Reference", "contact": "Contact", "description": "Description", "net_amount": "Net Amount", "vat_amount": "VAT Amount"}
REVIEW_COLUMNS = ["Exception type", "Date", "Reference", "Contact", "Net Amount", "VAT Amount", "Source File", "Note"]


@dataclass
class VatReconSettings:
    accounting_basis: str = "accrual"  # "accrual" | "cash"
    tolerance: float = DEFAULT_TOLERANCE

    @property
    def variance_threshold(self) -> float:
        """The tolerance the user asked for a match to be found within is
        also the threshold below which a found match's residual VAT
        difference isn't worth flagging - never below the float-noise
        floor, so "Exact Match (£0.00)" still absorbs penny rounding."""
        return max(self.tolerance, FLOAT_NOISE_FLOOR)


def _normalise_contact(name) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def _empty_gl() -> pd.DataFrame:
    return pd.DataFrame(columns=GL_COLUMNS)


def _basis_date_column(gl: pd.DataFrame, basis: str) -> pd.Series:
    if gl.empty:
        return gl["date"]
    if basis == "cash" and "payment_date" in gl.columns:
        return gl["payment_date"].fillna(gl["date"])
    return gl["date"]


def _amounts_close(a: pd.Series, b: float, tolerance: float) -> pd.Series:
    return (a.round(2) - round(float(b), 2)).abs() <= (tolerance + 1e-9)


def _prepare_pool(gl: pd.DataFrame | None, settings: VatReconSettings) -> pd.DataFrame:
    pool = gl.copy() if gl is not None and not gl.empty else _empty_gl()
    pool["_match_date"] = _basis_date_column(pool, settings.accounting_basis)
    pool["_contact_key"] = pool["contact"].apply(_normalise_contact)
    pool["_ref_key"] = pool["reference"].astype(str).str.strip().str.lower()
    pool["_pool_id"] = range(len(pool))
    return pool


def match_box(pool: pd.DataFrame, filed: pd.DataFrame | None, settings: VatReconSettings) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Matches filed rows against `pool` (a GL pool already prepared by
    _prepare_pool, and possibly already trimmed by an earlier box's
    matches - see reconcile()). Returns (matched, unmatched_filed,
    remaining_pool): remaining_pool is what's left of `pool` once this
    box's matches are removed, ready to be handed to the next box."""
    if filed is None or filed.empty:
        return pd.DataFrame(), pd.DataFrame(), pool

    filed = filed.copy()
    filed["_contact_key"] = filed["contact"].apply(_normalise_contact)
    filed["_ref_key"] = filed["reference"].astype(str).str.strip().str.lower()

    matched_rows = []
    unmatched_filed_idx = []

    for filed_idx, frow in filed.iterrows():
        candidates, match_basis = None, None

        if frow["_ref_key"]:
            cand = pool[pool["_ref_key"] == frow["_ref_key"]]
            if not cand.empty:
                candidates, match_basis = cand, "reference"

        if candidates is None:
            cand = pool[
                (pool["_match_date"] == frow["date"])
                & _amounts_close(pool["vat_amount"], frow["vat_amount"], settings.tolerance)
                & (pool["_contact_key"] == frow["_contact_key"])
            ]
            if not cand.empty:
                candidates, match_basis = cand, "date + amount + contact"

        if candidates is None:
            cand = pool[
                _amounts_close(pool["vat_amount"], frow["vat_amount"], settings.tolerance)
                & (pool["_contact_key"] == frow["_contact_key"])
            ]
            if not cand.empty:
                candidates, match_basis = cand, "amount + contact (verify)"

        if candidates is None or candidates.empty:
            unmatched_filed_idx.append(filed_idx)
            continue

        gl_row = candidates.iloc[0]
        pool = pool[pool["_pool_id"] != gl_row["_pool_id"]]
        matched_rows.append({
            "Match basis": match_basis,
            "Filed date": frow["date"], "Filed reference": frow["reference"], "Filed contact": frow["contact"],
            "Filed net": frow["net_amount"], "Filed VAT": frow["vat_amount"], "Source file": frow.get("source_file", ""),
            "GL date": gl_row["date"], "GL reference": gl_row["reference"], "GL contact": gl_row["contact"],
            "GL net": gl_row["net_amount"], "GL VAT": gl_row["vat_amount"],
            "VAT variance": round(float(frow["vat_amount"]) - float(gl_row["vat_amount"]), 2),
        })

    matched = pd.DataFrame(matched_rows)
    unmatched_filed = filed.loc[unmatched_filed_idx, [c for c in FILED_COLUMNS if c in filed.columns]]
    return matched, unmatched_filed, pool


def _box_review_rows(unmatched_filed: pd.DataFrame, matched: pd.DataFrame, variance_threshold: float) -> pd.DataFrame:
    rows = []
    for _, r in unmatched_filed.iterrows():
        rows.append({
            "Exception type": "Filed return item - no GL match", "Date": r["date"], "Reference": r["reference"],
            "Contact": r["contact"], "Net Amount": r["net_amount"], "VAT Amount": r["vat_amount"],
            "Source File": r.get("source_file", ""), "Note": "",
        })
    if not matched.empty:
        variances = matched[matched["VAT variance"].abs() > variance_threshold]
        for _, r in variances.iterrows():
            rows.append({
                "Exception type": "Matched but VAT amount differs", "Date": r["Filed date"], "Reference": r["Filed reference"],
                "Contact": r["Filed contact"], "Net Amount": r["Filed net"], "VAT Amount": r["Filed VAT"],
                "Source File": r["Source file"],
                "Note": f"GL shows VAT {r['GL VAT']:.2f} on {r['GL reference'] or '(no reference)'} - variance {r['VAT variance']:.2f}",
            })
    return pd.DataFrame(rows, columns=REVIEW_COLUMNS)


def _box_result(box_label: str, filed: pd.DataFrame | None, matched: pd.DataFrame, unmatched_filed: pd.DataFrame, settings: VatReconSettings) -> ReconResult:
    name = f"VAT Recon - {box_label}"
    if filed is None or filed.empty:
        return ReconResult(name, "n/a", f"No filed VAT return detail uploaded for {box_label}.")

    filed_total = round(float(filed["vat_amount"].sum()), 2)
    matched_total = round(float(matched["Filed VAT"].sum()), 2) if not matched.empty else 0.0
    unmatched_filed_total = round(float(unmatched_filed["vat_amount"].sum()), 2) if not unmatched_filed.empty else 0.0
    variance_threshold = settings.variance_threshold
    variance_count = int((matched["VAT variance"].abs() > variance_threshold).sum()) if not matched.empty else 0
    variance_total = round(float(matched.loc[matched["VAT variance"].abs() > variance_threshold, "VAT variance"].abs().sum()), 2) if not matched.empty else 0.0

    detail = pd.DataFrame([
        {"Measure": "Filed return - total VAT", "Amount": filed_total, "Count": len(filed)},
        {"Measure": "Matched to General Ledger", "Amount": matched_total, "Count": len(matched)},
        {"Measure": "Filed items with no GL match", "Amount": unmatched_filed_total, "Count": len(unmatched_filed)},
        {"Measure": "Matched items with a VAT variance", "Amount": variance_total, "Count": variance_count},
    ])

    n_exceptions = len(unmatched_filed) + variance_count
    status = "ok" if n_exceptions == 0 else "review"
    basis_note = "cash basis (payment date)" if settings.accounting_basis == "cash" else "accrual basis (transaction date)"
    tol_note = "exact match" if settings.tolerance == 0 else f"£{settings.tolerance:.2f} tolerance"
    msg = (
        f"All {len(filed)} filed {box_label} item(s) matched to the General Ledger ({basis_note}, {tol_note})."
        if status == "ok" else
        f"{len(unmatched_filed)} filed item(s) with no GL match, {variance_count} matched item(s) with a VAT variance "
        f"({basis_note}, {tol_note}) - review the exceptions below."
    )

    result = ReconResult(name, status, msg, detail)
    review = _box_review_rows(unmatched_filed, matched, variance_threshold)
    if not review.empty:
        result.extra_detail = review
        result.extra_detail_label = f"{box_label} - items requiring review"
    return result


def _gl_coverage_result(original_gl: pd.DataFrame | None, remaining_pool: pd.DataFrame) -> ReconResult:
    name = "VAT Recon - General Ledger Coverage"
    if original_gl is None or original_gl.empty:
        return ReconResult(name, "n/a", "No General Ledger uploaded for VAT reconciliation.")

    unmatched = remaining_pool[GL_COLUMNS] if remaining_pool is not None and not remaining_pool.empty else pd.DataFrame(columns=GL_COLUMNS)
    if unmatched.empty:
        return ReconResult(name, "ok", "Every General Ledger transaction was accounted for in the filed VAT return (Box 1 or Box 4).")

    total = round(float(unmatched["vat_amount"].sum()), 2)
    detail = pd.DataFrame([{"Measure": "GL items matching neither Box 1 nor Box 4 of the filed return", "Amount": total, "Count": len(unmatched)}])
    result = ReconResult(
        name, "review",
        f"{len(unmatched)} General Ledger item(s) (£{total:.2f} VAT) don't correspond to anything in the filed VAT "
        f"return's Box 1 or Box 4 detail - check these were correctly excluded, or that they're simply missing from what was filed.",
        detail,
    )
    result.extra_detail = unmatched.rename(columns=GL_DISPLAY_COLUMNS)
    result.extra_detail_label = "General Ledger transactions with no match in either box"
    return result


def reconcile(data: dict, settings: VatReconSettings) -> list[ReconResult]:
    """data keys: vat_gl, vat_filed_sales, vat_filed_purchases (see
    main.py's _load_canonical_data - each is the concatenation of every
    confirmed upload of that type, tagged with a source_file column on
    the filed sides). Always returns three results: Box 1, Box 4, then
    General Ledger Coverage (whatever neither box's filed detail
    accounted for) - "n/a" for a box with nothing filed against it, or
    for coverage when there's no General Ledger at all."""
    pool = _prepare_pool(data.get("vat_gl"), settings)

    box1_matched, box1_unmatched_filed, pool = match_box(pool, data.get("vat_filed_sales"), settings)
    box4_matched, box4_unmatched_filed, pool = match_box(pool, data.get("vat_filed_purchases"), settings)

    return [
        _box_result("Box 1 (Sales)", data.get("vat_filed_sales"), box1_matched, box1_unmatched_filed, settings),
        _box_result("Box 4 (Purchases)", data.get("vat_filed_purchases"), box4_matched, box4_unmatched_filed, settings),
        _gl_coverage_result(data.get("vat_gl"), pool),
    ]


def reconcile_box(box_label: str, gl: pd.DataFrame | None, filed: pd.DataFrame | None, settings: VatReconSettings) -> ReconResult:
    """Reconciles a single box in isolation (no shared-pool consumption
    against another box) - used directly by tests and by any caller that
    only cares about one box. reconcile() above is what main.py actually
    calls, since it needs Box 1 and Box 4 to share one GL pool."""
    pool = _prepare_pool(gl, settings)
    matched, unmatched_filed, _ = match_box(pool, filed, settings)
    return _box_result(box_label, filed, matched, unmatched_filed, settings)
