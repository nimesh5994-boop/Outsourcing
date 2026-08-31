"""VAT Reconciliation - General Ledger vs Filed VAT Return matching.

Box 1 (VAT due on sales - output tax) and Box 4 (VAT reclaimed on
purchases - input tax) are reconciled independently against the same
General Ledger transaction pool, Box 1 first then Box 4, through five
cascading passes per box - strongest, least ambiguous match first:

  1. reference (invoice/bill number) alone - not gated on amount, since
     the point of matching by invoice number is to catch a genuinely
     wrong VAT amount posted against a real invoice; that variance is
     then reported (see _review_rows), not hidden by requiring the
     amounts to already agree before the two sides count as the same
     transaction.
  2. date + VAT amount within tolerance + contact
  3. VAT amount within tolerance + contact (weakest one-to-one pass -
     flagged "verify")
  4. (cash basis only) a filed row against a COMBINATION of several
     remaining GL rows, same contact, summing to the filed VAT amount -
     the common cash-accounting shape where one invoice's VAT was
     recognised piecemeal in the GL as it was actually paid in
     instalments, while the filed return still reports it as one line.
  5. (cash basis only, run once after every filed row above has had its
     turn) the mirror image: a GL row against a COMBINATION of several
     still-unmatched filed rows, same contact, summing to the GL VAT
     amount - the return was filed at a finer grain than the GL posted
     it. Passes 4 and 5 never both claim the same row: 4 runs first and
     removes whatever it claims from the shared pool before 5 ever
     looks at what's left. See _find_combination for the (bounded,
     never-freezes-the-browser-equivalent) search.

Each GL row is claimed by at most one filed-return row (or filed-row
combination) across BOTH boxes: Box 1 matches first and removes what
it claims from the shared pool before Box 4 runs, so a sales
transaction can't also get claimed as a purchase. Whatever GL activity
neither box's filed return accounts for is reported once, as its own
"General Ledger Coverage" check, rather than being reported as an
"exception" under both boxes independently - a purchase transaction is
not a Box 1 exception just because Box 1's filed sales never mention
it, and vice versa.

"VAT Accounting Basis" (accrual vs cash) picks which General Ledger
date is used for the date-based pass: accrual uses the transaction/
invoice date (always present); cash uses a payment date when the GL
export actually carries one, falling back to the transaction date
row-by-row when it doesn't (a GL without payment dates can't be cash-
matched more precisely than that, so it degrades rather than drops
rows). Combination matching (passes 4/5) is cash-basis only - under
accrual accounting a genuine invoice-level VAT amount should already
be a single figure on both sides, so a split amount is more likely a
real discrepancy than a legitimate combination, and silently combining
it away would hide that.

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

# Standard UK VAT rates (zero/reduced/standard) - used only as an advisory
# sanity-check on matched pairs (see _implied_rate below), never to flag a
# box "review" by itself: a partial-exemption or flat-rate-scheme client
# can legitimately post other rates, so this is a heads-up to verify, not
# a finding.
STANDARD_VAT_RATES = (0.0, 0.05, 0.20)
RATE_TOLERANCE = 0.005

# Same thresholds nominal_matrix.py uses for its own contact-history
# suggestion, kept independent per module rather than shared, so each
# domain's suggestion logic can be tuned on its own terms later.
MIN_HISTORY_FOR_SUGGESTION = 2
DOMINANT_SHARE_THRESHOLD = 0.6

# Cash-basis combination matching (see match_box passes 4/5): bounded on
# every axis so a large same-contact cluster can't turn into a
# combinatorial explosion - only the rows closest in amount to the target
# are even considered, combinations are capped at a practical size (a
# genuine split payment is a handful of instalments, not dozens), and a
# hard ceiling on how many candidate combinations get examined guarantees
# the search always terminates quickly regardless of input.
MAX_COMBINATION_SIZE = 8
MAX_COMBINATION_CANDIDATES = 25
MAX_COMBINATIONS_EXAMINED = 20_000

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


def _implied_rate(net, vat) -> float | None:
    """VAT / net from the actual GL posting - None when there's no net
    amount to divide by (a rate is meaningless then, not zero)."""
    net = float(net)
    if abs(net) < 0.01:
        return None
    return round(abs(float(vat)) / abs(net), 4)


def _is_standard_rate(rate: float | None) -> bool:
    if rate is None:
        return True  # nothing to check a rate against - not itself a finding
    return any(abs(rate - r) <= RATE_TOLERANCE for r in STANDARD_VAT_RATES)


def _join_unique(values) -> str:
    """Semicolon-joined, sorted, de-duplicated text values - used to
    summarise a reference/source-file column across several combined
    rows into one display string, dropping blanks/NaN."""
    seen = sorted({str(v).strip() for v in values if str(v).strip() and str(v).strip().lower() != "nan"})
    return "; ".join(seen)


def _describe_dates(dates) -> object:
    """A single date if every combined row shares one, otherwise a short
    'N dates (earliest to latest)' summary - mixing a real date (for the
    common single-date case) with a text summary (for a genuine spread)
    in the same column is fine for display; Excel just shows the text
    ones as text."""
    unique = sorted(pd.Series(list(dates)).dropna().unique())
    if not unique:
        return ""
    if len(unique) == 1:
        return unique[0]
    return f"{len(unique)} dates ({unique[0]} to {unique[-1]})"


def _find_combination(
    candidates: pd.DataFrame, contact_key_col: str, contact_key: str,
    amount_col: str, target: float, tolerance: float, id_col: str | None = None,
) -> list | None:
    """Bounded search for a small subset of `candidates` rows (same
    contact, `amount_col` summing to `target` within `tolerance`) - the
    cash-basis combination-matching primitive both match_box passes 4
    and 5 share. Returns the list of matching row ids (>=2 - a size-1
    match would already have been found by a one-to-one pass before this
    ever runs), or None if nothing valid turns up within the search
    bounds. `id_col` names the column identifying each row (e.g. the GL
    pool's `_pool_id`); when omitted, the DataFrame's own index is used
    (the filed side, where the index is already a stable per-row id).

    Bounded on every axis so this can never run away on a large same-
    contact cluster: only the MAX_COMBINATION_CANDIDATES rows closest in
    amount to the target are considered at all, combinations are capped
    at MAX_COMBINATION_SIZE rows, sizes are tried smallest-first with an
    immediate return on the first valid combination found (a genuine
    split payment is rarely more than a handful of instalments, so the
    common case resolves fast), and a hard ceiling on combinations
    examined guarantees termination regardless of input shape."""
    same_contact = candidates[
        (candidates[contact_key_col] == contact_key) & (candidates[amount_col].round(2) != 0)
    ].copy()
    if len(same_contact) < 2:
        return None

    same_contact["_diff"] = (same_contact[amount_col] - target).abs()
    same_contact = same_contact.sort_values("_diff").head(MAX_COMBINATION_CANDIDATES)
    ids = same_contact[id_col].tolist() if id_col else list(same_contact.index)
    rows = list(zip(ids, same_contact[amount_col].tolist()))

    tol = tolerance + 1e-9
    target = round(float(target), 2)
    budget = [MAX_COMBINATIONS_EXAMINED]

    def search(size: int, start: int, chosen: list, total: float) -> list | None:
        if len(chosen) == size:
            return list(chosen) if abs(round(total, 2) - target) <= tol else None
        remaining_slots = size - len(chosen)
        for i in range(start, len(rows) - remaining_slots + 1):
            if budget[0] <= 0:
                return None
            budget[0] -= 1
            rid, amount = rows[i]
            result = search(size, i + 1, chosen + [rid], total + amount)
            if result is not None:
                return result
        return None

    for size in range(2, min(MAX_COMBINATION_SIZE, len(rows)) + 1):
        result = search(size, 0, [], 0.0)
        if result is not None:
            return result
    return None


def _combined_gl_legs_row(frow: pd.Series, combo_rows: pd.DataFrame) -> dict:
    """One filed row matched against several GL rows summing to it (e.g.
    an invoice's VAT recognised piecemeal across instalment payments in
    the GL, filed as one line)."""
    return {
        "Match basis": f"combined GL postings, cash basis ({len(combo_rows)} legs)",
        "Filed date": frow["date"], "Filed reference": frow["reference"], "Filed contact": frow["contact"],
        "Filed net": frow["net_amount"], "Filed VAT": frow["vat_amount"], "Source file": frow.get("source_file", ""),
        "GL date": _describe_dates(combo_rows["date"]), "GL reference": _join_unique(combo_rows["reference"]),
        "GL contact": combo_rows.iloc[0]["contact"],
        "GL net": round(float(combo_rows["net_amount"].sum()), 2), "GL VAT": round(float(combo_rows["vat_amount"].sum()), 2),
        "VAT variance": round(float(frow["vat_amount"]) - float(combo_rows["vat_amount"].sum()), 2),
    }


def _combined_filed_items_row(grow: pd.Series, combo_filed: pd.DataFrame) -> dict:
    """One GL row matched against several filed rows summing to it (the
    mirror case: the return was filed at a finer grain than the GL
    posted it)."""
    filed_vat_total = float(combo_filed["vat_amount"].sum())
    return {
        "Match basis": f"combined filed items, cash basis ({len(combo_filed)} items)",
        "Filed date": _describe_dates(combo_filed["date"]), "Filed reference": _join_unique(combo_filed["reference"]),
        "Filed contact": combo_filed.iloc[0]["contact"],
        "Filed net": round(float(combo_filed["net_amount"].sum()), 2), "Filed VAT": round(filed_vat_total, 2),
        "Source file": _join_unique(combo_filed["source_file"]) if "source_file" in combo_filed.columns else "",
        "GL date": grow["date"], "GL reference": grow["reference"], "GL contact": grow["contact"],
        "GL net": grow["net_amount"], "GL VAT": grow["vat_amount"],
        "VAT variance": round(filed_vat_total - float(grow["vat_amount"]), 2),
    }


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

        # Pass 4 (cash basis only): this filed row against a COMBINATION
        # of several remaining GL rows summing to it - see module
        # docstring and _find_combination.
        if candidates is None and settings.accounting_basis == "cash":
            combo_ids = _find_combination(
                pool, "_contact_key", frow["_contact_key"], "vat_amount", float(frow["vat_amount"]),
                settings.tolerance, id_col="_pool_id",
            )
            if combo_ids:
                combo_rows = pool[pool["_pool_id"].isin(combo_ids)]
                pool = pool[~pool["_pool_id"].isin(combo_ids)]
                matched_rows.append(_combined_gl_legs_row(frow, combo_rows))
                continue

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

    # Pass 5 (cash basis only, runs once after every filed row above has
    # had its turn): the mirror image of pass 4 - a still-unclaimed GL
    # row against a COMBINATION of several still-unmatched filed rows
    # summing to it. Deliberately sequenced after the whole pass 1-4 loop
    # (not interleaved with it) so pass 4 always gets first claim on any
    # row a combination could explain either way, keeping the outcome
    # deterministic.
    if settings.accounting_basis == "cash" and unmatched_filed_idx and not pool.empty:
        remaining_filed = filed.loc[unmatched_filed_idx].copy()
        for _, grow in list(pool.iterrows()):
            if remaining_filed.empty:
                break
            combo_filed_idx = _find_combination(
                remaining_filed, "_contact_key", grow["_contact_key"], "vat_amount", float(grow["vat_amount"]),
                settings.tolerance,
            )
            if combo_filed_idx:
                combo_filed_rows = remaining_filed.loc[combo_filed_idx]
                matched_rows.append(_combined_filed_items_row(grow, combo_filed_rows))
                pool = pool[pool["_pool_id"] != grow["_pool_id"]]
                remaining_filed = remaining_filed.drop(index=combo_filed_idx)
        unmatched_filed_idx = list(remaining_filed.index)

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

    matched_out = matched.copy()
    non_standard_count = 0
    if not matched_out.empty:
        rates = matched_out.apply(lambda r: _implied_rate(r["GL net"], r["GL VAT"]), axis=1)
        matched_out["Implied VAT rate %"] = rates.apply(lambda r: round(r * 100, 1) if r is not None else "")
        non_standard_count = int(rates.apply(lambda r: not _is_standard_rate(r)).sum())
    if non_standard_count:
        msg += (
            f" {non_standard_count} matched item(s) have an implied VAT rate (VAT ÷ net, from the actual GL "
            f"posting) that isn't a standard UK rate (0%/5%/20%) - see 'Implied VAT rate %' in the matched detail "
            f"below; advisory only; a partial-exemption or flat-rate-scheme client can legitimately post other rates."
        )

    result = ReconResult(name, status, msg, detail)
    review = _box_review_rows(unmatched_filed, matched, variance_threshold)
    if not review.empty:
        result.extra_detail = review
        result.extra_detail_label = f"{box_label} - items requiring review"
    if not matched_out.empty:
        result.matched_detail = matched_out
        result.matched_detail_label = f"{box_label} - full matched detail"
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


def _suggest_box_for_gl_coverage_gap(remaining_pool: pd.DataFrame, box1_matched: pd.DataFrame, box4_matched: pd.DataFrame) -> ReconResult:
    """For a General Ledger transaction that matched neither box (see
    _gl_coverage_result), suggests which box it's probably missing from -
    based on that SAME contact's other GL transactions that DID match,
    this client's own pattern rather than a generic guess, the same
    "client's own data as vocabulary" approach as fixed_assets/
    control_accounts/nominal_matrix's own suggestion checks. Only
    suggested when one box clearly dominates that contact's matched
    history (a real majority, over a real sample) - never matches
    anything itself."""
    name = "VAT Recon - suggested box for unmatched General Ledger items"
    unmatched = remaining_pool[GL_COLUMNS] if remaining_pool is not None and not remaining_pool.empty else pd.DataFrame(columns=GL_COLUMNS)
    if unmatched.empty:
        return ReconResult(name, "n/a", "No General Ledger items are unmatched to either box.")

    box1_counts = box1_matched["GL contact"].apply(_normalise_contact).value_counts() if not box1_matched.empty else pd.Series(dtype=int)
    box4_counts = box4_matched["GL contact"].apply(_normalise_contact).value_counts() if not box4_matched.empty else pd.Series(dtype=int)

    rows = []
    for _, row in unmatched.iterrows():
        key = _normalise_contact(row["contact"])
        if not key:
            continue
        b1, b4 = int(box1_counts.get(key, 0)), int(box4_counts.get(key, 0))
        total = b1 + b4
        if total < MIN_HISTORY_FOR_SUGGESTION:
            continue
        if b1 / total >= DOMINANT_SHARE_THRESHOLD:
            suggested, based_on = "Box 1 (Sales)", b1
        elif b4 / total >= DOMINANT_SHARE_THRESHOLD:
            suggested, based_on = "Box 4 (Purchases)", b4
        else:
            continue
        rows.append({
            "Date": row["date"], "Reference": row["reference"], "Contact": row["contact"],
            "Net Amount": row["net_amount"], "VAT Amount": row["vat_amount"],
            "Suggested box": suggested,
            "Based on": f"{based_on} of {total} of this contact's other GL transactions matched into that box",
        })

    if not rows:
        return ReconResult(name, "ok", "No unmatched General Ledger item had a clear enough pattern in that contact's other transactions to suggest a box.")

    detail = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
    msg = (
        f"{len(detail)} unmatched General Ledger item(s) have a likely box based on that contact's own other "
        f"transactions - a suggestion to check, not an automatic match."
    )
    return ReconResult(name, "review", msg, detail)


def reconcile(data: dict, settings: VatReconSettings) -> list[ReconResult]:
    """data keys: vat_gl, vat_filed_sales, vat_filed_purchases (see
    main.py's _load_canonical_data - each is the concatenation of every
    confirmed upload of that type, tagged with a source_file column on
    the filed sides). Always returns four results: Box 1, Box 4, General
    Ledger Coverage (whatever neither box's filed detail accounted for),
    then the suggested-box check for those coverage gaps - "n/a" for a
    box with nothing filed against it, or for coverage/suggestions when
    there's no General Ledger at all."""
    pool = _prepare_pool(data.get("vat_gl"), settings)

    box1_matched, box1_unmatched_filed, pool = match_box(pool, data.get("vat_filed_sales"), settings)
    box4_matched, box4_unmatched_filed, pool = match_box(pool, data.get("vat_filed_purchases"), settings)

    return [
        _box_result("Box 1 (Sales)", data.get("vat_filed_sales"), box1_matched, box1_unmatched_filed, settings),
        _box_result("Box 4 (Purchases)", data.get("vat_filed_purchases"), box4_matched, box4_unmatched_filed, settings),
        _gl_coverage_result(data.get("vat_gl"), pool),
        _suggest_box_for_gl_coverage_gap(pool, box1_matched, box4_matched),
    ]


def reconcile_box(box_label: str, gl: pd.DataFrame | None, filed: pd.DataFrame | None, settings: VatReconSettings) -> ReconResult:
    """Reconciles a single box in isolation (no shared-pool consumption
    against another box) - used directly by tests and by any caller that
    only cares about one box. reconcile() above is what main.py actually
    calls, since it needs Box 1 and Box 4 to share one GL pool."""
    pool = _prepare_pool(gl, settings)
    matched, unmatched_filed, _ = match_box(pool, filed, settings)
    return _box_result(box_label, filed, matched, unmatched_filed, settings)
