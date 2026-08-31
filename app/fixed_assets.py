"""Fixed asset register.

Two ways to build this, usable independently or together:

1. Category-level, derived entirely from TB + nominal activity (always
   available, no extra upload needed). Real Xero-exported TBs commonly
   split each fixed asset category into a COST/ADDITIONS code and a
   DEPRECIATION/ACCUMULATED DEPRECIATION code (e.g. "COMPUTER EQUIPMENT -
   COST", "COMPUTER EQUIPMENT - DEPRECIATION B/FWD") - these are detected
   and paired up automatically into a cost / accumulated depreciation / NBV
   rollforward per category, cross-checked against the TB.

2. Asset-level, rolling forward a prior year's register (uploaded via the
   generic mapping path, since firms keep these in whatever format their
   own working papers use - there's no standard software export for this).
   Each asset's depreciation for the year is computed from its own
   method/rate, new additions are picked up from nominal activity against
   fixed asset cost codes, and the totals are reconciled to the TB.

Both produce a status/message/schedule shape consistent with the rest of
the recon engine, so the Excel builder treats them the same way.
"""
import re
from dataclasses import dataclass, field

import pandas as pd

from app.recon import ReconResult

MATERIALITY_AMOUNT = 500.0

FIXED_ASSET_TYPE = "fixed asset"
_DEPRECIATION_KEYWORDS = ("depreciation", "accumulated depn", "acc dep")

# Real charts of accounts name cost/depreciation sub-accounts very
# inconsistently - some use punctuated suffixes ("COMPUTER EQUIPMENT -
# COST"), others just run the words together with no separator at all
# ("IT EQUIPMENT COST BROUGHT FORWARD"). Stripping a fixed suffix pattern
# only caught the first style; stripping these "structural" tokens
# wherever they appear catches both, so the remaining words are the
# category name either way.
_STRUCTURAL_WORDS = re.compile(
    r"\b(cost|additions?|disposals?|depreciation|accumulated|accum\.?|acc\.?|"
    r"brought|forward|carried|b/?fwd|c/?fwd|nbv|net|book|value|of)\b",
    re.IGNORECASE,
)
_PUNCT_EDGES = re.compile(r"^[\s\-:]+|[\s\-:]+$")
_MULTI_SPACE = re.compile(r"\s{2,}")

STRAIGHT_LINE = "straight_line"
REDUCING_BALANCE = "reducing_balance"

# Default depreciation basis per category, inferred from the category name
# itself (keyword -> (method, annual rate %)) - a generic SME starting point
# so the category-level rollforward (which has no asset register to read an
# actual method/rate from) can still show a system-estimated depreciation
# charge next to what's actually booked. Deliberately never used to flag a
# category "review" on its own - real accounting policies vary too much for
# a keyword guess to be authoritative - only ever shown as an advisory
# column, with the assumption stated plainly in the result message.
DEFAULT_CATEGORY_DEPRECIATION: dict[str, tuple[str, float]] = {
    "computer": (REDUCING_BALANCE, 25.0),
    "it equipment": (REDUCING_BALANCE, 25.0),
    "software": (STRAIGHT_LINE, 33.3),
    "motor vehicle": (REDUCING_BALANCE, 25.0),
    "vehicle": (REDUCING_BALANCE, 25.0),
    "fixture": (STRAIGHT_LINE, 15.0),
    "fitting": (STRAIGHT_LINE, 15.0),
    "furniture": (STRAIGHT_LINE, 15.0),
    "plant": (REDUCING_BALANCE, 20.0),
    "machinery": (REDUCING_BALANCE, 20.0),
    "tool": (REDUCING_BALANCE, 20.0),
    "equipment": (REDUCING_BALANCE, 20.0),
    "leasehold improvement": (STRAIGHT_LINE, 10.0),
    "building": (STRAIGHT_LINE, 2.0),
    "land": (STRAIGHT_LINE, 0.0),
}
_DEFAULT_RATE_FALLBACK = (STRAIGHT_LINE, 20.0)


def _infer_category_rate(category: str) -> tuple[str, float]:
    """First keyword match wins - dict order above goes from most to least
    specific ("it equipment" before the bare "equipment" it would otherwise
    also match), so a specific category name gets the specific rate."""
    name = str(category).lower()
    for keyword, rate in DEFAULT_CATEGORY_DEPRECIATION.items():
        if keyword in name:
            return rate
    return _DEFAULT_RATE_FALLBACK


@dataclass
class FixedAssetResult:
    name: str
    status: str  # "ok" | "review" | "n/a"
    message: str
    detail: pd.DataFrame = field(default_factory=pd.DataFrame)
    # Same purpose as recon.ReconResult.extra_detail: a second, differently-
    # shaped supporting table written below `detail` on the same sheet -
    # here, the actual transactions behind each category's "Additions"
    # figure, so a preparer isn't left trusting a total with nothing to
    # check it against.
    extra_detail: pd.DataFrame = field(default_factory=pd.DataFrame)
    extra_detail_label: str = ""


@dataclass
class AssetRegisterResult:
    status: str  # "ok" | "review" | "n/a"
    message: str
    summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    asset_schedule: pd.DataFrame = field(default_factory=pd.DataFrame)
    new_additions: pd.DataFrame = field(default_factory=pd.DataFrame)
    possible_disposals: pd.DataFrame = field(default_factory=pd.DataFrame)
    closing_register: pd.DataFrame = field(default_factory=pd.DataFrame)
    period_fraction: float = 1.0  # the period_days/365 fraction depreciation was prorated by - exposed so a formula-linked sheet can embed the same constant a formula recalculates against, rather than re-deriving it


def _category_and_kind(account_name: str) -> tuple[str, str]:
    """Returns (category, 'cost'|'depreciation') for a Fixed Asset TB account."""
    kind = "depreciation" if any(k in account_name.lower() for k in _DEPRECIATION_KEYWORDS) else "cost"
    stripped = _STRUCTURAL_WORDS.sub(" ", account_name)
    stripped = _PUNCT_EDGES.sub("", _MULTI_SPACE.sub(" ", stripped).strip())
    return stripped or account_name, kind


def group_fixed_asset_codes(tb_current: pd.DataFrame) -> dict[str, dict[str, list[str]]]:
    """Groups Fixed Asset TB account codes by category ("IT EQUIPMENT", ...)
    and kind ("cost"/"depreciation"), the same grouping category_level_
    rollforward uses internally - exposed separately so the formula-linked
    Excel builder can reference these exact account codes in its own
    SUMPRODUCT formulas instead of re-deriving (and potentially drifting
    from) the grouping logic."""
    if tb_current is None or tb_current.empty:
        return {}
    fa = tb_current[tb_current["account_type"].str.lower() == FIXED_ASSET_TYPE].copy()
    if fa.empty:
        return {}
    fa["category"], fa["kind"] = zip(*fa["account_name"].map(_category_and_kind))
    grouped: dict[str, dict[str, list[str]]] = {}
    for category, group in fa.groupby("category"):
        grouped[category] = {
            "cost": group.loc[group["kind"] == "cost", "account_code"].astype(str).tolist(),
            "depreciation": group.loc[group["kind"] == "depreciation", "account_code"].astype(str).tolist(),
        }
    return grouped


# --- Genuine addition vs. opening-balance migration journal --------------
#
# A debit to a fixed asset cost code looks identical in nominal activity
# whether it's a real in-year purchase or a data-migration journal that
# simply re-established a prior system's opening balances in this one -
# found against real client data, where entries in a TB's own "Opening
# Balance" section got flagged as additions needing review, which is the
# right call (they DO need reviewing) but misleads a preparer into
# looking for a purchase invoice that was never posted. Both directions
# below use only signals the client's own posting habits already carry -
# same "client's own data as vocabulary" idea as this module's other
# suggestion checks (suggest_capital_expenditure_reclassification) -
# never an invented one:
#
#   1. the transaction's own description/reference text - a genuine
#      migration entry is nearly always labelled as one by whoever
#      created it ("Opening Balance", "B/Fwd", "Data Migration", ...).
#   2. failing that, its date and source together - a migration is
#      characteristically a single journal struck at/near the very start
#      of the period, not a Bill/Invoice/Spend Money transaction spread
#      through the year the way a real purchase is.
#
# Purely advisory labelling: never changes whether a row is counted,
# flagged, or included - only which of the two explanations it's shown
# with, so a preparer isn't misled about WHY something needs reviewing.
_MIGRATION_KEYWORDS = (
    "opening balance", "opening bal", "b/fwd", "bfwd", "brought forward",
    "data migration", "migration", "migrated", "conversion balance", "trial balance conversion",
)
_MIGRATION_JOURNAL_SOURCE_TYPES = ("journal", "manual journal", "je")
_MIGRATION_DATE_WINDOW_DAYS = 14
_LIKELY_GENUINE_ADDITION = "Likely genuine addition"
_POSSIBLE_MIGRATION = "Possible opening-balance migration - review before treating as a new asset"


def _addition_type(row: pd.Series, period_start) -> str:
    text = f"{row.get('description', '')} {row.get('reference', '')}".lower()
    if any(k in text for k in _MIGRATION_KEYWORDS):
        return _POSSIBLE_MIGRATION
    if period_start is not None and str(row.get("source_type", "")).lower() in _MIGRATION_JOURNAL_SOURCE_TYPES:
        posting_date = row.get("date")
        if pd.notna(posting_date):
            days_from_start = abs((pd.Timestamp(posting_date) - pd.Timestamp(period_start)).days)
            if days_from_start <= _MIGRATION_DATE_WINDOW_DAYS:
                return _POSSIBLE_MIGRATION
    return _LIKELY_GENUINE_ADDITION


def far_additions_detail(
    tb_current: pd.DataFrame | None, nominal_activity: pd.DataFrame | None, period_start=None,
) -> pd.DataFrame:
    """Every transaction-level debit posted during the year to a Fixed
    Asset cost code, across every category - the actual postings behind
    each category's "Additions" total in category_level_rollforward, so a
    preparer can check the figure against real transactions rather than
    just trust an aggregate. This is "first, suggest anything already
    posted to a FAR category as an addition" - the part of the register
    that needs no inference at all, since the client's own chart of
    accounts already says these are fixed assets."""
    if tb_current is None or tb_current.empty or nominal_activity is None or nominal_activity.empty:
        return pd.DataFrame()

    grouped = group_fixed_asset_codes(tb_current)
    if not grouped:
        return pd.DataFrame()

    code_to_category: dict[str, str] = {}
    for category, codes in grouped.items():
        for code in codes["cost"]:
            code_to_category[code] = category
    if not code_to_category:
        return pd.DataFrame()

    movement = nominal_activity.copy()
    movement["account_code"] = movement["account_code"].astype(str)
    movement = movement[movement["account_code"].isin(code_to_category)]
    additions = movement[movement["debit"].astype(float) > 0].copy()
    if additions.empty:
        return pd.DataFrame()

    additions["Category"] = additions["account_code"].map(code_to_category)
    for col in ("date", "account_name", "reference", "description", "contact", "source_type"):
        if col not in additions.columns:
            additions[col] = ""
    additions["Addition type"] = additions.apply(lambda r: _addition_type(r, period_start), axis=1)
    return additions[["date", "Category", "account_code", "account_name", "reference", "description", "contact", "debit", "Addition type"]].rename(columns={
        "date": "Date", "account_code": "Nominal Code", "account_name": "Account",
        "reference": "Reference", "description": "Description", "contact": "Contact", "debit": "Addition (Cost)",
    }).sort_values(["Category", "Date"]).reset_index(drop=True)


def category_level_rollforward(
    tb_current: pd.DataFrame, tb_comparative: pd.DataFrame, nominal_activity: pd.DataFrame | None,
    period_days: int = 365, materiality: float = MATERIALITY_AMOUNT, period_start=None,
) -> FixedAssetResult:
    """Cost / accumulated depreciation / NBV rollforward per fixed asset
    category, built purely from the TB's own cost-vs-depreciation account
    split - no fixed asset register upload required. Also brings in, per
    category: the actual addition transactions behind the "Additions"
    total (far_additions_detail, above) and a system-estimated
    depreciation charge (see DEFAULT_CATEGORY_DEPRECIATION) to sanity-
    check what's actually booked against - both stated explicitly in the
    result rather than left for the preparer to have to go find."""
    if tb_current is None or tb_current.empty:
        return FixedAssetResult("Fixed asset register (category summary)", "n/a", "No trial balance uploaded.")

    grouped = group_fixed_asset_codes(tb_current)
    if not grouped:
        return FixedAssetResult("Fixed asset register (category summary)", "n/a", "No Fixed Asset accounts found in the trial balance.")

    period_fraction = min(1.0, max(0.0, period_days / 365.0))

    def current_balance(code: str) -> float:
        mask = tb_current["account_code"].astype(str) == code
        return float(tb_current.loc[mask, "balance"].sum())

    def comparative_balance(code: str) -> float:
        if tb_comparative is None or tb_comparative.empty:
            return 0.0
        mask = tb_comparative["account_code"].astype(str) == code
        return float(tb_comparative.loc[mask, "balance"].sum())

    def movement(code: str) -> tuple[float, float]:
        if nominal_activity is None or nominal_activity.empty:
            return 0.0, 0.0
        m = nominal_activity[nominal_activity["account_code"].astype(str) == code]
        return float(m["debit"].sum()), float(m["credit"].sum())

    rows = []
    for category, codes in grouped.items():
        cost_codes = codes["cost"]
        dep_codes = codes["depreciation"]

        cost_c_fwd = sum(current_balance(c) for c in cost_codes)
        cost_b_fwd = sum(comparative_balance(c) for c in cost_codes)
        additions = sum(movement(c)[0] for c in cost_codes)
        disposals_cost = sum(movement(c)[1] for c in cost_codes)

        # accumulated depreciation is credit-natured (negative in this TB's
        # debit-positive convention) - flipped to a positive "amount of
        # depreciation" here so it reads naturally and subtracts correctly
        # from cost to give NBV, same treatment as debtors/creditors elsewhere
        acc_dep_c_fwd = -sum(current_balance(c) for c in dep_codes)
        acc_dep_b_fwd = -sum(comparative_balance(c) for c in dep_codes)
        dep_charge = sum(movement(c)[1] for c in dep_codes)
        dep_on_disposals = sum(movement(c)[0] for c in dep_codes)

        computed_cost_c_fwd = cost_b_fwd + additions - disposals_cost
        computed_acc_dep_c_fwd = acc_dep_b_fwd + dep_charge - dep_on_disposals
        cost_diff = round(computed_cost_c_fwd - cost_c_fwd, 2)
        dep_diff = round(computed_acc_dep_c_fwd - acc_dep_c_fwd, 2)

        est_method, est_rate_pct = _infer_category_rate(category)
        est_base = cost_b_fwd if est_method == STRAIGHT_LINE else max(0.0, cost_b_fwd - acc_dep_b_fwd)
        system_est_dep = round(max(0.0, est_base) * (est_rate_pct / 100.0) * period_fraction, 2)

        rows.append({
            "Category": category,
            "Cost b/fwd": round(cost_b_fwd, 2),
            "Additions": round(additions, 2),
            "Disposals (cost)": round(disposals_cost, 2),
            "Cost c/fwd (per TB)": round(cost_c_fwd, 2),
            "Cost diff": cost_diff,
            "Acc. depreciation b/fwd": round(acc_dep_b_fwd, 2),
            "Depreciation charge": round(dep_charge, 2),
            "Depreciation on disposals": round(dep_on_disposals, 2),
            "Acc. depreciation c/fwd (per TB)": round(acc_dep_c_fwd, 2),
            "Depreciation diff": dep_diff,
            "NBV b/fwd": round(cost_b_fwd - acc_dep_b_fwd, 2),
            "NBV c/fwd": round(cost_c_fwd - acc_dep_c_fwd, 2),
            "System est. method": "Reducing balance" if est_method == REDUCING_BALANCE else "Straight line",
            "System est. rate %": est_rate_pct,
            "System est. depreciation": system_est_dep,
            "Booked vs system est. (diff)": round(dep_charge - system_est_dep, 2),
        })

    detail = pd.DataFrame(rows)
    flagged = detail[(detail["Cost diff"].abs() > materiality) | (detail["Depreciation diff"].abs() > materiality)]
    status = "ok" if flagged.empty else "review"
    msg = (
        "Every category's cost and depreciation rollforward ties to the TB."
        if status == "ok"
        else f"{len(flagged)} categor{'y' if len(flagged) == 1 else 'ies'} don't tie to the TB - check for movements outside the nominal activity supplied (e.g. a revaluation or manual journal)."
    )
    msg += (
        " The 'System est.' columns apply a default depreciation rate inferred from each category's name "
        "(e.g. computer/IT equipment 25% reducing balance, motor vehicles 25% reducing balance, fixtures/"
        "fittings/furniture 15% straight line, 20% straight line otherwise) as a sanity-check against the "
        "booked charge - a starting assumption to verify, not this client's actual accounting policy."
    )

    result = FixedAssetResult("Fixed asset register (category summary)", status, msg, detail)
    additions_detail = far_additions_detail(tb_current, nominal_activity, period_start)
    if not additions_detail.empty:
        result.extra_detail = additions_detail
        result.extra_detail_label = "Fixed asset additions posted during the year (per nominal activity)"
    return result


_METHOD_ALIASES = {
    "sl": STRAIGHT_LINE, "straightline": STRAIGHT_LINE, "straight-line": STRAIGHT_LINE, "straight line": STRAIGHT_LINE,
    "rb": REDUCING_BALANCE, "reducingbalance": REDUCING_BALANCE, "reducing balance": REDUCING_BALANCE,
    "wdv": REDUCING_BALANCE, "written down value": REDUCING_BALANCE, "diminishing balance": REDUCING_BALANCE,
}


def _normalise_method(raw: str) -> str:
    key = str(raw or "").strip().lower()
    return _METHOD_ALIASES.get(key, STRAIGHT_LINE if not key else key)


_DISPOSED_TRUTHY = {"yes", "y", "true", "1", "disposed", "sold", "written off", "wrote off"}


def _parse_disposed_flag(value) -> bool:
    """A "Disposed?" column round-tripped through a real upload (CSV/xlsx
    -> parsers.apply_mapping) always arrives as a string, not a Python
    bool - "No" is exactly as truthy as "Yes" under bool("No"), so
    reg["disposed"].astype(bool) (the old approach) marked EVERY asset as
    disposed the moment a real file supplied a "No" column, silently
    emptying the asset schedule. Only recognised affirmative text (or an
    already-Python bool True) counts as disposed; "No", "", NaN and
    anything else read as not disposed."""
    if isinstance(value, bool):
        return value
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    return str(value).strip().lower() in _DISPOSED_TRUTHY


# --- Suggested asset ID for a possible disposal --------------------------
#
# Same "client's own data as vocabulary" approach as suggest_capital_
# expenditure_reclassification below and control_accounts.py's/
# nominal_matrix.py's own suggestion checks: a credit movement on a fixed
# asset cost code is a candidate disposal, but which specific still-held
# register line it corresponds to isn't knowable from the movement alone
# - the preparer has always had to read the posting's own description/
# reference text and compare it against the register by eye. This does
# exactly that comparison automatically: if a still-held asset's own
# description shows up (as significant words) in the disposal posting's
# own description/reference text, and no OTHER asset's description
# matches comparably well, that's a strong enough signal to suggest -
# never strong enough to act on, so it only ever fills in an advisory
# "Suggested asset ID" column, next to the "Matched to asset ID (to
# complete)" column the preparer still fills in themselves.
_DISPOSAL_MATCH_MIN_SCORE = 0.6  # same threshold as DOMINANT_SHARE_THRESHOLD elsewhere in this codebase
_DISPOSAL_WORD_RE = re.compile(r"[a-z0-9]+")
_DISPOSAL_STOPWORDS = {"and", "the", "for", "of", "a", "an", "to", "in", "on", "with"}


def _disposal_keywords(text) -> set[str]:
    words = _DISPOSAL_WORD_RE.findall(str(text).lower())
    return {w for w in words if len(w) >= 3 and w not in _DISPOSAL_STOPWORDS}


def _suggest_disposal_matches(possible_disposals: pd.DataFrame, still_held: pd.DataFrame) -> pd.DataFrame:
    """Adds "Suggested asset ID" / "Suggested match reason" columns to
    possible_disposals: for each disposal, the still-held register line
    whose own description's significant words appear (a real majority -
    at least _DISPOSAL_MATCH_MIN_SCORE of them) in the disposal's own
    description/reference text - and no other asset's description scores
    just as well. Left blank when nothing clears the threshold, or when
    two or more assets tie for the best score (e.g. two near-identical
    "Dell Latitude laptop" units - the posting alone can't say which one
    was actually disposed, so this never guesses between them)."""
    if possible_disposals.empty or still_held.empty:
        return possible_disposals

    assets = still_held[["asset_id", "description"]].copy()
    assets["_keywords"] = assets["description"].apply(_disposal_keywords)
    assets = assets[assets["_keywords"].apply(len) > 0]
    if assets.empty:
        return possible_disposals

    suggested_ids, reasons = [], []
    for _, row in possible_disposals.iterrows():
        text_keywords = _disposal_keywords(f"{row.get('Description', '')} {row.get('Reference', '')}")
        scores = []
        if text_keywords:
            for _, asset in assets.iterrows():
                overlap = asset["_keywords"] & text_keywords
                score = len(overlap) / len(asset["_keywords"])
                if score >= _DISPOSAL_MATCH_MIN_SCORE:
                    scores.append((asset["asset_id"], asset["description"], score))
        if scores:
            top_score = max(s[2] for s in scores)
            top_matches = [s for s in scores if s[2] == top_score]
            if len(top_matches) == 1:
                asset_id, description, score = top_matches[0]
                suggested_ids.append(asset_id)
                reasons.append(f"'{description}' matched in the posting's own description/reference")
                continue
        suggested_ids.append("")
        reasons.append("")

    possible_disposals = possible_disposals.copy()
    possible_disposals["Suggested asset ID"] = suggested_ids
    possible_disposals["Suggested match reason"] = reasons
    return possible_disposals


def asset_level_rollforward(
    prior_register: pd.DataFrame,
    nominal_activity: pd.DataFrame | None,
    tb_current: pd.DataFrame | None,
    period_days: int = 365,
    materiality: float = MATERIALITY_AMOUNT,
    period_start=None,
) -> AssetRegisterResult:
    """Rolls a prior-year asset register forward: depreciates each asset by
    its own method/rate, flags additions found in nominal activity that
    aren't yet in the register, flags likely disposals, and reconciles
    totals to the TB."""
    if prior_register is None or prior_register.empty:
        return AssetRegisterResult("n/a", "No prior year fixed asset register uploaded.")

    reg = prior_register.copy()
    reg["depreciation_method"] = reg["depreciation_method"].map(_normalise_method)
    reg["disposed"] = reg.get("disposed", False).apply(_parse_disposed_flag) if "disposed" in reg.columns else False
    period_fraction = min(1.0, max(0.0, period_days / 365.0))

    def charge_for(row) -> float:
        rate = float(row["depreciation_rate"]) / 100.0 if row["depreciation_rate"] else 0.0
        base = row["cost"] if row["depreciation_method"] == STRAIGHT_LINE else (row["cost"] - row["accumulated_depreciation_b_fwd"])
        return round(max(0.0, base) * rate * period_fraction, 2)

    reg["depreciation_charge"] = reg.apply(charge_for, axis=1)
    reg["accumulated_depreciation_c_fwd"] = (reg["accumulated_depreciation_b_fwd"] + reg["depreciation_charge"]).clip(upper=reg["cost"])
    reg["nbv_b_fwd"] = reg["cost"] - reg["accumulated_depreciation_b_fwd"]
    reg["nbv_c_fwd"] = reg["cost"] - reg["accumulated_depreciation_c_fwd"]

    still_held = reg[~reg["disposed"]]
    asset_schedule = still_held[[
        "asset_id", "description", "category", "date_acquired", "cost", "depreciation_method", "depreciation_rate",
        "accumulated_depreciation_b_fwd", "depreciation_charge", "accumulated_depreciation_c_fwd", "nbv_b_fwd", "nbv_c_fwd",
    ]].rename(columns={
        "asset_id": "Asset ID", "description": "Description", "category": "Category", "date_acquired": "Date Acquired",
        "cost": "Cost", "depreciation_method": "Method", "depreciation_rate": "Rate %",
        "accumulated_depreciation_b_fwd": "Acc. Dep. b/fwd", "depreciation_charge": "Depreciation Charge",
        "accumulated_depreciation_c_fwd": "Acc. Dep. c/fwd", "nbv_b_fwd": "NBV b/fwd", "nbv_c_fwd": "NBV c/fwd",
    })

    new_additions = pd.DataFrame()
    possible_disposals = pd.DataFrame()
    if nominal_activity is not None and not nominal_activity.empty and tb_current is not None and not tb_current.empty:
        fa_accounts = tb_current[tb_current["account_type"].str.lower() == FIXED_ASSET_TYPE]
        cost_accounts = fa_accounts[~fa_accounts["account_name"].str.lower().str.contains("depreciation")]
        fa_codes = set(cost_accounts["account_code"].astype(str))
        movement = nominal_activity[nominal_activity["account_code"].astype(str).isin(fa_codes)]
        if not movement.empty:
            additions_raw = movement[movement["debit"] > 0]
            if not additions_raw.empty:
                addition_types = additions_raw.apply(lambda r: _addition_type(r, period_start), axis=1)
                new_additions = additions_raw[["date", "account_code", "account_name", "reference", "description", "contact", "debit"]].rename(
                    columns={"debit": "Cost", "account_name": "Account", "date": "Date", "reference": "Reference", "description": "Description", "contact": "Contact", "account_code": "Nominal Code"}
                )
                new_additions["Addition type"] = addition_types.values
                new_additions["Category (to complete)"] = ""
                new_additions["Depreciation rate % (to complete)"] = ""

            disposals_raw = movement[movement["credit"] > 0]
            if not disposals_raw.empty:
                possible_disposals = disposals_raw[["date", "account_code", "account_name", "reference", "description", "contact", "credit"]].rename(
                    columns={"credit": "Cost removed", "account_name": "Account", "date": "Date", "reference": "Reference", "description": "Description", "contact": "Contact", "account_code": "Nominal Code"}
                )
                possible_disposals["Matched to asset ID (to complete)"] = ""
                possible_disposals = _suggest_disposal_matches(possible_disposals, still_held)

    # Closing register: same canonical shape as the prior-year upload, so
    # next year's job can take this sheet's contents straight back in as
    # its opening register. Existing assets carry forward with this year's
    # closing accumulated depreciation as next year's brought-forward
    # figure; new additions are appended with cost known but category/
    # method/rate left for the preparer to complete before the next roll.
    closing_existing = still_held[["asset_id", "description", "category", "date_acquired", "cost", "depreciation_method", "depreciation_rate"]].copy()
    closing_existing["accumulated_depreciation_b_fwd"] = still_held["accumulated_depreciation_c_fwd"]
    closing_existing["disposed"] = False

    closing_new = pd.DataFrame()
    if not new_additions.empty:
        closing_new = pd.DataFrame({
            "asset_id": [f"NEW-{i + 1} (assign a proper ID)" for i in range(len(new_additions))],
            "description": new_additions["Description"].values,
            "category": "",
            "date_acquired": new_additions["Date"].values,
            "cost": new_additions["Cost"].values,
            "depreciation_method": "",
            "depreciation_rate": "",
            "accumulated_depreciation_b_fwd": 0.0,
            "disposed": False,
        })
    closing_register = pd.concat([closing_existing, closing_new], ignore_index=True).rename(columns={
        "asset_id": "Asset ID", "description": "Description", "category": "Category (complete for new additions)",
        "date_acquired": "Date Acquired", "cost": "Cost", "depreciation_method": "Depreciation Method (complete for new additions)",
        "depreciation_rate": "Depreciation Rate % (complete for new additions)",
        "accumulated_depreciation_b_fwd": "Accumulated Depreciation b/fwd (for next year's opening register)",
        "disposed": "Disposed?",
    })

    total_cost = float(still_held["cost"].sum()) + float(new_additions["Cost"].sum() if not new_additions.empty else 0)
    total_acc_dep = float(still_held["accumulated_depreciation_c_fwd"].sum())
    total_nbv = total_cost - total_acc_dep

    tb_nbv = None
    if tb_current is not None and not tb_current.empty:
        fa = tb_current[tb_current["account_type"].str.lower() == FIXED_ASSET_TYPE]
        if not fa.empty:
            tb_nbv = float(fa["balance"].sum())

    variance = None
    status, message = "n/a", "Register rolled forward - no TB fixed asset total available to compare against."
    if tb_nbv is not None:
        variance = round(total_nbv - tb_nbv, 2)
        parts = []
        if abs(variance) > materiality:
            parts.append(f"register NBV (£{total_nbv:,.2f}) differs from the TB fixed asset net balance (£{tb_nbv:,.2f}) by £{abs(variance):,.2f}")
        if not new_additions.empty:
            parts.append(f"{len(new_additions)} transaction(s) found in nominal activity against fixed asset cost codes not yet in the register - review and add")
        if not possible_disposals.empty:
            parts.append(f"{len(possible_disposals)} credit movement(s) on fixed asset cost codes - possible disposals to match and mark in the register")
        status = "ok" if not parts else "review"
        message = "Register ties to the TB, no new additions or disposals detected." if not parts else "; ".join(parts).capitalize() + "."

    summary = pd.DataFrame([{
        "Total cost (register + unrecorded additions)": round(total_cost, 2),
        "Total accumulated depreciation": round(total_acc_dep, 2),
        "Total NBV (register)": round(total_nbv, 2),
        "TB fixed asset net balance": round(tb_nbv, 2) if tb_nbv is not None else "",
        "Variance": variance if variance is not None else "",
    }])

    return AssetRegisterResult(status, message, summary, asset_schedule, new_additions, possible_disposals, closing_register, period_fraction)


# --- Capital expenditure coded elsewhere ---------------------------------
#
# Everything above starts from what the client has already coded to a
# Fixed Asset account. This section looks the other way: nominal activity
# coded to an ordinary expense account that, based on THIS client's own
# fixed asset categories (and common capex nouns), looks like it could be
# capital expenditure someone miscoded - a laptop put through "IT Costs"
# instead of "Computer Equipment", a van through "Repairs & Maintenance"
# instead of "Motor Vehicles". The vocabulary is built from the client's
# own register/TB categories wherever possible - "nature of business" as
# reflected in what this client actually owns already, not a generic
# industry guess - so a garage's existing "Workshop Equipment" category
# is what catches its own miscoded tool purchases, and a different
# client's categories catch different things. Never reclassifies
# anything automatically: only ever a suggestion for the preparer to
# check, since a keyword match on a description is a starting point, not
# proof of capital nature.

_EXPENSE_LIKE_TYPES = {
    "overhead", "overheads", "expense", "expenses", "direct costs", "direct cost",
    "cost of sales", "administrative expenses", "admin expenses", "operating expenses",
}

_GENERIC_CAPEX_KEYWORDS = {
    "laptop", "macbook", "computer", "desktop", "printer", "photocopier", "server", "monitor", "projector",
    "van", "vehicle", "car", "forklift", "trailer",
    "machinery", "machine", "plant", "generator", "tool", "tools",
    "furniture", "desk", "chair", "chairs", "cabinet", "shelving",
    "renovation", "refurbishment", "refurb", "fit-out", "fitout", "extension", "installation",
}

_VOCAB_STOPWORDS = {"and", "the", "for", "other", "general", "misc", "miscellaneous", "office", "assets", "asset"}
_VOCAB_WORD_RE = re.compile(r"[a-z0-9]+")


def _category_keywords(category: str) -> set[str]:
    words = _VOCAB_WORD_RE.findall(str(category).lower())
    return {w for w in words if len(w) >= 4 and w not in _VOCAB_STOPWORDS}


def _capex_vocabulary(tb_current: pd.DataFrame | None, prior_register: pd.DataFrame | None) -> set[str]:
    vocabulary = set(_GENERIC_CAPEX_KEYWORDS)
    for category in group_fixed_asset_codes(tb_current):
        vocabulary |= _category_keywords(category)
    if prior_register is not None and not prior_register.empty and "category" in prior_register.columns:
        for category in prior_register["category"].dropna().unique():
            vocabulary |= _category_keywords(str(category))
    return vocabulary


def suggest_capital_expenditure_reclassification(
    tb_current: pd.DataFrame | None,
    nominal_activity: pd.DataFrame | None,
    prior_register: pd.DataFrame | None = None,
    threshold: float = MATERIALITY_AMOUNT,
) -> ReconResult:
    name = "Fixed asset register - possible capital expenditure coded elsewhere"
    if tb_current is None or tb_current.empty or nominal_activity is None or nominal_activity.empty:
        return ReconResult(name, "n/a", "No trial balance / nominal activity available to scan.")

    # _capex_vocabulary always seeds with _GENERIC_CAPEX_KEYWORDS, so this
    # is never empty even for a client with no fixed asset categories yet
    # and no prior register - there's always at least the generic capex
    # vocabulary to check expense postings against.
    vocabulary = _capex_vocabulary(tb_current, prior_register)

    tb = tb_current.copy()
    tb["account_code"] = tb["account_code"].astype(str)
    type_by_code = dict(zip(tb["account_code"], tb["account_type"].astype(str).str.lower().str.strip()))

    candidates = nominal_activity.copy()
    candidates["account_code"] = candidates["account_code"].astype(str)
    candidates["_type"] = candidates["account_code"].map(type_by_code).fillna("")
    candidates = candidates[candidates["_type"].isin(_EXPENSE_LIKE_TYPES)]
    candidates = candidates[candidates["debit"].astype(float) > threshold]

    def matched_keywords(row) -> list[str]:
        text = f"{row.get('description', '')} {row.get('account_name', '')} {row.get('reference', '')}".lower()
        return sorted(kw for kw in vocabulary if kw in text)

    ok_message = (
        f"No expense postings above £{threshold:,.2f} matched this client's fixed asset vocabulary "
        f"({len(vocabulary)} keyword(s) drawn from its own fixed asset categories and common capex terms)."
    )
    if candidates.empty:
        return ReconResult(name, "ok", ok_message)

    candidates["_matches"] = candidates.apply(matched_keywords, axis=1)
    flagged = candidates[candidates["_matches"].map(bool)].copy()
    if flagged.empty:
        return ReconResult(name, "ok", ok_message)

    flagged["Matched on"] = flagged["_matches"].map(lambda ks: ", ".join(ks))
    for col in ("date", "account_name", "reference", "description", "contact"):
        if col not in flagged.columns:
            flagged[col] = ""
    detail = flagged[["date", "account_code", "account_name", "reference", "description", "contact", "debit", "Matched on"]].rename(columns={
        "date": "Date", "account_code": "Nominal Code", "account_name": "Account (currently coded to)",
        "reference": "Reference", "description": "Description", "contact": "Contact", "debit": "Amount",
    }).sort_values("Date").reset_index(drop=True)

    msg = (
        f"{len(detail)} posting(s) coded to a non-fixed-asset account look like they could be capital "
        f"expenditure, based on this client's own fixed asset categories and common capex terms - review "
        f"and add to the register if genuinely capital in nature. Nothing here is moved automatically; a "
        f"word match on the description is a starting point, not proof."
    )
    return ReconResult(name, "review", msg, detail)
