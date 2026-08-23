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


@dataclass
class FixedAssetResult:
    name: str
    status: str  # "ok" | "review" | "n/a"
    message: str
    detail: pd.DataFrame = field(default_factory=pd.DataFrame)


@dataclass
class AssetRegisterResult:
    status: str  # "ok" | "review" | "n/a"
    message: str
    summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    asset_schedule: pd.DataFrame = field(default_factory=pd.DataFrame)
    new_additions: pd.DataFrame = field(default_factory=pd.DataFrame)
    possible_disposals: pd.DataFrame = field(default_factory=pd.DataFrame)
    closing_register: pd.DataFrame = field(default_factory=pd.DataFrame)


def _category_and_kind(account_name: str) -> tuple[str, str]:
    """Returns (category, 'cost'|'depreciation') for a Fixed Asset TB account."""
    kind = "depreciation" if any(k in account_name.lower() for k in _DEPRECIATION_KEYWORDS) else "cost"
    stripped = _STRUCTURAL_WORDS.sub(" ", account_name)
    stripped = _PUNCT_EDGES.sub("", _MULTI_SPACE.sub(" ", stripped).strip())
    return stripped or account_name, kind


def category_level_rollforward(
    tb_current: pd.DataFrame, tb_comparative: pd.DataFrame, nominal_activity: pd.DataFrame | None
) -> FixedAssetResult:
    """Cost / accumulated depreciation / NBV rollforward per fixed asset
    category, built purely from the TB's own cost-vs-depreciation account
    split - no fixed asset register upload required."""
    if tb_current is None or tb_current.empty:
        return FixedAssetResult("Fixed asset register (category summary)", "n/a", "No trial balance uploaded.")

    fa = tb_current[tb_current["account_type"].str.lower() == FIXED_ASSET_TYPE].copy()
    if fa.empty:
        return FixedAssetResult("Fixed asset register (category summary)", "n/a", "No Fixed Asset accounts found in the trial balance.")

    fa["category"], fa["kind"] = zip(*fa["account_name"].map(_category_and_kind))
    codes = set(fa["account_code"].astype(str))

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
    for category, group in fa.groupby("category"):
        cost_rows = group[group["kind"] == "cost"]
        dep_rows = group[group["kind"] == "depreciation"]

        cost_c_fwd = float(cost_rows["balance"].sum())
        cost_b_fwd = sum(comparative_balance(c) for c in cost_rows["account_code"].astype(str))
        additions = sum(movement(c)[0] for c in cost_rows["account_code"].astype(str))
        disposals_cost = sum(movement(c)[1] for c in cost_rows["account_code"].astype(str))

        # accumulated depreciation is credit-natured (negative in this TB's
        # debit-positive convention) - flipped to a positive "amount of
        # depreciation" here so it reads naturally and subtracts correctly
        # from cost to give NBV, same treatment as debtors/creditors elsewhere
        acc_dep_c_fwd = -float(dep_rows["balance"].sum())
        acc_dep_b_fwd = -sum(comparative_balance(c) for c in dep_rows["account_code"].astype(str))
        dep_charge = sum(movement(c)[1] for c in dep_rows["account_code"].astype(str))
        dep_on_disposals = sum(movement(c)[0] for c in dep_rows["account_code"].astype(str))

        computed_cost_c_fwd = cost_b_fwd + additions - disposals_cost
        computed_acc_dep_c_fwd = acc_dep_b_fwd + dep_charge - dep_on_disposals
        cost_diff = round(computed_cost_c_fwd - cost_c_fwd, 2)
        dep_diff = round(computed_acc_dep_c_fwd - acc_dep_c_fwd, 2)

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
        })

    detail = pd.DataFrame(rows)
    flagged = detail[(detail["Cost diff"].abs() > MATERIALITY_AMOUNT) | (detail["Depreciation diff"].abs() > MATERIALITY_AMOUNT)]
    status = "ok" if flagged.empty else "review"
    msg = (
        "Every category's cost and depreciation rollforward ties to the TB."
        if status == "ok"
        else f"{len(flagged)} categor{'y' if len(flagged) == 1 else 'ies'} don't tie to the TB - check for movements outside the nominal activity supplied (e.g. a revaluation or manual journal)."
    )
    return FixedAssetResult("Fixed asset register (category summary)", status, msg, detail)


_METHOD_ALIASES = {
    "sl": STRAIGHT_LINE, "straightline": STRAIGHT_LINE, "straight-line": STRAIGHT_LINE, "straight line": STRAIGHT_LINE,
    "rb": REDUCING_BALANCE, "reducingbalance": REDUCING_BALANCE, "reducing balance": REDUCING_BALANCE,
    "wdv": REDUCING_BALANCE, "written down value": REDUCING_BALANCE, "diminishing balance": REDUCING_BALANCE,
}


def _normalise_method(raw: str) -> str:
    key = str(raw or "").strip().lower()
    return _METHOD_ALIASES.get(key, STRAIGHT_LINE if not key else key)


def asset_level_rollforward(
    prior_register: pd.DataFrame,
    nominal_activity: pd.DataFrame | None,
    tb_current: pd.DataFrame | None,
    period_days: int = 365,
) -> AssetRegisterResult:
    """Rolls a prior-year asset register forward: depreciates each asset by
    its own method/rate, flags additions found in nominal activity that
    aren't yet in the register, flags likely disposals, and reconciles
    totals to the TB."""
    if prior_register is None or prior_register.empty:
        return AssetRegisterResult("n/a", "No prior year fixed asset register uploaded.")

    reg = prior_register.copy()
    reg["depreciation_method"] = reg["depreciation_method"].map(_normalise_method)
    period_fraction = min(1.0, max(0.0, period_days / 365.0))

    def charge_for(row) -> float:
        rate = float(row["depreciation_rate"]) / 100.0 if row["depreciation_rate"] else 0.0
        base = row["cost"] if row["depreciation_method"] == STRAIGHT_LINE else (row["cost"] - row["accumulated_depreciation_b_fwd"])
        return round(max(0.0, base) * rate * period_fraction, 2)

    reg["depreciation_charge"] = reg.apply(charge_for, axis=1)
    reg["accumulated_depreciation_c_fwd"] = (reg["accumulated_depreciation_b_fwd"] + reg["depreciation_charge"]).clip(upper=reg["cost"])
    reg["nbv_b_fwd"] = reg["cost"] - reg["accumulated_depreciation_b_fwd"]
    reg["nbv_c_fwd"] = reg["cost"] - reg["accumulated_depreciation_c_fwd"]

    still_held = reg[~reg.get("disposed", pd.Series(False, index=reg.index)).astype(bool)]
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
                new_additions = additions_raw[["date", "account_code", "account_name", "reference", "description", "contact", "debit"]].rename(
                    columns={"debit": "Cost", "account_name": "Account", "date": "Date", "reference": "Reference", "description": "Description", "contact": "Contact", "account_code": "Nominal Code"}
                )
                new_additions["Category (to complete)"] = ""
                new_additions["Depreciation rate % (to complete)"] = ""

            disposals_raw = movement[movement["credit"] > 0]
            if not disposals_raw.empty:
                possible_disposals = disposals_raw[["date", "account_code", "account_name", "reference", "description", "contact", "credit"]].rename(
                    columns={"credit": "Cost removed", "account_name": "Account", "date": "Date", "reference": "Reference", "description": "Description", "contact": "Contact", "account_code": "Nominal Code"}
                )
                possible_disposals["Matched to asset ID (to complete)"] = ""

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
        if abs(variance) > MATERIALITY_AMOUNT:
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

    return AssetRegisterResult(status, message, summary, asset_schedule, new_additions, possible_disposals, closing_register)
