"""Column-header mapping: turns an arbitrary export (Xero, QBO, Sage, or
anything else) into the canonical schema in models.py.

Two layers:
1. Best-effort auto-suggestion from alias dictionaries below (common header
   text seen in Xero / QBO / Sage exports for each report type).
2. Manual override via the mapping UI. Once a user confirms a mapping for a
   client + report type + platform, it is saved as a reusable profile so the
   *next* period's file for that same client maps itself.
"""
import json
import re
from pathlib import Path

from app.models import REPORT_SCHEMAS

PROFILES_DIR = Path("data/clients")


def _normalise(header: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(header).lower())


# canonical_field -> set of normalised header aliases we've seen across
# Xero / QBO / Sage exports. Best-effort; extend as real exports are seen.
ALIASES = {
    "trial_balance": {
        "account_code": {"accountcode", "code", "nominalcode", "accountnumber", "acctnum", "account"},
        "account_name": {"accountname", "name", "account", "description", "accountdescription"},
        "account_type": {"accounttype", "type", "accountclass"},
        "debit": {"debit", "debitgbp", "debitamount", "dr"},
        "credit": {"credit", "creditgbp", "creditamount", "cr"},
    },
    "nominal_activity": {
        "date": {"date", "transactiondate", "txndate"},
        "account_code": {"accountcode", "nominalcode", "code", "account"},
        "account_name": {"accountname", "account"},
        "reference": {"reference", "refno", "docnum", "num", "invoiceno"},
        "description": {"description", "narrative", "memo", "memodescription", "details"},
        "contact": {"contact", "name", "payee", "customer", "supplier", "vendor"},
        "source_type": {"sourcetype", "transactiontype", "type"},
        "debit": {"debit", "debitamount", "dr"},
        "credit": {"credit", "creditamount", "cr"},
    },
    "aged_debtors": {
        "customer": {"customer", "customername", "name", "contact"},
        "current": {"current", "notdue"},
        "d1_30": {"130days", "1to30days", "030days", "current130"},
        "d31_60": {"3160days", "31to60days"},
        "d61_90": {"6190days", "61to90days"},
        "d90_plus": {"90daysplus", "over90days", "90plusdays", "older"},
        "total": {"total", "totaldue", "balance", "amountdue"},
    },
    "aged_creditors": {
        "supplier": {"supplier", "suppliername", "name", "vendor", "contact"},
        "current": {"current", "notdue"},
        "d1_30": {"130days", "1to30days", "030days"},
        "d31_60": {"3160days", "31to60days"},
        "d61_90": {"6190days", "61to90days"},
        "d90_plus": {"90daysplus", "over90days", "90plusdays", "older"},
        "total": {"total", "totaldue", "balance", "amountdue"},
    },
    "vat_return": {
        "box1": {"box1", "vatduesales"},
        "box2": {"box2", "vatdueacquisitions"},
        "box3": {"box3", "totalvatdue"},
        "box4": {"box4", "vatreclaimed", "vatreclaimedoncurrentpurchases"},
        "box5": {"box5", "netvatdue", "netvat"},
        "box6": {"box6", "totalsalesexvat", "totalvaluesalesexvat"},
        "box7": {"box7", "totalpurchasesexvat", "totalvaluepurchasesexvat"},
        "box8": {"box8", "suppliestoeu", "totalvaluegoodssuppliedexvat"},
        "box9": {"box9", "acquisitionsfromeu", "totalacquisitionsexvat"},
    },
    "bank_statement": {
        "account_name": {"accountname", "bankaccount", "account", "name"},
        "statement_date": {"statementdate", "date", "asatdate", "closingdate"},
        "closing_balance": {"closingbalance", "balance", "statementbalance", "endingbalance"},
    },
    "profit_and_loss": {
        "account_code": {"accountcode", "nominalcode", "code"},
        "account_name": {"accountname", "account", "description", "name"},
        "category": {"category", "group", "section", "accountgroup"},
        "amount": {"amount", "total", "value", "currentyear", "ytd", "balance"},
    },
    "balance_sheet": {
        "account_code": {"accountcode", "nominalcode", "code"},
        "account_name": {"accountname", "account", "description", "name"},
        "category": {"category", "group", "section", "accountgroup"},
        "amount": {"amount", "total", "value", "currentyear", "ytd", "balance"},
    },
    "fixed_asset_register": {
        "asset_id": {"assetid", "assetref", "assetreference", "id", "ref", "reference", "assetnumber"},
        "description": {"description", "asset", "assetdescription", "item", "name"},
        "category": {"category", "assetcategory", "assettype", "group", "class"},
        "date_acquired": {"dateacquired", "acquisitiondate", "purchasedate", "dateofpurchase", "date"},
        "cost": {"cost", "originalcost", "grosscost", "purchaseprice"},
        "depreciation_method": {"depreciationmethod", "method"},
        "depreciation_rate": {"depreciationrate", "rate", "raterate", "annualrate"},
        "accumulated_depreciation_b_fwd": {"accumulateddepreciation", "accumulateddepreciationbfwd", "depreciationbfwd", "accdepbfwd", "nbvbfwd", "wdvbfwd"},
        "disposed": {"disposed", "disposal", "soldwrittenoff"},
    },
}


def suggest_mapping(report_type: str, source_columns: list[str]) -> dict:
    """Return {source_column: canonical_field or None} best-effort guess."""
    aliases = ALIASES.get(report_type, {})
    suggestion = {}
    for col in source_columns:
        norm = _normalise(col)
        matched = None
        for canonical_field, alias_set in aliases.items():
            if norm in alias_set or norm == canonical_field.replace("_", ""):
                matched = canonical_field
                break
        suggestion[col] = matched
    return suggestion


def profile_path(client_id: str, report_type: str, platform: str) -> Path:
    return PROFILES_DIR / client_id / "mapping_profiles" / f"{report_type}__{platform}.json"


def load_profile(client_id: str, report_type: str, platform: str) -> dict | None:
    path = profile_path(client_id, report_type, platform)
    if path.exists():
        return json.loads(path.read_text())
    return None


def save_profile(client_id: str, report_type: str, platform: str, mapping: dict) -> None:
    path = profile_path(client_id, report_type, platform)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(mapping, indent=2))
