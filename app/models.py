"""Canonical schema definitions.

Every source file (whatever platform it came from) is normalised into one of
these column layouts before any reconciliation logic runs. Reconciliation
and the Excel builder only ever deal with canonical columns, so a new source
platform is just a new mapping into this shape, not a new code path.
"""

REPORT_TYPES = [
    "trial_balance",
    "nominal_activity",
    "aged_debtors",
    "aged_creditors",
    "vat_return",
    "bank_statement",
    "profit_and_loss",
    "balance_sheet",
    "fixed_asset_register",
]

REPORT_LABELS = {
    "trial_balance": "Trial Balance",
    "nominal_activity": "Nominal Activity / General Ledger Detail",
    "aged_debtors": "Aged Debtors",
    "aged_creditors": "Aged Creditors",
    "vat_return": "VAT Return",
    "bank_statement": "Bank Closing Statement",
    "profit_and_loss": "Profit & Loss",
    "balance_sheet": "Balance Sheet",
    "fixed_asset_register": "Fixed Asset Register (prior year closing)",
}

# canonical column -> friendly label, used to drive the mapping UI
REPORT_SCHEMAS = {
    "trial_balance": {
        "account_code": "Nominal/Account Code",
        "account_name": "Account Name",
        "account_type": "Account Type (Sales/Direct Costs/Overhead/Bank/Current Asset/Current Liability/Equity/etc.)",
        "debit": "Debit",
        "credit": "Credit",
        "comparative_balance": "Comparative Year Balance (if embedded in the same export, net debit-credit convention)",
    },
    "nominal_activity": {
        "date": "Date",
        "account_code": "Nominal/Account Code",
        "account_name": "Account Name",
        "reference": "Reference",
        "description": "Description / Narrative",
        "contact": "Contact / Payee",
        "source_type": "Source (Journal/Invoice/Bill/Bank/etc.)",
        "debit": "Debit",
        "credit": "Credit",
    },
    "aged_debtors": {
        "customer": "Customer",
        "current": "Current",
        "bucket_1": "Overdue bucket 1 (e.g. <1 Month / 1-30 days)",
        "bucket_2": "Overdue bucket 2 (e.g. 1 Month / 31-60 days)",
        "bucket_3": "Overdue bucket 3 (e.g. 2 Months / 61-90 days)",
        "bucket_4": "Overdue bucket 4 (e.g. 3 Months)",
        "older": "Older / 90+ days",
        "total": "Total",
    },
    "aged_creditors": {
        "supplier": "Supplier",
        "current": "Current",
        "bucket_1": "Overdue bucket 1 (e.g. <1 Month / 1-30 days)",
        "bucket_2": "Overdue bucket 2 (e.g. 1 Month / 31-60 days)",
        "bucket_3": "Overdue bucket 3 (e.g. 2 Months / 61-90 days)",
        "bucket_4": "Overdue bucket 4 (e.g. 3 Months)",
        "older": "Older / 90+ days",
        "total": "Total",
    },
    "vat_return": {
        "box1": "Box 1 - VAT due on sales",
        "box2": "Box 2 - VAT due on acquisitions",
        "box3": "Box 3 - Total VAT due",
        "box4": "Box 4 - VAT reclaimed",
        "box5": "Box 5 - Net VAT due",
        "box6": "Box 6 - Total sales ex VAT",
        "box7": "Box 7 - Total purchases ex VAT",
        "box8": "Box 8 - Supplies to EU",
        "box9": "Box 9 - Acquisitions from EU",
    },
    "bank_statement": {
        "account_name": "Bank Account Name",
        "statement_date": "Statement Date",
        "closing_balance": "Closing Balance (per bank statement)",
    },
    "profit_and_loss": {
        "account_code": "Nominal/Account Code",
        "account_name": "Account Name",
        "category": "P&L Category (Turnover/COS/Overheads/etc.)",
        "amount": "Amount",
    },
    "balance_sheet": {
        "account_code": "Nominal/Account Code",
        "account_name": "Account Name",
        "category": "B/S Category (Fixed Assets/Current Assets/Current Liabilities/etc.)",
        "amount": "Amount",
    },
    "fixed_asset_register": {
        "asset_id": "Asset ID / Reference",
        "description": "Description",
        "category": "Category (Motor Vehicles/Fixtures & Fittings/Computer Equipment/etc.)",
        "date_acquired": "Date Acquired",
        "cost": "Cost",
        "depreciation_method": "Depreciation Method (Straight Line/Reducing Balance)",
        "depreciation_rate": "Depreciation Rate (% per annum)",
        "accumulated_depreciation_b_fwd": "Accumulated Depreciation Brought Forward",
        "disposed": "Disposed? (Yes/No)",
    },
}

# Fields required for a period upload to be usable in reconciliation.
REQUIRED_FIELDS = {
    "trial_balance": ["account_code", "account_name"],
    "nominal_activity": ["date", "account_code", "description"],
    "aged_debtors": ["customer", "total"],
    "aged_creditors": ["supplier", "total"],
    "vat_return": ["box5", "box6", "box7"],
    "bank_statement": ["account_name", "closing_balance"],
    "profit_and_loss": ["account_name", "amount"],
    "balance_sheet": ["account_name", "amount"],
    "fixed_asset_register": ["description", "cost"],
}

PLATFORMS = ["xero", "qbo", "sage", "other"]

PERIODS = ["current", "comparative"]
