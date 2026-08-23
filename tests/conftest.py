"""Shared fixtures for the test suite - the canonical dataset used by both
test_pipeline.py (end-to-end pipeline checks) and test_formulas.py
(formula-linked schedule verification), sourced from sample_data/, which
mirrors the real Xero export structures for a fictional client."""
from pathlib import Path

import pytest

from app import mapping, parsers, xero_reports

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


@pytest.fixture
def canonical_data():
    tb_current, tb_comparative = xero_reports.parse_trial_balance(
        parsers.FileDataSource(SAMPLE_DIR / "trial_balance_current_xero.xlsx")
    )
    nominal_current = xero_reports.parse_account_transactions(
        parsers.FileDataSource(SAMPLE_DIR / "account_transactions_current_xero.xlsx")
    )
    aged_debtors = xero_reports.parse_aged_report(
        parsers.FileDataSource(SAMPLE_DIR / "aged_receivables_current_xero.xlsx"), "customer"
    )
    aged_creditors = xero_reports.parse_aged_report(
        parsers.FileDataSource(SAMPLE_DIR / "aged_payables_current_xero.xlsx"), "supplier"
    )
    pl_current, bs_current = xero_reports.derive_pl_bs_from_tb(tb_current)

    vat_source = parsers.FileDataSource(SAMPLE_DIR / "vat_return_current.csv")
    vat_mapping = mapping.suggest_mapping("vat_return", vat_source.raw_columns())
    vat_return = parsers.apply_mapping(vat_source, "vat_return", vat_mapping)

    bank_source = parsers.FileDataSource(SAMPLE_DIR / "bank_statement_current.csv")
    bank_mapping = mapping.suggest_mapping("bank_statement", bank_source.raw_columns())
    bank_statement = parsers.apply_mapping(bank_source, "bank_statement", bank_mapping)

    return {
        "tb_current": tb_current, "tb_comparative": tb_comparative,
        "nominal_current": nominal_current,
        "aged_debtors": aged_debtors, "aged_creditors": aged_creditors,
        "pl_current": pl_current, "bs_current": bs_current,
        "vat_return": vat_return, "bank_statement": bank_statement,
    }
