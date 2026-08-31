"""Unit tests for recon.bank_reconciliation's account-name matching - see
app/recon.py. No dedicated test file existed for this check before (it
was only exercised end-to-end via test_pipeline.py/test_configurable_
materiality.py with exact-matching fixture names), so the fuzzy-match
fallback's real behaviour had never actually been exercised."""
import pandas as pd

from app import recon


def _stmt(account_name, closing_balance, statement_date="2025-12-31"):
    return pd.DataFrame([{"account_name": account_name, "statement_date": statement_date, "closing_balance": closing_balance}])


def _tb_row(code, name, balance):
    return {"account_code": code, "account_name": name, "account_type": "Bank", "balance": balance}


def test_exact_name_match_ties_out():
    bank_statement = _stmt("BANK CURRENT ACCOUNT", 5000.0)
    tb = pd.DataFrame([_tb_row("1200", "BANK CURRENT ACCOUNT", 5000.0)])
    result = recon.bank_reconciliation(bank_statement, tb, materiality=500)
    assert result.status == "ok"
    assert result.detail.iloc[0]["TB book balance"] == 5000.0


def test_fuzzy_match_ties_out_when_unique():
    # Regression test for a real bug: the substring-fallback match used to
    # be gated behind an unused `bank_account_names` parameter that no
    # caller ever actually passed - its own content played no part in the
    # match, it was purely an always-off toggle - so a bank statement
    # account name that didn't match the TB byte-for-byte (the normal
    # case for a real export, not the exception: "NATWEST CURRENT
    # ACCOUNT" on the statement vs "BANK - NATWEST CURRENT ACCOUNT" in
    # the TB) always came back as a bogus "unreconciled" finding for the
    # account's entire balance, even though the two obviously refer to
    # the same real account.
    bank_statement = _stmt("NATWEST CURRENT ACCOUNT", 50000.0)
    tb = pd.DataFrame([_tb_row("1200", "BANK - NATWEST CURRENT ACCOUNT", 50000.0)])
    result = recon.bank_reconciliation(bank_statement, tb, materiality=500)
    assert result.status == "ok"
    assert result.detail.iloc[0]["TB book balance"] == 50000.0
    assert result.detail.iloc[0]["Unreconciled variance"] == 0.0


def test_fuzzy_match_refuses_to_guess_when_ambiguous():
    # A bare "Bank" on the statement is a substring of BOTH "Bank Current
    # Account" and "Bank Deposit Account" in the TB - silently summing
    # both would produce a meaningless combined balance and could mask a
    # real variance on either individual account, so the fuzzy fallback
    # must only apply when it uniquely identifies exactly one TB account.
    bank_statement = _stmt("BANK", 5000.0)
    tb = pd.DataFrame([
        _tb_row("1200", "BANK CURRENT ACCOUNT", 5000.0),
        _tb_row("1210", "BANK DEPOSIT ACCOUNT", 20000.0),
    ])
    result = recon.bank_reconciliation(bank_statement, tb, materiality=500)
    assert result.detail.iloc[0]["TB book balance"] == 0.0
    assert result.status == "review"


def test_no_matching_tb_account_reports_full_balance_as_variance():
    bank_statement = _stmt("UNKNOWN BUILDING SOCIETY ACCOUNT", 1000.0)
    tb = pd.DataFrame([_tb_row("1200", "BANK CURRENT ACCOUNT", 5000.0)])
    result = recon.bank_reconciliation(bank_statement, tb, materiality=500)
    assert result.detail.iloc[0]["TB book balance"] == 0.0
    assert result.detail.iloc[0]["Unreconciled variance"] == 1000.0
    assert result.status == "review"


def test_na_without_bank_statement():
    assert recon.bank_reconciliation(None, None).status == "n/a"
    assert recon.bank_reconciliation(pd.DataFrame(), pd.DataFrame()).status == "n/a"
