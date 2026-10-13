"""Xero API module tests - OAuth, client retry/refresh behaviour, and the
converters from API JSON to the canonical DataFrames. Uses fake HTTP
transports throughout (no network, no real credentials), so the API JSON
shapes here are Xero's documented structure as understood when this was
written - the converters key off header names and the account-id attribute
rather than fixed positions to tolerate variation, and the first pull
against a real organisation should be checked against the raw report the
app stores with each pull (see main.py's xero pull route)."""
import json
import urllib.parse
from pathlib import Path

import pandas as pd
import pytest

from app import parsers, xero_api, xero_reports

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"


@pytest.fixture(autouse=True)
def _env(monkeypatch):
    monkeypatch.setenv("SECRET_KEY", "test-secret-key")
    monkeypatch.setenv("XERO_CLIENT_ID", "cid")
    monkeypatch.setenv("XERO_CLIENT_SECRET", "csecret")
    monkeypatch.delenv("XERO_SCOPES", raising=False)


def _tb_report(rows_by_section: dict[str, list[tuple]], header=("Account", "Debit", "Credit", "YTD Debit", "YTD Credit")) -> dict:
    """rows: (label, account_id, ytd_debit, ytd_credit); the monthly Debit/
    Credit columns are filled with deliberately different junk so a parser
    that reads the wrong pair fails."""
    sections = []
    for title, rows in rows_by_section.items():
        sections.append({"RowType": "Section", "Title": title, "Rows": [
            {"RowType": "Row", "Cells": [
                {"Value": label, "Attributes": [{"Value": acc_id, "Id": "account"}] if acc_id else []},
                {"Value": "999.00"}, {"Value": "888.00"},
                {"Value": f"{ytd_dr:.2f}"}, {"Value": f"{ytd_cr:.2f}"},
            ]} for label, acc_id, ytd_dr, ytd_cr in rows
        ]})
    return {"Reports": [{"ReportID": "TrialBalance", "Rows": [
        {"RowType": "Header", "Cells": [{"Value": h} for h in header]}, *sections,
        {"RowType": "Section", "Title": "", "Rows": [{"RowType": "SummaryRow", "Cells": [{"Value": "Total"}, {"Value": "1"}, {"Value": "1"}, {"Value": "1"}, {"Value": "1"}]}]},
    ]}]}


def test_trial_balance_frame_matches_the_export_parser_on_the_same_data():
    expected, _ = xero_reports.parse_trial_balance(parsers.FileDataSource(SAMPLE_DIR / "trial_balance_current_xero.xlsx"))
    accounts = {f"id-{i}": {"code": r.account_code, "name": r.account_name, "type": r.account_type} for i, r in enumerate(expected.itertuples())}
    report = _tb_report({"All": [(f"{r.account_name} ({r.account_code})", f"id-{i}", r.debit, r.credit) for i, r in enumerate(expected.itertuples())]})

    got = xero_api.trial_balance_frame(report, accounts)

    pd.testing.assert_frame_equal(got.reset_index(drop=True), expected.reset_index(drop=True), check_dtype=False)


def test_trial_balance_uses_ytd_columns_not_the_monthly_pair():
    report = _tb_report({"Revenue": [("Sales (200)", "a1", 0.0, 1234.0)]})
    df = xero_api.trial_balance_frame(report, {"a1": {"code": "200", "name": "Sales", "type": "Revenue"}})
    assert df.iloc[0]["credit"] == 1234.0 and df.iloc[0]["debit"] == 0.0 and df.iloc[0]["balance"] == -1234.0


def test_trial_balance_falls_back_to_plain_debit_credit_when_no_ytd_columns():
    report = _tb_report({"x": [("Sales (200)", "a1", 5.0, 6.0)]}, header=("Account", "Debit", "Credit"))
    report["Reports"][0]["Rows"][1]["Rows"][0]["Cells"] = [{"Value": "Sales (200)", "Attributes": [{"Value": "a1", "Id": "account"}]}, {"Value": "10"}, {"Value": "0"}]
    df = xero_api.trial_balance_frame(report, {"a1": {"code": "200", "name": "Sales", "type": "Revenue"}})
    assert df.iloc[0]["debit"] == 10.0


def test_name_only_account_is_keyed_by_its_name_like_the_export_parser_does():
    report = _tb_report({"Bank": [("Starling Business Account", "a2", 100.0, 0.0)]})
    df = xero_api.trial_balance_frame(report, {"a2": {"code": "", "name": "Starling Business Account", "type": "Bank"}})
    assert df.iloc[0]["account_code"] == "Starling Business Account"


def test_unknown_account_id_falls_back_to_the_code_in_the_label():
    report = _tb_report({"x": [("Office costs (412)", None, 50.0, 0.0)]})
    df = xero_api.trial_balance_frame(report, {})
    assert (df.iloc[0]["account_code"], df.iloc[0]["account_name"]) == ("412", "Office costs")


def test_account_types_translate_to_the_labels_the_rest_of_the_app_keys_off():
    # these exact strings are matched elsewhere (PL_ACCOUNT_TYPES, stock/NCL/bank detection)
    assert xero_api.account_type_label("TERMLIAB") == "Non-current Liability"
    assert xero_api.account_type_label("INVENTORY") == "Inventory"
    assert xero_api.account_type_label("DIRECTCOSTS") == "Direct Costs"
    assert xero_api.account_type_label("OVERHEADS") == "Overhead"
    assert xero_api.account_type_label("BANK") == "Bank"
    assert xero_api.account_type_label("SOMETHINGNEW") == "Somethingnew"
    # every label Xero's real exports use for P&L must be recognised as P&L by the existing derive step
    for token in ("REVENUE", "SALES", "DIRECTCOSTS", "OVERHEADS", "EXPENSE", "OTHERINCOME"):
        assert xero_api.account_type_label(token).lower() in xero_reports.PL_ACCOUNT_TYPES


# ---------- OAuth + client ----------

class FakeHttp:
    def __init__(self, responses):
        self.responses, self.calls = list(responses), []

    def __call__(self, method, url, headers, body=None):
        self.calls.append((method, url, headers, body))
        resp = self.responses.pop(0)
        status, payload, extra = resp[0], resp[1], (resp[2] if len(resp) > 2 else {})
        return status, {k.lower(): v for k, v in extra.items()}, json.dumps(payload).encode()


def test_authorize_url_requests_default_scopes_and_carries_state():
    url = xero_api.authorize_url("https://wp.example/xero/callback", "STATE123")
    q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
    assert q["state"] == ["STATE123"] and q["client_id"] == ["cid"] and q["response_type"] == ["code"]
    assert "offline_access" in q["scope"][0] and "accounting.reports.trialbalance.read" in q["scope"][0]
    assert "journals" not in q["scope"][0]


def test_scopes_are_overridable_to_request_journals(monkeypatch):
    monkeypatch.setenv("XERO_SCOPES", "offline_access accounting.settings accounting.reports.trialbalance.read accounting.journals.read")
    assert "accounting.journals.read" in xero_api.requested_scopes()


def test_exchange_code_stores_tokens_encrypted_never_plaintext():
    http = FakeHttp([(200, {"access_token": "ACCESS", "refresh_token": "REFRESH", "expires_in": 1800, "scope": "offline_access accounting.settings"})])
    fields = xero_api.exchange_code("code", "https://x/cb", http=http, now=lambda: 1000)
    assert "ACCESS" not in json.dumps(fields) and "REFRESH" not in json.dumps(fields)
    assert xero_api.decrypt_token(fields["refresh_token_enc"]) == "REFRESH"
    assert fields["expires_at"] == 2800
    method, url, headers, body = http.calls[0]
    assert headers["Authorization"].startswith("Basic ") and b"grant_type=authorization_code" in body


def test_failed_token_exchange_raises_auth_error():
    with pytest.raises(xero_api.XeroAuthError):
        xero_api.exchange_code("bad", "https://x/cb", http=FakeHttp([(400, {"error": "invalid_grant"})]))


def _connection(expires_at=10_000, scopes="accounting.settings"):
    return {"tenant_id": "T1", "scopes": scopes, "expires_at": expires_at,
            "access_token_enc": xero_api.encrypt_token("OLD_ACCESS"), "refresh_token_enc": xero_api.encrypt_token("OLD_REFRESH")}


def test_client_refreshes_an_expired_token_and_saves_the_rotated_credentials():
    saved = []
    http = FakeHttp([
        (200, {"access_token": "NEW_ACCESS", "refresh_token": "NEW_REFRESH", "expires_in": 1800, "scope": "accounting.settings"}),
        (200, {"Organisations": []}),
    ])
    conn = _connection(expires_at=0)
    client = xero_api.XeroClient(conn, saved.append, http=http, now=lambda: 5000)
    client.get("Organisation")
    assert len(saved) == 1  # single-use refresh token: the rotated one must be persisted
    assert xero_api.decrypt_token(conn["refresh_token_enc"]) == "NEW_REFRESH"
    assert http.calls[1][2]["Authorization"] == "Bearer NEW_ACCESS" and http.calls[1][2]["xero-tenant-id"] == "T1"


def test_client_does_not_refresh_a_fresh_token():
    http = FakeHttp([(200, {"ok": 1})])
    xero_api.XeroClient(_connection(expires_at=99_999), lambda c: None, http=http, now=lambda: 5000).get("Accounts")
    assert len(http.calls) == 1 and http.calls[0][2]["Authorization"] == "Bearer OLD_ACCESS"


def test_client_retries_after_rate_limit_using_retry_after():
    sleeps = []
    http = FakeHttp([(429, {}, {"Retry-After": "3"}), (200, {"ok": 1})])
    out = xero_api.XeroClient(_connection(99_999), lambda c: None, http=http, sleep=sleeps.append, now=lambda: 0).get("Accounts")
    assert out == {"ok": 1} and sleeps == [3.0]


def test_client_retries_a_401_once_after_refreshing_then_gives_up_as_auth_error():
    http = FakeHttp([
        (401, {}), (200, {"access_token": "A2", "refresh_token": "R2", "expires_in": 1800}), (401, {}),
    ])
    with pytest.raises(xero_api.XeroAuthError):
        xero_api.XeroClient(_connection(99_999), lambda c: None, http=http, now=lambda: 0).get("Accounts")


def test_client_403_is_forbidden_not_auth_error():
    with pytest.raises(xero_api.XeroForbidden):
        xero_api.XeroClient(_connection(99_999), lambda c: None, http=FakeHttp([(403, {"Detail": "no"})]), now=lambda: 0).get("Journals")


# ---------- journals ----------

def _journal(date_ms, source, lines, ref="REF"):
    return {"JournalID": "j", "JournalDate": f"/Date({date_ms}+0000)/", "SourceType": source, "SourceID": "inv-1", "Reference": ref, "JournalLines": lines}


def _line(code, name, net, tax=0.0, desc="d"):
    return {"AccountCode": code, "AccountName": name, "NetAmount": net, "TaxAmount": tax, "Description": desc}


MS_2025_06_15 = 1749945600000  # 2025-06-15 UTC


def test_journals_frame_signs_dates_contra_and_source_labels():
    j = _journal(MS_2025_06_15, "ACCREC", [_line("200", "Sales", -100.0, -20.0), _line("610", "Trade Debtors", 120.0), _line("820", "VAT", -20.0)])
    df = xero_api.journals_frame([j], "2025-04-01", "2026-03-31", contacts_by_source_id={"inv-1": "Acme Ltd"})

    sales = df[df.account_code == "200"].iloc[0]
    assert (sales.debit, sales.credit, sales.net) == (0.0, 100.0, -100.0)
    assert sales.date == pd.Timestamp("2025-06-15")
    assert sales.source_type == "Receivable Invoice" and sales.contact == "Acme Ltd"
    assert sales.contra_name == "Trade Debtors" and bool(sales.contra_needs_review) is True  # debtors + VAT = more than one other
    debtors = df[df.account_code == "610"].iloc[0]
    assert debtors.debit == 120.0 and debtors.contra_name == "Sales"


def test_journals_outside_the_period_are_excluded():
    inside = _journal(MS_2025_06_15, "MANJOURNAL", [_line("200", "Sales", -5.0), _line("610", "Debtors", 5.0)])
    before = _journal(1700000000000, "MANJOURNAL", [_line("200", "Sales", -9.0), _line("610", "Debtors", 9.0)])
    df = xero_api.journals_frame([inside, before], "2025-04-01", "2026-03-31")
    assert set(df.debit) | set(df.credit) == {0.0, 5.0}
    assert (df.source_type == "Manual Journal").all()


def test_journals_frame_has_every_column_the_account_transactions_parser_produces():
    expected = xero_reports.parse_account_transactions(parsers.FileDataSource(SAMPLE_DIR / "account_transactions_current_xero.xlsx"))
    got = xero_api.journals_frame([], "2025-01-01", "2025-12-31")
    assert list(got.columns) == list(expected.columns)


def test_manual_journal_source_label_is_recognised_by_the_anomaly_checks():
    # anomaly_detection / recon match on "journal" in the source label
    assert "journal" in xero_api.SOURCE_TYPE_LABELS["MANJOURNAL"].lower()


# ---------- period check + frame persistence ----------

def test_financial_year_warning_only_when_the_period_cant_be_a_ytd_tb():
    org = {"Organisations": [{"FinancialYearEndDay": 31, "FinancialYearEndMonth": 3}]}
    assert xero_api.financial_year_warning(org, "2025-04-01", "2026-03-31") is None
    assert xero_api.financial_year_warning(org, "2025-07-01", "2026-03-31") is not None  # part-year
    assert xero_api.financial_year_warning(org, "2025-01-01", "2025-12-31") is not None  # wrong year end
    assert xero_api.financial_year_warning({"Organisations": [{}]}, "2025-01-01", "2025-12-31") is None


def test_frame_bytes_roundtrip_keeps_leading_zero_codes_and_dates():
    df = pd.DataFrame({"date": [pd.Timestamp("2025-06-15")], "account_code": ["0100"], "debit": [1.5]})
    back = xero_api.frame_from_bytes(xero_api.frame_to_bytes(df))
    assert back.iloc[0]["account_code"] == "0100" and back.iloc[0]["date"] == pd.Timestamp("2025-06-15")
