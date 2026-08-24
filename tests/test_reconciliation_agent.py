"""Unit tests for the opt-in AI reconciliation-explanation agent.

Never calls the real Anthropic API - the client is mocked throughout, so
these tests run offline, free, and deterministically. What they verify is
the contract this feature must hold regardless of what the model says:
degrades to no note (or a diagnosable message) on any failure, never
raises, and the "no additional insight" marker is treated as an empty
note rather than clutter."""
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from app import reconciliation_agent as agent


def _detail_df():
    return pd.DataFrame([{"Check": "Box 5 vs VAT control", "VAT return": 800.0, "Nominal ledger": 700.0, "Variance": 100.0}])


def test_no_api_key_returns_diagnosable_message(monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    note = agent.explain_flagged_result(
        "VAT return cross-check", "review", "VAT return does not agree to the nominal ledger.",
        _detail_df(), pd.DataFrame(), "", "", "Acme Ltd", "Year ended 31 Dec 2025",
    )
    assert "ANTHROPIC_API_KEY" in note


def _mock_response(text: str):
    block = MagicMock()
    block.type = "text"
    block.text = text
    response = MagicMock()
    response.content = [block]
    return response


def test_successful_call_returns_model_text(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    mock_client = MagicMock()
    mock_client.messages.create.return_value = _mock_response(
        "The £100 variance likely relates to the March VAT journal posted to the control account."
    )
    with patch("anthropic.Anthropic", return_value=mock_client):
        note = agent.explain_flagged_result(
            "VAT return cross-check", "review", "VAT return does not agree to the nominal ledger.",
            _detail_df(), pd.DataFrame(), "", "", "Acme Ltd", "Year ended 31 Dec 2025",
        )
    assert "March VAT journal" in note
    mock_client.messages.create.assert_called_once()
    call_kwargs = mock_client.messages.create.call_args.kwargs
    assert call_kwargs["model"] == agent.MODEL
    assert "VAT return cross-check" in call_kwargs["messages"][0]["content"]


def test_no_additional_insight_marker_becomes_empty_string(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    mock_client = MagicMock()
    mock_client.messages.create.return_value = _mock_response(agent.NO_INSIGHT_MARKER)
    with patch("anthropic.Anthropic", return_value=mock_client):
        note = agent.explain_flagged_result(
            "VAT return cross-check", "review", "VAT return does not agree to the nominal ledger.",
            _detail_df(), pd.DataFrame(), "", "", "Acme Ltd", "Year ended 31 Dec 2025",
        )
    assert note == ""


def test_api_error_degrades_to_diagnosable_message_not_a_raise(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    mock_client = MagicMock()
    mock_client.messages.create.side_effect = RuntimeError("connection reset")
    with patch("anthropic.Anthropic", return_value=mock_client):
        note = agent.explain_flagged_result(
            "VAT return cross-check", "review", "VAT return does not agree to the nominal ledger.",
            _detail_df(), pd.DataFrame(), "", "", "Acme Ltd", "Year ended 31 Dec 2025",
        )
    assert "unavailable" in note
    assert "RuntimeError" in note


def test_instruction_note_and_extra_detail_reach_the_prompt(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    mock_client = MagicMock()
    mock_client.messages.create.return_value = _mock_response("Some note.")
    extra = pd.DataFrame([{"date": "2025-03-01", "account_name": "VAT Control", "description": "Qtr VAT journal"}])
    with patch("anthropic.Anthropic", return_value=mock_client):
        agent.explain_flagged_result(
            "VAT return cross-check", "review", "VAT return does not agree to the nominal ledger.",
            _detail_df(), extra, "Candidate reconciling items",
            "This client's VAT export has a Detail tab - use it for reconciliation.",
            "Acme Ltd", "Year ended 31 Dec 2025",
        )
    prompt = mock_client.messages.create.call_args.kwargs["messages"][0]["content"]
    assert "Candidate reconciling items" in prompt
    assert "Qtr VAT journal" in prompt
    assert "Detail tab" in prompt
