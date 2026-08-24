"""Opt-in LLM assistance for flagged reconciliation checks.

Deliberately narrow and clearly separated from the deterministic engine
(recon.py, control_accounts.py, etc. make no API calls and never will -
this module is called from outside them, in main.py's generate(), and
only ever *adds* a note to an already-computed ReconResult; it never
changes a status or a number). The model sees exactly what a human
reviewer already sees - the flagged check's own detail tables and the
practice's own instruction note for that report type - and is asked for
a short, hedged suggestion of where to look, not a verdict. Every note
that reaches the workbook is prefixed so it's unmistakably AI-assisted,
not a system finding.

Off by default (see DEFAULT_TEMPLATE_CONFIG in storage.py); requires
ANTHROPIC_API_KEY. Any failure (missing key, API error, timeout) degrades
to no note rather than breaking generation - a working paper must still
generate correctly with zero AI involvement.
"""
import os

import pandas as pd

MODEL = "claude-haiku-4-5-20251001"
MAX_TOKENS = 300
NO_INSIGHT_MARKER = "NO_ADDITIONAL_INSIGHT"

SYSTEM_PROMPT = """You are assisting a UK accountant reviewing a working paper during an \
accounts preparation/review process. You will be shown one reconciliation check that has \
already been flagged for review by deterministic code, along with whatever supporting data \
is available.

Your job: suggest, in 2-4 short sentences, where the reviewer should look or what might \
explain the variance - using only the data given to you. Be specific about which line items \
or accounts look relevant if any stand out. Never invent figures that aren't in the data \
given to you. Never state a conclusion as fact - this is a hint for a human reviewer, who \
will verify everything themselves, not an audit finding.

If the data given to you offers no real additional insight beyond the check's own message \
(e.g. there's no supporting detail, or nothing in it stands out), respond with exactly this \
and nothing else: NO_ADDITIONAL_INSIGHT"""


def _df_to_text(df: pd.DataFrame, max_rows: int = 30) -> str:
    if df is None or df.empty:
        return "(none)"
    if len(df) > max_rows:
        return df.head(max_rows).to_string(index=False) + f"\n... ({len(df) - max_rows} more rows not shown)"
    return df.to_string(index=False)


def _client() -> "object | None":
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None
    import anthropic

    return anthropic.Anthropic(api_key=api_key)


def explain_flagged_result(
    check_name: str, status: str, message: str,
    detail: pd.DataFrame, extra_detail: pd.DataFrame, extra_detail_label: str,
    instruction_note: str, client_name: str, period_label: str,
) -> str:
    """Returns a short AI-assisted note, or "" if the feature can't run
    (no API key) or the model had nothing useful to add. Never raises -
    generation must succeed with or without this."""
    client = _client()
    if client is None:
        return "(AI-assisted notes are enabled but ANTHROPIC_API_KEY isn't configured - see README)"

    prompt_parts = [
        f"Client: {client_name}\nPeriod: {period_label}",
        f"Check: {check_name}\nStatus: {status.upper()}\nMessage: {message}",
        f"Supporting detail:\n{_df_to_text(detail)}",
    ]
    if extra_detail is not None and not extra_detail.empty:
        prompt_parts.append(f"{extra_detail_label}:\n{_df_to_text(extra_detail)}")
    if instruction_note:
        prompt_parts.append(f"Practice's own note on how this client's data should be read: {instruction_note}")

    try:
        response = client.messages.create(
            model=MODEL, max_tokens=MAX_TOKENS, system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": "\n\n".join(prompt_parts)}],
        )
        text = "".join(block.text for block in response.content if block.type == "text").strip()
    except Exception as exc:
        return f"(AI-assisted note unavailable: {type(exc).__name__})"

    if not text or text == NO_INSIGHT_MARKER:
        return ""
    return text
