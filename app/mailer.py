"""Minimal Resend-based email sending - the app only ever sends two kinds
of transactional email (a practice invite link, a password reset link),
so this is deliberately thin: plain urllib rather than pulling in a new
HTTP-client dependency for two outbound calls.

Degrades silently when RESEND_API_KEY/RESEND_FROM_ADDRESS aren't
configured - logs a warning and returns False rather than raising, since
callers always show a fallback (the admin invite screen shows the raw
invite link; forgot-password shows a generic "check your email" message
either way), so a missing/failed email is an inconvenience, not a broken
feature."""
from __future__ import annotations

import json
import logging
import os
import urllib.error
import urllib.request

logger = logging.getLogger(__name__)

RESEND_API_URL = "https://api.resend.com/emails"


def _send(to_email: str, subject: str, html: str) -> bool:
    api_key = os.getenv("RESEND_API_KEY")
    from_address = os.getenv("RESEND_FROM_ADDRESS")
    if not api_key or not from_address:
        logger.warning("RESEND_API_KEY/RESEND_FROM_ADDRESS not configured - not emailing %s", to_email)
        return False

    payload = {"from": from_address, "to": [to_email], "subject": subject, "html": html}
    request = urllib.request.Request(
        RESEND_API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            # Resend sits behind Cloudflare, which blocks the default
            # "Python-urllib/x.y" user agent as a bot signature (seen as
            # Cloudflare error code 1010) - any normal-looking UA passes.
            "User-Agent": "outsourcing-working-papers/1.0",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            return 200 <= response.status < 300
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        logger.error(
            "Resend rejected email to %s: HTTP %s - %s (key_len=%s key_prefix=%r from=%r)",
            to_email, exc.code, body, len(api_key), api_key[:6], from_address,
        )
        return False
    except urllib.error.URLError:
        logger.exception("Failed to send email to %s", to_email)
        return False


def send_invite_email(to_email: str, invite_link: str) -> bool:
    return _send(
        to_email,
        "You're invited to Working Paper Automation",
        (
            "<p>You've been invited to set up a practice on Working Paper Automation.</p>"
            f'<p><a href="{invite_link}">Click here to create your practice</a></p>'
            f"<p>This link is single-use and tied to {to_email}.</p>"
        ),
    )


def send_new_user_email(to_email: str, practice_name: str, role: str, setup_link: str) -> bool:
    return _send(
        to_email,
        f"You've been added to {practice_name} on Working Paper Automation",
        (
            f"<p>You've been added to <strong>{practice_name}</strong> on Working Paper Automation as a {role}.</p>"
            f'<p><a href="{setup_link}">Click here to set your password and log in</a></p>'
            "<p>This link is single-use and expires in 7 days.</p>"
        ),
    )


def send_password_reset_email(to_email: str, reset_link: str) -> bool:
    return _send(
        to_email,
        "Reset your Working Paper Automation password",
        (
            "<p>Someone requested a password reset for this account.</p>"
            f'<p><a href="{reset_link}">Click here to choose a new password</a></p>'
            "<p>This link is single-use and expires in 1 hour. If you didn't request this, "
            "you can ignore this email.</p>"
        ),
    )
