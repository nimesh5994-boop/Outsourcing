"""Minimal Resend-based email sending - the app only ever sends one kind
of email (a practice invite link), so this is deliberately thin: a single
function, plain urllib rather than pulling in a new HTTP-client dependency
for one outbound call.

Degrades silently when RESEND_API_KEY/RESEND_FROM_ADDRESS aren't
configured - logs a warning and returns False rather than raising, since
the admin invite screen always shows the raw invite link too (see
main.py's /admin/invites route), so a missing/failed email is an
inconvenience for the admin, not a broken feature."""
from __future__ import annotations

import json
import logging
import os
import urllib.error
import urllib.request

logger = logging.getLogger(__name__)

RESEND_API_URL = "https://api.resend.com/emails"


def send_invite_email(to_email: str, invite_link: str) -> bool:
    api_key = os.getenv("RESEND_API_KEY")
    from_address = os.getenv("RESEND_FROM_ADDRESS")
    if not api_key or not from_address:
        logger.warning("RESEND_API_KEY/RESEND_FROM_ADDRESS not configured - not emailing invite to %s", to_email)
        return False

    payload = {
        "from": from_address,
        "to": [to_email],
        "subject": "You're invited to Working Paper Automation",
        "html": (
            "<p>You've been invited to set up a practice on Working Paper Automation.</p>"
            f'<p><a href="{invite_link}">Click here to create your practice</a></p>'
            f"<p>This link is single-use and tied to {to_email}.</p>"
        ),
    }
    request = urllib.request.Request(
        RESEND_API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=10) as response:
            return 200 <= response.status < 300
    except urllib.error.URLError:
        logger.exception("Failed to send invite email to %s", to_email)
        return False
