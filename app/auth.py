"""Email+password authentication and practice/role/client-scoped
authorization.

Sessions are a signed cookie (itsdangerous), not a server-side session
table: the cookie carries the user id and an issue time, verified against
SECRET_KEY on every request. That keeps auth stateless, which matters on
serverless - no session store to write on login or clean up on expiry,
just one user lookup per request.

Roles, each scoped to exactly one practice per user:
  partner  - full access: manage users, templates, all clients/jobs
  manager  - manage templates, work every client/job (no user management)
  preparer - only the clients explicitly granted via client_access; no
             template or user management

Password hashes are computed by callers (main.py) via hash_password() and
passed into storage.create_user() as plain strings - this module imports
storage (to load users / check client access), so storage can't import
back into this module without a cycle.
"""
import os

import bcrypt
from fastapi import Request
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

from app import storage

SESSION_COOKIE = "session"
SESSION_MAX_AGE = 14 * 24 * 3600  # 14 days

ROLES = ("partner", "manager", "preparer")


class Unauthenticated(Exception):
    """No valid session - the exception handler redirects to /login."""


class Forbidden(Exception):
    """Valid session, but not permitted for this resource/action."""

    def __init__(self, message: str = "You don't have permission to do that."):
        self.message = message


def _serializer() -> URLSafeTimedSerializer:
    return URLSafeTimedSerializer(os.environ["SECRET_KEY"], salt="wpa-session")


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(password: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(password.encode("utf-8"), password_hash.encode("utf-8"))
    except ValueError:
        return False


def create_session_value(user_id: str) -> str:
    return _serializer().dumps({"user_id": user_id})


def _read_session_value(cookie_value: str) -> str | None:
    try:
        data = _serializer().loads(cookie_value, max_age=SESSION_MAX_AGE)
    except (BadSignature, SignatureExpired):
        return None
    return data.get("user_id")


def set_session_cookie(response, request: Request, user_id: str) -> None:
    response.set_cookie(
        SESSION_COOKIE, create_session_value(user_id),
        max_age=SESSION_MAX_AGE, httponly=True,
        secure=request.url.scheme == "https", samesite="lax",
    )


def get_current_user(request: Request) -> dict | None:
    cookie = request.cookies.get(SESSION_COOKIE)
    if not cookie:
        return None
    user_id = _read_session_value(cookie)
    if not user_id:
        return None
    return storage.get_user(user_id)


def current_user_dep(request: Request) -> dict:
    """FastAPI dependency: every logged-in-only route takes
    `user: dict = Depends(auth.current_user_dep)`."""
    user = get_current_user(request)
    if not user:
        raise Unauthenticated()
    return user


def require_role(user: dict, *roles: str) -> None:
    if user["role"] not in roles:
        raise Forbidden(f"This action needs one of these roles: {', '.join(roles)}.")


def safe_next_path(next_path: str | None) -> str:
    """Only ever redirect to a relative in-app path - a `next` value taken
    from a query/form param could otherwise be used for an open redirect."""
    if next_path and next_path.startswith("/") and not next_path.startswith("//"):
        return next_path
    return "/practices"
