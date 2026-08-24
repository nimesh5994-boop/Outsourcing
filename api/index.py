"""Vercel entrypoint: exposes the FastAPI app as an ASGI callable.

Vercel's Python runtime looks for a module-level ASGI/WSGI app in
api/*.py and routes requests to it directly - no separate server process,
no local filesystem persistence between invocations (which is exactly why
app/storage.py is Postgres-backed rather than JSON-file-backed).
"""
from app.main import app  # noqa: F401
