"""Admin authentication: pbkdf2 password hashing + stateless HMAC session tokens.

Credentials are seeded from :data:`DEFAULT_ADMIN_USERNAME` / :data:`DEFAULT_ADMIN_PASSWORD`
and then persisted (hashed) in the config folder by ``config_store``; the admin changes
them from the Admin page.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
from typing import Any, Dict, Optional, Tuple

DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "changeme"  # forced change on first successful login
_PBKDF2_ROUNDS = 200_000
TOKEN_TTL_SECONDS = 8 * 3600


def _b64e(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def _b64d(txt: str) -> bytes:
    return base64.urlsafe_b64decode(txt + "=" * (-len(txt) % 4))


def hash_password(password: str, salt_hex: Optional[str] = None) -> Tuple[str, str]:
    salt = bytes.fromhex(salt_hex) if salt_hex else os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, _PBKDF2_ROUNDS)
    return salt.hex(), dk.hex()


def verify_password(password: str, salt_hex: str, hash_hex: str) -> bool:
    _, calc = hash_password(password, salt_hex)
    return hmac.compare_digest(calc, hash_hex)


def default_admin() -> Dict[str, Any]:
    salt, h = hash_password(DEFAULT_ADMIN_PASSWORD)
    return {"username": DEFAULT_ADMIN_USERNAME, "salt": salt, "hash": h,
            "must_change": True}


def make_token(username: str, secret: str, now: Optional[float] = None) -> str:
    payload = {"u": username, "exp": int((now or time.time()) + TOKEN_TTL_SECONDS)}
    body = _b64e(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
    sig = _b64e(hmac.new(secret.encode("utf-8"), body.encode("ascii"), hashlib.sha256).digest())
    return f"{body}.{sig}"


def verify_token(token: str, secret: str, now: Optional[float] = None) -> Optional[str]:
    try:
        body, sig = token.split(".", 1)
        expected = _b64e(hmac.new(secret.encode("utf-8"), body.encode("ascii"),
                                  hashlib.sha256).digest())
        if not hmac.compare_digest(sig, expected):
            return None
        payload = json.loads(_b64d(body))
        if payload.get("exp", 0) < (now or time.time()):
            return None
        return payload.get("u")
    except Exception:  # noqa: BLE001
        return None


def new_secret() -> str:
    return os.urandom(32).hex()
