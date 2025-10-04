# token_helper.py
# Freeze-safe Xero OAuth2 token helper:
# - Always reads/writes xero_tokens.json next to this file (or next to the EXE)
# - Loads .env from the same folder
# - Auto-refreshes tokens with a small buffer
# - Atomic writes to avoid corrupt files

import os
import sys
import time
import json
from pathlib import Path

import requests
from dotenv import load_dotenv

# ---------- locate files next to script/EXE ----------
if getattr(sys, "frozen", False):  # running as PyInstaller EXE
    BASE_DIR = Path(sys.executable).parent
else:                              # running as .py
    BASE_DIR = Path(__file__).parent

TOKENS_PATH = BASE_DIR / "xero_tokens.json"
ENV_PATH    = BASE_DIR / ".env"

load_dotenv(ENV_PATH)

# Support both XERO_* and plain names just in case
CLIENT_ID     = os.getenv("XERO_CLIENT_ID")     or os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("XERO_CLIENT_SECRET") or os.getenv("CLIENT_SECRET")
TOKEN_URL     = "https://identity.xero.com/connect/token"

# ---------- low-level file helpers ----------
def _load_tokens() -> dict:
    try:
        with open(TOKENS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError as e:
        raise FileNotFoundError(
            f"xero_tokens.json not found at: {TOKENS_PATH}\n"
            f"Put your existing token file here, or re-run your code-exchange step to create it."
        ) from e

def _save_tokens(raw: dict) -> dict:
    """
    Persist tokens atomically. Xero usually returns a new refresh_token on each refresh.
    Returns the normalized dict we stored.
    """
    expires_in  = int(raw.get("expires_in", 1800))
    expires_at  = int(time.time()) + expires_in - 60  # small buffer
    normalized = {
        "access_token":  raw["access_token"],
        "refresh_token": raw.get("refresh_token", ""),  # sometimes missing; we’ll patch below
        "expires_at":    expires_at,
    }
    tmp = TOKENS_PATH.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(normalized, f, indent=2)
    os.replace(tmp, TOKENS_PATH)
    return normalized

# ---------- OAuth refresh ----------
def _refresh_access_token(refresh_token: str) -> dict:
    if not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError("Missing CLIENT_ID/CLIENT_SECRET in .env (XERO_CLIENT_ID / XERO_CLIENT_SECRET).")

    data = {
        "grant_type":    "refresh_token",
        "refresh_token": refresh_token,
    }
    # Xero supports HTTP Basic auth with client id/secret
    resp = requests.post(TOKEN_URL, data=data, auth=(CLIENT_ID, CLIENT_SECRET), timeout=30)
    try:
        resp.raise_for_status()
    except requests.HTTPError as e:
        # Surface useful details for debugging
        detail = ""
        try:
            detail = f" | response={resp.status_code} {resp.text[:500]}"
        except Exception:
            pass
        raise RuntimeError(f"Token refresh failed{detail}") from e

    raw = resp.json()
    saved = _save_tokens(raw)

    # If Xero didn't return a new refresh_token, keep the old one
    if not saved.get("refresh_token"):
        saved["refresh_token"] = refresh_token
        _save_tokens(saved)

    return saved

# ---------- public API ----------
def ensure_access_token() -> str:
    """
    Returns a valid access token. Refreshes automatically if expired (or nearly).
    Reads/writes xero_tokens.json beside this file/EXE.
    """
    tokens = _load_tokens()
    # refresh if expired or inside the last ~30 seconds
    now = int(time.time())
    if int(tokens.get("expires_at", 0)) <= now + 30:
        tokens = _refresh_access_token(tokens["refresh_token"])
    return tokens["access_token"]
