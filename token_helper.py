# token_helper.py — Cloud-friendly & local-friendly
import os
import sys
import time
import json
from pathlib import Path

import requests

# --- detect Streamlit + session state (Cloud mode) ---
try:
    import streamlit as st  # type: ignore
    _ss = st.session_state
except Exception:
    st = None
    _ss = {}

# --- freeze-safe base dir for local file mode ---
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

TOKENS_PATH = BASE_DIR / "xero_tokens.json"
TOKEN_URL = "https://identity.xero.com/connect/token"

def _in_cloud_mode() -> bool:
    # treat presence of Streamlit secrets as Cloud mode
    try:
        return st is not None and hasattr(st, "secrets") and ("XERO_CLIENT_ID" in st.secrets or "XERO_REFRESH_TOKEN" in st.secrets)
    except Exception:
        return False

def _get_client_id() -> str:
    return (st.secrets.get("XERO_CLIENT_ID") if _in_cloud_mode() else os.getenv("XERO_CLIENT_ID")) or ""

def _get_client_secret() -> str:
    return (st.secrets.get("XERO_CLIENT_SECRET") if _in_cloud_mode() else os.getenv("XERO_CLIENT_SECRET")) or ""

def _get_baseline_refresh_token() -> str:
    # Cloud: from secrets. Local: from env or tokens file.
    if _in_cloud_mode():
        return st.secrets.get("XERO_REFRESH_TOKEN", "")
    rt = os.getenv("XERO_REFRESH_TOKEN", "")
    if rt:
        return rt
    # last resort: local tokens file
    if TOKENS_PATH.exists():
        try:
            with open(TOKENS_PATH, "r", encoding="utf-8") as f:
                return (json.load(f) or {}).get("refresh_token", "")
        except Exception:
            pass
    return ""

# ---------- local file helpers ----------
def _load_tokens_file() -> dict:
    with open(TOKENS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def _save_tokens_file(tokens: dict):
    tmp = TOKENS_PATH.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(tokens, f, indent=2)
    os.replace(tmp, TOKENS_PATH)

# ---------- refresh core ----------
def _refresh_access_token(refresh_token: str) -> dict:
    cid = _get_client_id()
    csecret = _get_client_secret()
    if not cid or not csecret:
        raise RuntimeError("Missing XERO_CLIENT_ID / XERO_CLIENT_SECRET.")

    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }
    resp = requests.post(TOKEN_URL, data=data, auth=(cid, csecret), timeout=30)
    if not resp.ok:
        # helpful error detail
        try:
            detail = resp.json()
        except Exception:
            detail = resp.text[:400]
        raise RuntimeError(f"Token refresh failed ({resp.status_code}): {detail}")

    raw = resp.json()
    expires_in = int(raw.get("expires_in", 1800))
    tokens = {
        "access_token": raw["access_token"],
        "refresh_token": raw.get("refresh_token", refresh_token),  # keep old if not returned
        "expires_at": int(time.time()) + expires_in - 60,
    }
    return tokens

# ---------- public API ----------
def ensure_access_token() -> str:
    now = int(time.time())

    if _in_cloud_mode():
        # 1) use in-session cache if valid
        tok = _ss.get("xero_tokens")
        if tok and int(tok.get("expires_at", 0)) > now + 30:
            return tok["access_token"]

        # 2) refresh using the latest we have, else baseline from secrets
        base_rt = (tok or {}).get("refresh_token") or _get_baseline_refresh_token()
        if not base_rt:
            raise FileNotFoundError(
                "No refresh token available. Set XERO_REFRESH_TOKEN in Streamlit Secrets."
            )
        new_tok = _refresh_access_token(base_rt)
        _ss["xero_tokens"] = new_tok
        return new_tok["access_token"]

    # ---- Local file mode (uses xero_tokens.json) ----
    try:
        t = _load_tokens_file()
    except FileNotFoundError as e:
        raise FileNotFoundError(
            f"xero_tokens.json not found at: {TOKENS_PATH}\n"
            "Run your local code-exchange step once to create it."
        ) from e

    if int(t.get("expires_at", 0)) > now + 30:
        return t["access_token"]

    new_tok = _refresh_access_token(t["refresh_token"])
    _save_tokens_file(new_tok)
    return new_tok["access_token"]
