# app.py — Streamlit Xero Sales Day Book (invoices + credit notes)
import os
import sys
import io
import time
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv

# ---------- page config ----------
st.set_page_config(page_title="Xero Sales Day Book", page_icon="📒", layout="wide")

# ---------- freeze-safe base dir ----------
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

DOTENV_PATH = BASE_DIR / ".env"
TOKENS_PATH = BASE_DIR / "xero_tokens.json"
os.environ.setdefault("XERO_TOKENS_PATH", str(TOKENS_PATH))

# ---------- secrets helper (Cloud-first, local fallback) ----------
def get_secret(key: str, default: str | None = None) -> str | None:
    try:
        return st.secrets[key]  # Streamlit Cloud
    except Exception:
        return os.getenv(key, default)  # local env

# load .env only for local dev
if DOTENV_PATH.exists():
    load_dotenv(DOTENV_PATH)

# expose secrets/env for the token helper
os.environ.setdefault("XERO_CLIENT_ID",     get_secret("XERO_CLIENT_ID", "") or "")
os.environ.setdefault("XERO_CLIENT_SECRET", get_secret("XERO_CLIENT_SECRET", "") or "")
os.environ.setdefault("XERO_TENANT_ID",     get_secret("XERO_TENANT_ID", "") or "")
# OPTIONAL: if your token helper supports baseline refresh via env
os.environ.setdefault("XERO_REFRESH_TOKEN", get_secret("XERO_REFRESH_TOKEN", "") or "")

# support either module name you’ve used
try:
    from token_helper import ensure_access_token  # preferred
except Exception:
    from tokenHelper import ensure_access_token  # fallback

TENANT_ID = os.getenv("XERO_TENANT_ID")
BASE_URL = "https://api.xero.com/api.xro/2.0"

st.title("📒 Xero Sales Day Book (Streamlit)")
st.caption("Pull ACCREC invoices + ACCRECCREDIT notes, preview, and export to Excel.")

# ---------- sidebar controls ----------
with st.sidebar:
    st.header("Filters")

    year = st.number_input("Year", min_value=2000, max_value=2100, value=2025, step=1)

    default_start = date(year, 1, 1)
    default_end = date(year + 1, 1, 1)
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start date", value=default_start, format="YYYY-MM-DD")
    with col2:
        end_date = st.date_input("End date (exclusive)", value=default_end, format="YYYY-MM-DD")

    status_options = ["DRAFT", "SUBMITTED", "AUTHORISED", "PAID", "VOIDED"]
    statuses = st.multiselect("Statuses", status_options, default=["AUTHORISED", "PAID"])

    st.divider()
    exclude_contacts_raw = st.text_area(
        "Exclude contacts (comma or newline)", value="IMIS USER", height=80
    )

    st.divider()
    output_override = st.text_input(
        "Output folder (optional, absolute path)",
        value=os.getenv("REPORT_OUTPUT_DIR", ""),
        help=(
            "Leave blank to save in ./output next to this app (ephemeral on Cloud). "
            "Windows: use forward slashes or escape backslashes."
        ),
    )

    run_btn = st.button("Fetch & Export", type="primary")

# ---------- helpers ----------
def where_between(d1: date, d2: date) -> str:
    return (
        f"Date >= DateTime({d1.year},{d1.month},{d1.day}) AND "
        f"Date < DateTime({d2.year},{d2.month},{d2.day})"
    )

def status_clause(vals) -> str:
    if not vals:
        return ""
    ors = " OR ".join([f'Status=="{s}"' for s in vals])
    return f"({ors})"

def xero_get(resource: str, params=None) -> dict:
    headers = {
        "Authorization": f"Bearer {ensure_access_token()}",
        "xero-tenant-id": TENANT_ID,
        "Accept": "application/json",
    }
    r = requests.get(f"{BASE_URL}/{resource}", headers=headers, params=params or {}, timeout=60)
    if r.status_code == 429:
        time.sleep(1.0)
        r = requests.get(f"{BASE_URL}/{resource}", headers=headers, params=params or {}, timeout=60)
    r.raise_for_status()
    return r.json()

def fetch_paged(resource: str, where: str, order: str = "Date ASC", page_limit: int = 500, key_name: str = "") -> list:
    rows, page = [], 1
    while page <= page_limit:
        payload = xero_get(resource, params={"where": where, "order": order, "page": page})
        items = payload.get(key_name, [])
        if not items:
            break
        rows.extend(items)
        page += 1
    return rows

def fetch_sales_invoices(d1: date, d2: date, statuses_sel) -> list:
    parts = [f'Type=="ACCREC"', where_between(d1, d2)]
    s = status_clause(statuses_sel)
    if s: parts.append(s)
    where = " AND ".join(parts)
    return fetch_paged("Invoices", where=where, key_name="Invoices")

def fetch_sales_credit_notes(d1: date, d2: date, statuses_sel) -> list:
    parts = [f'Type=="ACCRECCREDIT"', where_between(d1, d2)]
    s = status_clause(statuses_sel)
    if s: parts.append(s)
    where = " AND ".join(parts)
    return fetch_paged("CreditNotes", where=where, key_name="CreditNotes")

def _parse_xero_dt_to_ts(x):
    if x is None:
        return pd.NaT
    if isinstance(x, str) and x.startswith("/Date("):
        try:
            ms = int(x.split("/Date(")[1].split("+")[0].rstrip(")/"))
            return pd.to_datetime(ms, unit="ms", errors="coerce")
        except Exception:
            return pd.NaT
    return pd.to_datetime(x, errors="coerce")

def _latest_payment_date_from_doc(doc: dict):
    cand = []
    for p in (doc.get("Payments") or []):
        cand.append(p.get("DateString") or p.get("Date"))
    for alloc in (doc.get("Allocations") or []):
        cand.append(alloc.get("AppliedDate"))
    if not cand:
        return None
    ts = pd.to_datetime(pd.Series(cand).apply(_parse_xero_dt_to_ts), errors="coerce")
    if ts.isna().all():
        return None
    try:
        return ts.max().date()
    except Exception:
        return None

def tidy_invoices(items: list) -> pd.DataFrame:
    rows = []
    for inv in items:
        rows.append({
            "doc_kind": "INVOICE",
            "type": inv.get("Type"),
            "number": inv.get("InvoiceNumber"),
            "contact": (inv.get("Contact") or {}).get("Name"),
            "date": inv.get("DateString") or inv.get("Date"),
            "due_date": inv.get("DueDateString") or inv.get("DueDate"),
            "payment_date": _latest_payment_date_from_doc(inv),
            "status": inv.get("Status"),
            "currency": inv.get("CurrencyCode"),
            "subtotal": inv.get("SubTotal"),
            "tax": inv.get("TotalTax"),
            "total": inv.get("Total"),
            "amount_paid": inv.get("AmountPaid"),
            "amount_due": inv.get("AmountDue"),
            "amount_credited": inv.get("AmountCredited"),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        for c in ["date", "due_date", "payment_date"]:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
        for c in ["subtotal", "tax", "total", "amount_paid", "amount_due", "amount_credited"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def tidy_credit_notes(items: list) -> pd.DataFrame:
    rows = []
    for cn in items:
        rows.append({
            "doc_kind": "CREDIT_NOTE",
            "type": cn.get("Type"),
            "number": cn.get("CreditNoteNumber"),
            "contact": (cn.get("Contact") or {}).get("Name"),
            "date": cn.get("DateString") or cn.get("Date"),
            "due_date": cn.get("DueDateString") or cn.get("DueDate"),
            "payment_date": _latest_payment_date_from_doc(cn),
            "status": cn.get("Status"),
            "currency": cn.get("CurrencyCode"),
            "subtotal": cn.get("SubTotal"),
            "tax": cn.get("TotalTax"),
            "total": cn.get("Total"),
            "amount_paid": None,
            "amount_due": None,
            "amount_credited": cn.get("AmountCredited"),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        for c in ["date", "due_date", "payment_date"]:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
        for c in ["subtotal", "tax", "total", "amount_paid", "amount_due", "amount_credited"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def clean_exclusions(df: pd.DataFrame, exclude_raw: str) -> pd.DataFrame:
    if df.empty or not exclude_raw.strip():
        return df
    tokens = [t.strip() for chunk in exclude_raw.split("\n") for t in chunk.split(",")]
    targets = set(t.upper() for t in tokens if t)
    if not targets:
        return df
    return df[~df["contact"].fillna("").str.upper().isin(targets)].copy()

@st.cache_data(show_spinner=False, ttl=300)
def fetch_and_tidy(d1: date, d2: date, statuses_sel, exclude_raw: str):
    inv_raw = fetch_sales_invoices(d1, d2, statuses_sel)
    cn_raw  = fetch_sales_credit_notes(d1, d2, statuses_sel)
    inv = tidy_invoices(inv_raw)
    cns = tidy_credit_notes(cn_raw)

    cols = [
        "doc_kind", "type", "number", "contact",
        "date", "due_date", "payment_date",
        "status", "currency",
        "subtotal", "tax", "total", "amount_paid", "amount_due", "amount_credited"
    ]
    frames = []
    if not inv.empty: frames.append(inv.reindex(columns=cols))
    if not cns.empty: frames.append(cns.reindex(columns=cols))
    df_all = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=cols)

    df_all = clean_exclusions(df_all, exclude_raw)
    df_authorised = df_all[df_all["status"] == "AUTHORISED"].copy()
    return df_all, df_authorised

def build_workbook_bytes(df_all: pd.DataFrame, df_auth: pd.DataFrame, year: int) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sheet_all = f"{year}_ALL"
        sheet_auth = f"{year}_SALES_DAY_BOOK"
        df_all.to_excel(writer, index=False, sheet_name=sheet_all)
        df_auth.to_excel(writer, index=False, sheet_name=sheet_auth)

        for sheet in (sheet_all, sheet_auth):
            ws = writer.book[sheet]
            header = [c.value for c in ws[1]]
            idx = {name: i + 1 for i, name in enumerate(header)}
            for name in ["subtotal", "tax", "total", "amount_paid", "amount_due", "amount_credited"]:
                col_i = idx.get(name)
                if col_i:
                    for col in ws.iter_cols(min_col=col_i, max_col=col_i, min_row=2, max_row=ws.max_row):
                        for cell in col:
                            cell.number_format = "#,##0.00"
            for name in ["date", "due_date", "payment_date"]:
                col_i = idx.get(name)
                if col_i:
                    for col in ws.iter_cols(min_col=col_i, max_col=col_i, min_row=2, max_row=ws.max_row):
                        for cell in col:
                            cell.number_format = "yyyy-mm-dd"
    buf.seek(0)
    return buf.getvalue()

def resolve_output_dir(override: str) -> Path:
    if override.strip():
        raw = override.strip().strip('"').strip("'")
        raw = os.path.expandvars(raw)
        safe = raw.replace("\\", "/")
        p = Path(safe).expanduser()
    else:
        p = BASE_DIR / "output"  # ephemeral on Cloud
    p.mkdir(parents=True, exist_ok=True)
    return p

# ---------- main action ----------
if run_btn:
    if not TENANT_ID:
        st.error("Missing **XERO_TENANT_ID** (set in Streamlit Secrets or .env).")
        st.stop()
    try:
        with st.spinner("Talking to Xero…"):
            df_all, df_auth = fetch_and_tidy(start_date, end_date, statuses, exclude_contacts_raw)

        c1, c2, c3 = st.columns(3)
        c1.metric("Rows (ALL)", len(df_all))
        c2.metric("Rows (AUTHORISED)", len(df_auth))
        c3.metric("Date range", f"{start_date} → {end_date} (exclusive)")

        st.subheader("Preview — ALL")
        st.dataframe(df_all, use_container_width=True, height=320)

        st.subheader("Preview — SALES DAY BOOK (Authorised)")
        st.dataframe(df_auth, use_container_width=True, height=240)

        xlsx_bytes = build_workbook_bytes(df_all, df_auth, year)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dl_name = f"sales_daybook_{year}_{stamp}.xlsx"
        st.download_button("⬇️ Download Excel", data=xlsx_bytes, file_name=dl_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        out_dir = resolve_output_dir(output_override)
        file_path = out_dir / f"sales_daybook_{year}.xlsx"
        try:
            with open(file_path, "wb") as f:
                f.write(xlsx_bytes)
            st.success(f"Saved to: `{file_path}`")
        except Exception:
            st.info("Saved to app memory only (download above). Disk writes may be blocked on Cloud.")

    except requests.HTTPError as e:
        st.error(f"HTTP error from Xero: {e} — check credentials, tenant, and scopes.")
    except Exception as e:
        st.error(f"Unexpected error: {type(e).__name__}: {e}")
else:
    st.info("Set filters in the sidebar and click **Fetch & Export**.")
