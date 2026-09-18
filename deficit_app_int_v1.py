#!/usr/bin/env python3
"""
INT SMB CSE Deficit Watermark Calculator — Streamlit App  (v1)
Connects to Salesforce via Session ID, runs INT Sales + Disconnects reports,
computes FIFO watermarks, and displays results with Market / Team / Leader / Rep
filters.  Covers UK, Canada, and Australia in a single combined view.

Org-chart hierarchy is hardcoded from the 2026 SMB quota file and used as the
primary source for Market / Team / Leader assignments.
"""

import io, time, calendar, requests
from datetime import date as _date, date
import streamlit as st
import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── Column constants ─────────────────────────────────────────────────────────
SALES_ID_COL     = "18 Digit Account ID"
SALES_AMOUNT_COL = "BMI Net New ARR (converted)"
SALES_DATE_COL   = "Order Effective Date"
SALES_OWNER_COL  = "Account Owner"
SALES_NAME_COL   = "Account Name"

TERMS_ID_COL     = "Account: 18 Digit Account ID"
TERMS_AMOUNT_COL = "Total Lost Revenue (converted)"
TERMS_DATE_COL   = "Billing End Date"
TERMS_OWNER_COL  = "Account: Account Owner"
TERMS_NAME_COL   = "Account: Account Name"

# ─── Color palette ────────────────────────────────────────────────────────────
_COL_HDR_BG   = "0070F2"
_COL_HDR_FONT = "FFFFFF"
_ALT_ROW      = "E1F4FF"
_LEADER_BG    = "00144A"
_REP_BG       = "4CB1FF"


# ─── Core FIFO watermark logic ────────────────────────────────────────────────

def _clear_date(effective_date):
    return effective_date + relativedelta(months=6)


def process_accounts(combined: pd.DataFrame) -> list:
    combined = combined.sort_values(
        ["account_id", "effective_date", "amount"],
        ascending=[True, True, True]
    ).reset_index(drop=True)

    results = []
    today   = pd.Timestamp(date.today())

    for account_id, group in combined.groupby("account_id", sort=False):
        last_row      = group.iloc[-1]
        account_owner = last_row["account_owner"]
        account_name  = last_row["account_name"]
        deficits      = []

        for _, row in group.iterrows():
            trans_date = row["effective_date"]
            amount     = float(row["amount"])
            source     = row["source"]
            deficits   = [d for d in deficits if trans_date <= d["clear_date"]]

            if amount < 0:
                deficits.append({
                    "amount":     abs(amount),
                    "eff_date":   trans_date,
                    "clear_date": _clear_date(trans_date),
                    "source":     source,
                })
            elif amount > 0:
                remaining = amount
                i = 0
                while i < len(deficits) and remaining > 0:
                    d = deficits[i]
                    if remaining >= d["amount"]:
                        remaining -= d["amount"]
                        deficits.pop(i)
                    else:
                        d["amount"] -= remaining
                        remaining = 0
                        i += 1

        deficits = [d for d in deficits if today <= d["clear_date"]]
        results.append({
            "account_owner": account_owner,
            "account_name":  account_name,
            "account_id":    account_id,
            "deficits":      deficits,
        })
    return results


# ─── Excel output helpers ─────────────────────────────────────────────────────

def _deficit_col_headers(max_deficits: int) -> list:
    headers = []
    for n in range(1, max_deficits + 1):
        headers += [
            f"Deficit {n} Amount",
            f"Deficit {n} Effective Date",
            f"Deficit {n} Clears/Expires",
            f"Deficit {n} Source",
        ]
    return headers


def _write_col_header_row(ws, col_headers: list):
    font  = Font(bold=True, color=_COL_HDR_FONT, name="Calibri", size=10)
    fill  = PatternFill(start_color=_COL_HDR_BG, end_color=_COL_HDR_BG, fill_type="solid")
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for ci, h in enumerate(col_headers, start=1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font, c.fill, c.alignment = font, fill, align
    ws.row_dimensions[1].height = 30


def _write_group_header(ws, row: int, label: str, total_cols: int,
                        bg: str, font_color: str = "FFFFFF",
                        font_size: int = 11, indent: int = 0):
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=total_cols)
    prefix = "  " * indent
    c = ws.cell(row=row, column=1, value=f"{prefix}{label}")
    c.font      = Font(bold=True, color=font_color, name="Calibri", size=font_size)
    c.fill      = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 18


def _write_account_row(ws, row: int, rep_name: str,
                       account_name: str, account_id: str,
                       deficits: list, use_alt: bool):
    alt_fill = (PatternFill(start_color=_ALT_ROW, end_color=_ALT_ROW, fill_type="solid")
                if use_alt else None)

    def _cell(col, value, num_fmt=None, h_align=None):
        c = ws.cell(row=row, column=col, value=value)
        if alt_fill:
            c.fill = alt_fill
        if num_fmt:
            c.number_format = num_fmt
        if h_align:
            c.alignment = Alignment(horizontal=h_align)
        return c

    _cell(1, rep_name)
    _cell(2, account_name)
    _cell(3, account_id)
    ci = 4
    for d in deficits:
        _cell(ci,   -round(d["amount"], 2), num_fmt='#,##0.00', h_align="right")
        _cell(ci+1, d["eff_date"].date(),   num_fmt="MM/DD/YYYY", h_align="center")
        _cell(ci+2, d["clear_date"].date(), num_fmt="MM/DD/YYYY", h_align="center")
        _cell(ci+3, d["source"])
        ci += 4


def _autofit_freeze(ws):
    for col in ws.columns:
        max_len = max(
            (len(str(cell.value)) if cell.value is not None else 0)
            for cell in col
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 45)
    ws.freeze_panes = "A2"


def _write_results_sheet(ws, df: pd.DataFrame):
    hdr_font  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    hdr_fill  = PatternFill(start_color=_COL_HDR_BG, end_color=_COL_HDR_BG, fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alt_fill  = PatternFill(start_color=_ALT_ROW,    end_color=_ALT_ROW,    fill_type="solid")

    for cell in ws[1]:
        cell.font, cell.fill, cell.alignment = hdr_font, hdr_fill, hdr_align
    ws.row_dimensions[1].height = 32

    for row_idx in range(2, ws.max_row + 1):
        if row_idx % 2 == 0:
            for cell in ws[row_idx]:
                cell.fill = alt_fill

    for col in ws.iter_cols(min_row=2, max_row=ws.max_row):
        hdr = ws.cell(1, col[0].column).value or ""
        for cell in col:
            if cell.value is None:
                continue
            if "Amount" in hdr:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            elif any(k in hdr for k in ("Date", "Expires", "Clears")):
                cell.number_format = "MM/DD/YYYY"
                cell.alignment = Alignment(horizontal="center")

    for col in ws.columns:
        max_len = max(
            (len(str(c.value)) if c.value is not None else 0) for c in col
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 45)
    ws.freeze_panes = "A2"


def _build_results_df(results: list) -> pd.DataFrame:
    rows = []
    for r in results:
        row = {
            "Account Owner":       r["account_owner"],
            "Account Name":        r["account_name"],
            "18 Digit Account ID": r["account_id"],
        }
        for n, d in enumerate(r["deficits"], start=1):
            row[f"Deficit {n} Amount"]         = -round(d["amount"], 2)
            row[f"Deficit {n} Effective Date"] = d["eff_date"].date()
            row[f"Deficit {n} Clears/Expires"] = d["clear_date"].date()
            row[f"Deficit {n} Source"]         = d["source"]
        rows.append(row)
    return pd.DataFrame(rows)


def write_excel(results: list, output_path):
    """
    Write Excel output.  Single Results sheet — no CSEs dependency.
    output_path can be a file path string or a BytesIO buffer.
    """
    active       = [r for r in results if r["deficits"]]
    df_results   = _build_results_df(active)

    with pd.ExcelWriter(output_path, engine="openpyxl",
                        date_format="MM/DD/YYYY") as writer:
        df_results.to_excel(writer, index=False, sheet_name="Results")
        _write_results_sheet(writer.sheets["Results"], df_results)

SF_API_VERSION = "v59.0"

# ─────────────────────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="INT SMB CSE Deficit Watermark Calculator",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# Configuration  — read from Streamlit secrets with hardcoded fallbacks.
# To update without touching code:
#   Streamlit Cloud → App settings → Secrets → add/edit the keys below.
# ─────────────────────────────────────────────────────────────────────────────
SALES_REPORT_ID  = st.secrets.get("SALES_REPORT_ID",  "00OPg00000Qf3xR")
TERMS_REPORT_ID  = st.secrets.get("TERMS_REPORT_ID",  "00OPg00000Qf3uD")

# ── Division column — identifies which market each record belongs to ──────────
# These column labels must match exactly what the INT Salesforce reports return.
# If results look wrong, expand the Diagnostics panel to see the actual column
# names returned and update these constants to match.
SALES_DIVISION_COL = "Account Owner: Division"
TERMS_DIVISION_COL = "Account: Account Owner: Division"

# Division string (partial match) → market label
DIVISION_MARKET_MAP = {
    "UK":        "UK SMB Client Sales",
    "Canada":    "Canada SMB Client Sales",
    "Australia": "Australia SMB Client Sales",
}

# Currency symbol per market (used in KPI cards and Amount column formatting)
MARKET_CURRENCY_SYMBOL = {
    "UK":        "£",
    "Canada":    "CA$",
    "Australia": "A$",
}

# ── Org-chart hierarchy — UK, Canada, Australia ───────────────────────────────
# LEADER_META  : leader name  →  {team, region, market}
# _RAW_ORG     : rep/FLSM name → {team, leader, market}
# Market is determined first from the org-chart lookup (rep name match), with
# the report's Account Owner Division column used as a secondary fallback for
# any rep whose name is not yet listed here.
LEADER_META = {
    # ── UK ────────────────────────────────────────────────────────────────────
    "Adam Bazeley" : {"team": "UK SMB Client Sales",           "region": "UK",                 "market": "UK"},
    "Bret Edis"    : {"team": "UK SMB Client Sales Strategic", "region": "UK Strategic/Premier","market": "UK"},
    "Katie Brown"  : {"team": "UK SMB Client Sales Key",       "region": "UK Key",              "market": "UK"},
    # ── Canada ────────────────────────────────────────────────────────────────
    "Brian Veloso" : {"team": "Canada SMB Client Sales",       "region": "Canada",              "market": "Canada"},
    "Lesley Nunes" : {"team": "Canada SMB Client Sales",       "region": "Canada",              "market": "Canada"},
    # ── Australia ─────────────────────────────────────────────────────────────
    "Fabian Calle" : {"team": "Australia SMB Client Sales",    "region": "Australia",           "market": "Australia"},
    "Peter Soukos" : {"team": "Australia SMB Client Sales",    "region": "Australia",           "market": "Australia"},
}

_RAW_ORG = {
    # ── UK — Strategic (under Bret Edis) ──────────────────────────────────────
    "Karl Perkins"          : {"team": "UK SMB Client Sales Strategic", "leader": "Bret Edis",   "market": "UK"},
    "Izabella Krawczyk Patel":{"team": "UK SMB Client Sales Strategic", "leader": "Bret Edis",   "market": "UK"},
    "Joanna Blackmore"      : {"team": "UK SMB Client Sales Strategic", "leader": "Bret Edis",   "market": "UK"},
    "Michael Benn"          : {"team": "UK SMB Client Sales Strategic", "leader": "Bret Edis",   "market": "UK"},
    "Ryan Hale"             : {"team": "UK SMB Client Sales Strategic", "leader": "Bret Edis",   "market": "UK"},
    # ── UK — Premier (under Bret Edis) ────────────────────────────────────────
    "Alan Donohoe"          : {"team": "UK SMB Client Sales Premier",   "leader": "Bret Edis",   "market": "UK"},
    "Lydia Holloway"        : {"team": "UK SMB Client Sales Premier",   "leader": "Bret Edis",   "market": "UK"},
    "Kammie Flitton"        : {"team": "UK SMB Client Sales Premier",   "leader": "Bret Edis",   "market": "UK"},
    # ── UK — Key (under Katie Brown) ──────────────────────────────────────────
    "Charlie Mason"         : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    "Adrian Sage"           : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    "Zak Jones"             : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    "Ross Greetham"         : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    "Lucy Collins"          : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    "George Smith"          : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    "Abbie Lewis"           : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    "Alexa Parritt"         : {"team": "UK SMB Client Sales Key",       "leader": "Katie Brown", "market": "UK"},
    # ── UK — FLSMs as account owners (leader = Market Leader) ─────────────────
    "Bret Edis"             : {"team": "UK SMB Client Sales Strategic", "leader": "Adam Bazeley","market": "UK"},
    "Katie Brown"           : {"team": "UK SMB Client Sales Key",       "leader": "Adam Bazeley","market": "UK"},

    # ── Canada — all CSE reps (under Lesley Nunes) ────────────────────────────
    "Brad Holder"           : {"team": "Canada SMB Client Sales Premier",   "leader": "Lesley Nunes", "market": "Canada"},
    "Chelsea Salonek"       : {"team": "Canada SMB Client Sales Premier",   "leader": "Lesley Nunes", "market": "Canada"},
    "Carol Murray"          : {"team": "Canada SMB Client Sales Strategic", "leader": "Lesley Nunes", "market": "Canada"},
    "Cristian Kawa"         : {"team": "Canada SMB Client Sales Strategic", "leader": "Lesley Nunes", "market": "Canada"},
    "Hannah Peach"          : {"team": "Canada SMB Client Sales Strategic", "leader": "Lesley Nunes", "market": "Canada"},
    "Tyler Witt"            : {"team": "Canada SMB Client Sales Key",       "leader": "Lesley Nunes", "market": "Canada"},
    "Deanna Burgess"        : {"team": "Canada SMB Client Sales Key",       "leader": "Lesley Nunes", "market": "Canada"},
    # ── Canada — FLSM as account owner (leader = Market Leader) ──────────────
    "Lesley Nunes"          : {"team": "Canada SMB Client Sales",           "leader": "Brian Veloso", "market": "Canada"},

    # ── Australia — all CSE reps (under Peter Soukos) ─────────────────────────
    "Andrew Cooksley"       : {"team": "Australia SMB Client Sales", "leader": "Peter Soukos", "market": "Australia"},
    "Kate Hulmston"         : {"team": "Australia SMB Client Sales", "leader": "Peter Soukos", "market": "Australia"},
    "Amanda Player"         : {"team": "Australia SMB Client Sales", "leader": "Peter Soukos", "market": "Australia"},
    "Steve Kavanagh"        : {"team": "Australia SMB Client Sales", "leader": "Peter Soukos", "market": "Australia"},
    "Claire van der Vegt"   : {"team": "Australia SMB Client Sales", "leader": "Peter Soukos", "market": "Australia"},
    "Hung Do"               : {"team": "Australia SMB Client Sales", "leader": "Peter Soukos", "market": "Australia"},
    # ── Australia — FLSM as account owner (leader = Market Leader) ────────────
    "Peter Soukos"          : {"team": "Australia SMB Client Sales", "leader": "Fabian Calle", "market": "Australia"},
}
# NOTE: Salesforce stores legal/full names; add aliases below if a rep's
# Salesforce name differs from the quota file (same pattern as the US app).
# Example:
#   "Izabella Krawczyk-Patel": {"team": "UK SMB Client Sales Strategic",
#                                "leader": "Bret Edis", "market": "UK"},

# Pre-build a lowercase-keyed lookup for fast, case-insensitive rep matching.
_ORG_LOOKUP = {k.lower(): v for k, v in _RAW_ORG.items()}

# ─────────────────────────────────────────────────────────────────────────────
# Styles
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: '72', Arial, sans-serif; }

.app-header {
    background: #0070F2;
    color: #fff;
    padding: 1rem 1.5rem;
    border-radius: 8px;
    margin-bottom: 1.25rem;
}
.app-header h1 { margin: 0; font-size: 1.35rem; font-weight: 700; color: #fff; }
.app-header p  { margin: 0.2rem 0 0; font-size: 0.82rem; opacity: 0.88; }

.kpi-wrap { display: flex; gap: 0.85rem; margin: 0.75rem 0 1.25rem; }
.kpi-card {
    flex: 1;
    background: #E1F4FF;
    border-left: 4px solid #0070F2;
    border-radius: 6px;
    padding: 0.75rem 1rem;
}
.kpi-val { font-size: 1.75rem; font-weight: 700; color: #00144A; line-height: 1.15; }
.kpi-lbl { font-size: 0.72rem; color: #555; margin-top: 4px; letter-spacing: 0.04em; text-transform: uppercase; }

.filter-wrap {
    background: #EAECEE;
    border-radius: 8px;
    padding: 0.75rem 1rem 0.25rem;
    margin-bottom: 0.75rem;
}

section[data-testid="stSidebar"] { background: #f5f6f7; }
section[data-testid="stSidebar"] .stMarkdown h3 { color: #00144A; }

.stButton > button {
    background: #0070F2 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
}
.stButton > button:hover { background: #0134BF !important; }
.stButton > button:disabled {
    background: #EAECEE !important;
    color: #aaa !important;
    cursor: not-allowed !important;
}
.stDownloadButton > button {
    background: #fff !important;
    color: #0070F2 !important;
    border: 1.5px solid #0070F2 !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
}
.stDownloadButton > button:hover {
    background: #E1F4FF !important;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Salesforce helpers
# ─────────────────────────────────────────────────────────────────────────────

def _hdr(session_id: str) -> dict:
    return {"Authorization": f"Bearer {session_id}", "Content-Type": "application/json"}


def detect_instance_url(session_id: str):
    """
    Auto-detect the Salesforce instance URL from a session ID alone.

    Calls the standard OAuth2 userinfo endpoint at login.salesforce.com (and
    test.salesforce.com as a fallback for sandboxes).  Salesforce returns the
    org's REST base URL in the response regardless of whether the org uses a
    custom domain, so the correct instance URL is always recovered.

    Returns (ok: bool, instance_url: str, info: dict).
    """
    from urllib.parse import urlparse
    for base in ("https://login.salesforce.com", "https://test.salesforce.com"):
        try:
            r = requests.get(
                f"{base}/services/oauth2/userinfo",
                headers=_hdr(session_id),
                timeout=10,
                allow_redirects=True,
            )
            if r.status_code == 200:
                info = r.json()
                # urls.rest is "https://<instance>/services/data/" — parse the origin
                rest = info.get("urls", {}).get("rest", "")
                if rest:
                    p = urlparse(rest)
                    return True, f"{p.scheme}://{p.netloc}", info
        except Exception:
            pass
    return False, "", {}


def search_reports(instance_url: str, session_id: str, term: str) -> list:
    term_safe = term.replace("'", "\\'")
    q = f"SELECT Id, Name FROM Report WHERE Name LIKE '%{term_safe}%' LIMIT 30"
    r = requests.get(
        f"{instance_url}/services/data/{SF_API_VERSION}/query",
        headers=_hdr(session_id), params={"q": q}, timeout=15,
    )
    r.raise_for_status()
    return r.json().get("records", [])


def _parse_sf_number(raw: str):
    """
    Parse a Salesforce-formatted number string to float.
    Handles:
      "$1,234.56"   →  1234.56
      "(1,234.56)"  →  -1234.56   (accounting notation for negatives)
      "($1,234.56)" →  -1234.56
      "-1234.56"    →  -1234.56
    Returns None if unparseable.
    """
    s = str(raw).replace(",", "").replace("$", "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _parse_report_response(data: dict):
    """
    Parse an Analytics API report instance response → (DataFrame, all_data: bool).

    factMap layout differs by report format:
      TABULAR  — all detail rows under "T!T"
      SUMMARY  — detail rows distributed across group keys ("0!T", "0_0!T", …);
                 "T!T" holds only grand-total *aggregates* (no rows).

    This function iterates every factMap node and collects all rows,
    so it works for both formats without needing to know which one is in use.
    """
    meta        = data.get("reportMetadata", {})
    ext         = data.get("reportExtendedMetadata", {})
    detail_cols = meta.get("detailColumns", [])
    col_info    = ext.get("detailColumnInfo", {})
    col_labels  = [col_info.get(c, {}).get("label", c) for c in detail_cols]
    numeric_t   = {"currency", "double", "int", "percent", "number"}
    col_types   = {c: col_info.get(c, {}).get("dataType", "string") for c in detail_cols}

    # Collect rows from every node in the factMap (handles tabular + summary).
    all_rows = []
    for node in data.get("factMap", {}).values():
        if isinstance(node, dict):
            all_rows.extend(node.get("rows", []))

    parsed = []
    for row in all_rows:
        pr = []
        for i, cell in enumerate(row.get("dataCells", [])):
            col_key = detail_cols[i] if i < len(detail_cols) else ""
            dtype   = col_types.get(col_key, "string")
            if dtype in numeric_t:
                val = cell.get("value")
                # Multi-currency fields return a dict: {"amount": 21168, "currency": "USD"}
                # Extract the numeric amount from it.
                if isinstance(val, dict):
                    val = val.get("amount")
                # value=null fallback: parse the formatted label string
                if val is None:
                    val = _parse_sf_number(cell.get("label", ""))
                pr.append(val)
            else:
                pr.append(cell.get("label"))
        parsed.append(pr)

    return pd.DataFrame(parsed, columns=col_labels), data.get("allData", True)


@st.cache_data(ttl=300, show_spinner=False)
def run_sf_report(instance_url: str, session_id: str, report_id: str,
                  date_col_label: str = ""):
    """
    Fetch ALL report rows — no 2,000-row ceiling.  Three-tier strategy:

    Tier 1 — SOQL + pagination
        GET the report metadata to discover the Salesforce object type and each
        column's SOQL field name (entityColumnName).  Issues a paginated /query
        and follows every nextRecordsUrl until done.  Unlimited row count.
        Works only when every column has a real SOQL equivalent (no formula /
        converted-currency columns).

    Tier 2 — Chunked Analytics API  ← the key fix for this report
        When converted-currency or formula columns have no entityColumnName,
        SOQL is unavailable.  Instead, split the date range into monthly windows
        and execute one async report instance per window.  Each month stays well
        under the 2,000-row cap; all chunks are concatenated.
        Requires date_col_label to be the exact column header of the date field
        in the report (e.g. "Order Effective Date").  The column's Analytics ID
        is resolved from the report metadata so no hardcoding is needed.
        The report's own existing filters are preserved; only a tight date
        bracket is added on top for each chunk.

    Tier 3 — plain async Analytics API  (last resort, 2,000-row cap)
        Used only when both SOQL and chunking fail.  Caller receives
        all_data=False as a signal.

    Returns (DataFrame, all_data: bool).
    """
    h   = _hdr(session_id)
    api = f"{instance_url}/services/data/{SF_API_VERSION}/analytics/reports/{report_id}"

    # ── Tier 0: fetch report metadata (always needed) ────────────────────────
    meta_r = requests.get(api, headers=h, timeout=15)
    meta_r.raise_for_status()
    m           = meta_r.json()
    rpt_meta    = m.get("reportMetadata", {})
    ext_meta    = m.get("reportExtendedMetadata", {})
    detail_cols = rpt_meta.get("detailColumns", [])
    col_info    = ext_meta.get("detailColumnInfo", {})
    col_labels  = [col_info.get(c, {}).get("label", c) for c in detail_cols]
    obj_type    = rpt_meta.get("reportType", {}).get("type", "")
    ent_cols    = [
        col_info.get(c, {}).get("entityColumnName", "") or ""
        for c in detail_cols
    ]

    # showDetails is only valid for Summary / Matrix reports.
    # Sending it for a Tabular report causes a 400 Bad Request.
    report_fmt   = rpt_meta.get("reportFormat", "TABULAR")
    needs_detail = report_fmt in ("SUMMARY", "MATRIX")

    # ── Tier 1: SOQL path ────────────────────────────────────────────────────
    if obj_type and all(ent_cols):
        try:
            soql     = f"SELECT {', '.join(ent_cols)} FROM {obj_type}"
            all_rows = []
            next_url = None
            while True:
                if next_url:
                    resp = requests.get(f"{instance_url}{next_url}",
                                        headers=h, timeout=30)
                else:
                    resp = requests.get(
                        f"{instance_url}/services/data/{SF_API_VERSION}/query",
                        headers=h, params={"q": soql}, timeout=30,
                    )
                resp.raise_for_status()
                data = resp.json()
                for rec in data.get("records", []):
                    row = []
                    for field in ent_cols:
                        val = rec
                        for part in field.split("."):
                            val = val.get(part) if isinstance(val, dict) else None
                        row.append(val)
                    all_rows.append(row)
                if data.get("done", True):
                    break
                next_url = data.get("nextRecordsUrl")
            return pd.DataFrame(all_rows, columns=col_labels), True, 1, []
        except Exception:
            pass  # fall through

    # ── Tier 2: chunked Analytics API ────────────────────────────────────────
    # Resolve the Analytics column ID for the date field by matching its label.
    date_col_id = None
    if date_col_label:
        for col_id in detail_cols:
            if col_info.get(col_id, {}).get("label", "") == date_col_label:
                date_col_id = col_id
                break

    if date_col_id:
        today   = _date.today()
        all_dfs = []
        t2_errs = []
        cap_hit = False

        # ── Tier 2: async chunked Analytics API ──────────────────────────────
        # Each monthly window is POSTed to /instances (async endpoint) instead
        # of the synchronous /reports/{id} endpoint.  This uses the async limit
        # (1,200 instances/hour) rather than the sync limit (500/hour), which
        # matters when 20+ leaders are running the app concurrently.
        # Each chunk stays well under the 2,000-row async cap because it covers
        # only one calendar month.
        # standardDateFilter overrides only the date range; all saved report
        # filters (e.g. division) remain intact.
        for i in range(9):
            if cap_hit:
                break
            y, mo = today.year, today.month - i
            while mo <= 0:
                mo += 12; y -= 1
            _, last = calendar.monthrange(y, mo)
            start = f"{y:04d}-{mo:02d}-01"
            end   = f"{y:04d}-{mo:02d}-{last:02d}"

            chunk_meta = {
                "standardDateFilter": {
                    "column":        date_col_id,
                    "durationValue": "CUSTOM",
                    "startDate":     start,
                    "endDate":       end,
                }
            }
            if needs_detail:
                chunk_meta["showDetails"] = True

            # Fire the async instance
            post_r = requests.post(
                f"{api}/instances",
                headers=h,
                json={"reportMetadata": chunk_meta},
                timeout=30,
            )

            if post_r.status_code >= 400:
                try:
                    err_body = post_r.json()
                except Exception:
                    err_body = post_r.text[:800]
                t2_errs.append({
                    "month":  f"{y}-{mo:02d}",
                    "status": post_r.status_code,
                    "body":   err_body,
                })
                continue

            inst_id = post_r.json().get("id")
            if not inst_id:
                t2_errs.append({
                    "month":  f"{y}-{mo:02d}",
                    "status": "no_id",
                    "body":   str(post_r.json())[:400],
                })
                continue

            # Poll until complete (up to 60 × 2s = 120s per chunk)
            for _ in range(60):
                time.sleep(2)
                poll = requests.get(
                    f"{api}/instances/{inst_id}",
                    headers=h,
                    timeout=30,
                )
                poll.raise_for_status()
                payload = poll.json()
                status  = payload.get("attributes", {}).get("status", "")

                if status == "Success":
                    chunk_df, chunk_all = _parse_report_response(payload)
                    if not chunk_all:
                        t2_errs.append({
                            "month":  f"{y}-{mo:02d}",
                            "status": "cap_exceeded",
                            "body":   "Single month > 2,000 rows; weekly chunking would be needed.",
                        })
                        all_dfs  = []
                        cap_hit  = True
                    elif not chunk_df.empty:
                        all_dfs.append(chunk_df)
                    break

                if status in ("Error", "Cancelled"):
                    code = payload.get("attributes", {}).get("errorCode", "Unknown")
                    t2_errs.append({
                        "month":  f"{y}-{mo:02d}",
                        "status": f"async_{status}",
                        "body":   code,
                    })
                    break
            else:
                t2_errs.append({
                    "month":  f"{y}-{mo:02d}",
                    "status": "timeout",
                    "body":   "Async instance did not complete within 120s.",
                })

        if all_dfs:
            return pd.concat(all_dfs, ignore_index=True), True, 2, t2_errs

        # All chunks failed or empty — fall through to Tier 3.

    # ── Tier 3: plain async Analytics API (2,000-row cap) ────────────────────
    t2_errs = t2_errs if "t2_errs" in dir() else []
    t3_body = {"reportMetadata": {"showDetails": True}} if needs_detail else {}
    r = requests.post(f"{api}/instances", headers=h, json=t3_body, timeout=30)
    r.raise_for_status()
    inst_id = r.json()["id"]
    for _ in range(90):
        time.sleep(2)
        r = requests.get(f"{api}/instances/{inst_id}", headers=h, timeout=30)
        r.raise_for_status()
        payload = r.json()
        status  = payload.get("attributes", {}).get("status", "")
        if status == "Success":
            df, all_data = _parse_report_response(payload)
            return df, all_data, 3, t2_errs
        if status in ("Error", "Cancelled"):
            code = payload.get("attributes", {}).get("errorCode", "Unknown")
            raise RuntimeError(f"Report failed: {code}")
    raise TimeoutError("Report timed out after 3 minutes.")


# ─────────────────────────────────────────────────────────────────────────────
# Data prep (DataFrame-based, replacing the file-path loaders)
# ─────────────────────────────────────────────────────────────────────────────

def _coerce_numeric(s: pd.Series) -> pd.Series:
    """
    Convert a Series to numeric, handling all Salesforce Analytics API formats:
      21168                          →  21168.0   (int from multi-currency dict)
      {"amount": 21168, ...}         →  21168.0   (dict not yet unwrapped)
      "$1,234.56"                    →  1234.56
      "(1,234.56)" / "($1,234.56)"   →  -1234.56  (accounting notation)
    """
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    def _extract(x):
        if isinstance(x, dict):          # {"amount": 21168, "currency": "USD"}
            return x.get("amount")
        return _parse_sf_number(x)

    return pd.to_numeric(s.map(_extract), errors="coerce")


def prep_sales(df: pd.DataFrame) -> pd.DataFrame:
    need = [SALES_ID_COL, SALES_AMOUNT_COL, SALES_DATE_COL, SALES_OWNER_COL, SALES_NAME_COL]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"Sales report is missing columns: {miss}\nColumns found: {list(df.columns)}")
    df = df.rename(columns={
        SALES_ID_COL: "account_id", SALES_AMOUNT_COL: "amount",
        SALES_DATE_COL: "effective_date", SALES_OWNER_COL: "account_owner",
        SALES_NAME_COL: "account_name",
    })[["account_id", "amount", "effective_date", "account_owner", "account_name"]].copy()
    df["source"]         = "Sales"
    df["amount"]         = _coerce_numeric(df["amount"])
    df["effective_date"] = pd.to_datetime(df["effective_date"], errors="coerce")
    return df[
        df["amount"].notna() & (df["amount"] != 0)
        & df["effective_date"].notna()
        & df["account_id"].notna()
        & (df["account_id"].astype(str).str.strip() != "")
    ].reset_index(drop=True)


def prep_terms(df: pd.DataFrame) -> pd.DataFrame:
    need = [TERMS_ID_COL, TERMS_AMOUNT_COL, TERMS_DATE_COL, TERMS_OWNER_COL, TERMS_NAME_COL]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"Terminations report is missing columns: {miss}\nColumns found: {list(df.columns)}")
    df = df.rename(columns={
        TERMS_ID_COL: "account_id", TERMS_AMOUNT_COL: "amount",
        TERMS_DATE_COL: "effective_date", TERMS_OWNER_COL: "account_owner",
        TERMS_NAME_COL: "account_name",
    })[["account_id", "amount", "effective_date", "account_owner", "account_name"]].copy()
    df["source"]         = "Terminations"
    df["amount"]         = _coerce_numeric(df["amount"]).abs() * -1
    df["effective_date"] = pd.to_datetime(df["effective_date"], errors="coerce")
    return df[
        df["amount"].notna() & (df["amount"] != 0)
        & df["effective_date"].notna()
        & df["account_id"].notna()
        & (df["account_id"].astype(str).str.strip() != "")
    ].reset_index(drop=True)


def prep_cses(df: pd.DataFrame) -> pd.DataFrame:
    need = [CSES_OWNER_COL, CSES_MANAGER_COL, CSES_TEAM_COL]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f"CSEs report is missing columns: {miss}\nColumns found: {list(df.columns)}")
    df = df.rename(columns={
        CSES_OWNER_COL: "owner_name", CSES_MANAGER_COL: "manager_name", CSES_TEAM_COL: "team",
    })[["owner_name", "manager_name", "team"]].copy()
    df["owner_name"]   = df["owner_name"].str.strip()
    df["manager_name"] = df["manager_name"].fillna("Unassigned").str.strip()
    df["team"]         = df["team"].fillna("").str.strip()
    return (df.dropna(subset=["owner_name"])
              .drop_duplicates(subset=["owner_name"])
              .reset_index(drop=True))


# ─────────────────────────────────────────────────────────────────────────────
# Master DataFrame — one row per account, all context + deficit columns
# ─────────────────────────────────────────────────────────────────────────────

def build_master_df(active: list, div_lookup: dict = None) -> pd.DataFrame:
    """
    Build the master DataFrame from process_accounts results.

    div_lookup: optional dict of account_id → division string (from raw report
    data).  Used only as a fallback to determine Market for reps not yet in the
    hardcoded org chart.
    """
    div_lookup = div_lookup or {}

    rows = []
    for r in active:
        rep       = (r["account_owner"] or "").strip()
        rep_lower = rep.lower()

        # Primary: org-chart lookup (case-insensitive exact match on rep name)
        org = _ORG_LOOKUP.get(rep_lower)
        if org:
            team   = org["team"]
            leader = org["leader"]
            region = LEADER_META.get(leader, {}).get("region", "")
            market = org.get("market", "")
        else:
            # Fallback: infer market from the division column in the raw report
            team   = ""
            leader = "Unassigned"
            region = ""
            market = ""
            div = div_lookup.get(str(r["account_id"]).strip(), "")
            for mkt, div_substr in DIVISION_MARKET_MAP.items():
                if div_substr in str(div):
                    market = mkt
                    break

        base = {
            "Market":       market,
            "Team":         team,
            "Leader":       leader,
            "Region":       region,
            "Rep":          rep,
            "Account Name": r["account_name"],
            "Account ID":   r["account_id"],
        }
        for n, d in enumerate(r["deficits"], 1):
            base[f"Deficit {n} Amount"]         = -round(d["amount"], 2)
            base[f"Deficit {n} Effective Date"] = (
                d["eff_date"].date()
                if hasattr(d["eff_date"], "date") else d["eff_date"]
            )
            base[f"Deficit {n} Clears/Expires"] = (
                d["clear_date"].date()
                if hasattr(d["clear_date"], "date") else d["clear_date"]
            )
            base[f"Deficit {n} Source"]         = d["source"]
        rows.append(base)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    return df.sort_values(["Market", "Team", "Leader", "Rep", "Account Name"]).reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# Excel download helper
# ─────────────────────────────────────────────────────────────────────────────

def excel_bytes(results: list) -> bytes:
    active = [r for r in results if r["deficits"]]
    if not active:
        return b""
    buf = io.BytesIO()
    write_excel(results, buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar — Connection + Report Selection
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### Salesforce Connection")

    session_id = st.text_input(
        "Session ID",
        type="password",
        placeholder="00D…",
        key="sf_sid",
    )

    if st.button("Connect", use_container_width=True):
        if not session_id.strip():
            st.error("Session ID is required.")
        else:
            with st.spinner("Detecting instance and verifying session…"):
                ok, instance_url, info = detect_instance_url(session_id.strip())
            if ok:
                st.session_state.update({
                    "connected":   True,
                    "sf_instance": instance_url,
                    "sf_token":    session_id.strip(),
                })
                name = info.get("name") or info.get("preferred_username", "")
                st.success(f"Connected{' — ' + name if name else ''}")
            else:
                st.session_state["connected"] = False
                st.error("Connection failed. Check your Session ID.")

    if st.session_state.get("connected"):
        st.markdown("---")
        run_btn = st.button("Run Calculation", use_container_width=True)

        if st.session_state.get("results_ready"):
            if st.button("Refresh Data", use_container_width=True):
                run_sf_report.clear()
                for k in ["results_ready", "results", "active", "master_df",
                          "run_date", "warnings", "div_lookup"]:
                    st.session_state.pop(k, None)
                st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# Page header
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>INT SMB CSE Deficit Watermark Calculator</h1>
  <p>Live FIFO netting of SOF losses and disconnects &mdash; 6-month watermark windows per account &mdash; UK &bull; Canada &bull; Australia</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Pre-connection / pre-run states
# ─────────────────────────────────────────────────────────────────────────────
if not st.session_state.get("connected"):
    st.info("Enter your Salesforce Session ID in the sidebar to get started.")
    with st.expander("How to get your Session ID"):
        st.markdown("""
**Option A — Developer Console**
1. Log in to Salesforce and open the gear menu → **Developer Console**.
2. Go to **Debug → Open Execute Anonymous Window**.
3. Paste and run: `System.debug(UserInfo.getSessionId());`
4. Open the log and copy the value after `DEBUG|`.

**Option B — Browser URL** *(Classic only)*
The session ID may appear in the page URL after `sid=`.

The session ID expires when you log out or after your org's session timeout (typically 8 hours).
The instance URL is detected automatically — no need to enter it manually.
        """)
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Run reports & compute watermarks
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.get("connected") and "run_btn" in dir() and run_btn:
    bar   = st.progress(0, "Fetching Sales report…")
    warns = []
    try:
        with st.spinner("Running Sales report (monthly chunks)…"):
            s_df, s_all, s_tier, s_t2errs = run_sf_report(
                st.session_state["sf_instance"], st.session_state["sf_token"],
                SALES_REPORT_ID, date_col_label=SALES_DATE_COL,
            )
        if not s_all:
            warns.append(
                "Sales report used the plain Analytics API (2,000-row cap). "
                "Monthly chunking produced no rows — see Tier 2 errors in Diagnostics "
                "for the exact Salesforce rejection reason."
            )
        bar.progress(30, "Fetching Disconnects report…")

        with st.spinner("Running Disconnects report (monthly chunks)…"):
            t_df, t_all, t_tier, t_t2errs = run_sf_report(
                st.session_state["sf_instance"], st.session_state["sf_token"],
                TERMS_REPORT_ID, date_col_label=TERMS_DATE_COL,
            )
        if not t_all:
            warns.append(
                "Disconnects report used the plain Analytics API (2,000-row cap). "
                "Monthly chunking produced no rows — see Tier 2 errors in Diagnostics "
                "for the exact Salesforce rejection reason."
            )
        bar.progress(60, "Processing data…")

        # ── Build division lookup (account_id → division) before prep ─────────
        # Used as fallback market attribution for reps not in the org chart.
        div_lookup: dict = {}
        for raw_df, id_col, div_col in [
            (s_df, SALES_ID_COL,  SALES_DIVISION_COL),
            (t_df, TERMS_ID_COL,  TERMS_DIVISION_COL),
        ]:
            if div_col in raw_df.columns and id_col in raw_df.columns:
                for _, row in raw_df[[id_col, div_col]].dropna(subset=[id_col]).iterrows():
                    aid = str(row[id_col]).strip()
                    div = str(row[div_col]).strip() if pd.notna(row.get(div_col)) else ""
                    if aid and div and aid not in div_lookup:
                        div_lookup[aid] = div

        bar.progress(70, "Computing watermarks…")

        sales_df  = prep_sales(s_df)
        terms_df  = prep_terms(t_df)
        combined  = pd.concat([sales_df, terms_df], ignore_index=True)
        results   = process_accounts(combined)
        active    = [r for r in results if r["deficits"]]
        master_df = build_master_df(active, div_lookup=div_lookup)

        bar.progress(100, "Done.")
        bar.empty()

        st.session_state.update({
            "results_ready": True,
            "results":       results,
            "active":        active,
            "master_df":     master_df,
            "div_lookup":    div_lookup,
            "run_date":      date.today(),
            "warnings":      warns,
            "diag": {
                "sales_rows":    len(s_df),
                "sales_cols":    list(s_df.columns),
                "sales_sample":  s_df.iloc[0].to_dict() if not s_df.empty else None,
                "sales_all":     s_all,
                "sales_tier":    s_tier,
                "sales_t2errs":  s_t2errs,
                "terms_rows":    len(t_df),
                "terms_cols":    list(t_df.columns),
                "terms_sample":  t_df.iloc[0].to_dict() if not t_df.empty else None,
                "terms_all":     t_all,
                "terms_tier":    t_tier,
                "terms_t2errs":  t_t2errs,
                "sales_prepped": len(sales_df),
                "terms_prepped": len(terms_df),
                "combined_rows": len(combined),
                "results_total": len(results),
                "active_total":  len(active),
            },
        })

    except Exception as e:
        bar.empty()
        st.error(f"Error: {e}")
        st.stop()


if not st.session_state.get("results_ready"):
    st.info("Click **Run Calculation** in the sidebar to load data.")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Results
# ─────────────────────────────────────────────────────────────────────────────

for w in st.session_state.get("warnings", []):
    st.warning(w)

# ── Diagnostics panel ─────────────────────────────────────────────────────────
if st.session_state.get("diag"):
    d = st.session_state["diag"]
    with st.expander("Diagnostics — expand if results look wrong", expanded=(d["active_total"] == 0)):
        c1, c2, c3 = st.columns(3)
        c1.metric("Sales rows fetched",    f"{d['sales_rows']:,}")
        c1.metric("Sales rows after prep", f"{d['sales_prepped']:,}")
        c2.metric("Terms rows fetched",    f"{d['terms_rows']:,}")
        c2.metric("Terms rows after prep", f"{d['terms_prepped']:,}")
        c3.metric("Combined rows",         f"{d['combined_rows']:,}")
        c3.metric("Accounts with deficits",f"{d['active_total']:,}")

        tier_s = "Chunked (Tier 2)" if d.get("sales_tier") == 2 else "Plain async / Tier 3 (≤2,000 rows)"
        tier_t = "Chunked (Tier 2)" if d.get("terms_tier") == 2 else "Plain async / Tier 3 (≤2,000 rows)"
        st.caption(f"Fetch strategy — Sales: **{tier_s}** | Terms: **{tier_t}**")

        # If Tier 2 failed for any chunks, show the exact Salesforce errors so
        # the filter-rejection reason is visible without guessing.
        for label, err_key in [("Sales", "sales_t2errs"), ("Terms", "terms_t2errs")]:
            errs = d.get(err_key) or []
            if errs:
                with st.expander(f"Tier 2 chunk errors — {label} ({len(errs)} failed)"):
                    st.json(errs)

        st.markdown("**Expected column names**")
        exp_s = [SALES_ID_COL, SALES_AMOUNT_COL, SALES_DATE_COL, SALES_OWNER_COL, SALES_NAME_COL]
        exp_t = [TERMS_ID_COL, TERMS_AMOUNT_COL, TERMS_DATE_COL, TERMS_OWNER_COL, TERMS_NAME_COL]
        miss_s = [c for c in exp_s if c not in d["sales_cols"]]
        miss_t = [c for c in exp_t if c not in d["terms_cols"]]

        st.markdown("*Sales report columns returned from Salesforce:*")
        st.code(", ".join(d["sales_cols"]) or "(none)")
        if miss_s:
            st.error(f"Sales: missing expected columns — {miss_s}")

        st.markdown("*Terminations report columns returned from Salesforce:*")
        st.code(", ".join(d["terms_cols"]) or "(none)")
        if miss_t:
            st.error(f"Terminations: missing expected columns — {miss_t}")

        # Raw sample — shows actual cell values so numeric-format issues are visible
        if d.get("sales_sample"):
            st.markdown("*Sales report — first raw row (before prep):*")
            st.json(d["sales_sample"])
        if d.get("terms_sample"):
            st.markdown("*Terminations report — first raw row (before prep):*")
            st.json(d["terms_sample"])

master_df = st.session_state["master_df"]
results   = st.session_state["results"]
run_date  = st.session_state["run_date"]
has_org   = not master_df.empty and "Leader" in master_df.columns


# ── Filters ──────────────────────────────────────────────────────────────────
st.markdown('<div class="filter-wrap">', unsafe_allow_html=True)

# Four-level cascade: Market → Team → Leader → Account Owner
fc = st.columns([0.9, 1.1, 1.1, 1.1, 1.6, 0.5])

# Level 0 — Market
all_markets = sorted(master_df["Market"].dropna().replace("", pd.NA).dropna().unique()) if has_org else []
sel_markets = fc[0].multiselect(
    "Market",
    all_markets,
    placeholder="All markets",
    key="f_market",
)

# Level 1 — Team (cascades from market)
if has_org:
    team_pool_df = master_df[master_df["Market"].isin(sel_markets)] if sel_markets else master_df
    all_teams = sorted(team_pool_df["Team"].dropna().replace("", pd.NA).dropna().unique())
else:
    all_teams = []
sel_teams = fc[1].multiselect(
    "Team",
    all_teams,
    placeholder="All teams",
    key="f_team",
)

# Level 2 — Leader (cascades from team, then market)
if has_org:
    if sel_teams:
        ldr_pool_df = master_df[master_df["Team"].isin(sel_teams)]
    elif sel_markets:
        ldr_pool_df = master_df[master_df["Market"].isin(sel_markets)]
    else:
        ldr_pool_df = master_df
    leader_display: dict[str, str] = {}
    for ldr in ldr_pool_df["Leader"].dropna().replace("", pd.NA).dropna().unique():
        region_vals = ldr_pool_df[ldr_pool_df["Leader"] == ldr]["Region"].dropna()
        region_str  = region_vals.iloc[0] if not region_vals.empty else ""
        leader_display[ldr] = f"{ldr} ({region_str})" if region_str else ldr
    all_leaders_display = sorted(leader_display.values())
    display_to_leader   = {v: k for k, v in leader_display.items()}
else:
    all_leaders_display = []
    display_to_leader   = {}

sel_leaders_display = fc[2].multiselect(
    "Leader",
    all_leaders_display,
    placeholder="All leaders",
    key="f_leader",
)
sel_leaders = [display_to_leader.get(d, d) for d in sel_leaders_display]

# Level 3 — Account Owner (cascades from leader → team → market → all)
if sel_leaders:
    rep_pool = master_df[master_df["Leader"].isin(sel_leaders)]["Rep"]
elif sel_teams:
    rep_pool = master_df[master_df["Team"].isin(sel_teams)]["Rep"]
elif sel_markets:
    rep_pool = master_df[master_df["Market"].isin(sel_markets)]["Rep"]
else:
    rep_pool = master_df["Rep"] if not master_df.empty else pd.Series([], dtype=str)
all_reps = sorted(rep_pool.dropna().unique())
sel_reps = fc[3].multiselect(
    "Account Owner",
    all_reps,
    placeholder="All reps",
    key="f_rep",
)

search = fc[4].text_input(
    "Search account name", placeholder="Type to search…", key="f_search"
)
if fc[5].button("Clear", use_container_width=True, key="f_clear"):
    for k in ["f_market", "f_team", "f_leader", "f_rep", "f_search"]:
        st.session_state.pop(k, None)
    st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Apply filters
filtered = master_df.copy()
if sel_markets:
    filtered = filtered[filtered["Market"].isin(sel_markets)]
if sel_teams:
    filtered = filtered[filtered["Team"].isin(sel_teams)]
if sel_leaders:
    filtered = filtered[filtered["Leader"].isin(sel_leaders)]
if sel_reps:
    filtered = filtered[filtered["Rep"].isin(sel_reps)]
if search:
    filtered = filtered[filtered["Account Name"].str.contains(search, case=False, na=False)]

filtered = filtered.reset_index(drop=True)

# Determine currency symbol for KPI display
# Use the selected market's symbol when exactly one market is active; otherwise generic.
active_markets = sel_markets if sel_markets else all_markets
if len(active_markets) == 1:
    cur_sym = MARKET_CURRENCY_SYMBOL.get(active_markets[0], "$")
else:
    cur_sym = ""  # mixed currencies — omit symbol


# ── KPI cards ────────────────────────────────────────────────────────────────
amt_cols         = [c for c in filtered.columns if "Amount" in c]
total_amt        = filtered[amt_cols].apply(pd.to_numeric, errors="coerce").sum().sum()
total_watermarks = int(filtered[amt_cols].apply(pd.to_numeric, errors="coerce").notna().sum().sum())

st.markdown(f"""
<div class="kpi-wrap">
  <div class="kpi-card">
    <div class="kpi-val">{len(filtered):,}</div>
    <div class="kpi-lbl">Accounts with Deficits</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-val">{cur_sym}{abs(total_amt):,.0f}</div>
    <div class="kpi-lbl">Total Deficit Amount</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-val">{total_watermarks:,}</div>
    <div class="kpi-lbl">Active Watermarks</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-val">{run_date.strftime("%b %d, %Y")}</div>
    <div class="kpi-lbl">Data As Of</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Table ─────────────────────────────────────────────────────────────────────
if filtered.empty:
    st.info("No accounts match the current filters.")
else:
    # ── Dynamic column trimming ───────────────────────────────────────────────
    # Only show as many Deficit N groups as the filtered rows actually contain.
    # This prevents empty Amount/Clears columns from squeezing the visible data.
    populated_ns = set()
    for col in filtered.columns:
        if col.startswith("Deficit ") and " Amount" in col:
            try:
                n = int(col.split()[1])
                if pd.to_numeric(filtered[col], errors="coerce").notna().any():
                    populated_ns.add(n)
            except (ValueError, IndexError):
                pass
    max_deficit_n = max(populated_ns) if populated_ns else 0

    # Identity columns — Market and Team kept in master_df for filtering but not displayed
    identity_cols = [
        c for c in ["Leader", "Rep", "Account Name", "Account ID"]
        if c in filtered.columns
    ]
    # Deficit group columns in order, capped at max_deficit_n
    deficit_group_order = ("Amount", "Effective Date", "Clears/Expires", "Source")
    deficit_cols = []
    for n in range(1, max_deficit_n + 1):
        for suffix in deficit_group_order:
            cname = f"Deficit {n} {suffix}"
            if cname in filtered.columns:
                deficit_cols.append(cname)

    display_cols = identity_cols + deficit_cols

    # ── Column config ─────────────────────────────────────────────────────────
    # Dynamic currency format for Amount columns
    amt_fmt = f"{cur_sym}%.2f" if cur_sym else "%.2f"

    col_cfg = {}
    col_cfg["Leader"]       = st.column_config.TextColumn("Leader",       width="medium")
    col_cfg["Rep"]          = st.column_config.TextColumn("Rep",          width="medium")
    col_cfg["Account Name"] = st.column_config.TextColumn("Account Name", width="large")
    col_cfg["Account ID"]   = st.column_config.TextColumn("Account ID",   width="medium")

    for c in display_cols:
        if "Amount" in c:
            col_cfg[c] = st.column_config.NumberColumn(c, format=amt_fmt,        width="medium")
        elif "Effective Date" in c or "Clears" in c:
            col_cfg[c] = st.column_config.DateColumn(c,   format="MM/DD/YYYY",   width="medium")
        elif "Source" in c:
            col_cfg[c] = st.column_config.TextColumn(c, width="small")

    row_height = 35
    table_h    = min(620, 45 + len(filtered) * row_height)

    st.dataframe(
        filtered[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config=col_cfg,
        height=table_h,
    )


# ── Downloads ─────────────────────────────────────────────────────────────────
st.markdown("---")
dl1, dl2, _ = st.columns([1.2, 1.2, 4])

with dl1:
    xls = excel_bytes(results)
    if xls:
        st.download_button(
            label="Download Full Excel",
            data=xls,
            file_name=f"DeficitResults_{run_date.strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.button("Download Full Excel", disabled=True, use_container_width=True)

with dl2:
    if not filtered.empty:
        # display_cols was built by the table block above (same filtering logic)
        csv_buf = io.StringIO()
        filtered[display_cols].to_csv(csv_buf, index=False)
        st.download_button(
            label="Download Filtered CSV",
            data=csv_buf.getvalue(),
            file_name=f"Deficits_filtered_{run_date.strftime('%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
