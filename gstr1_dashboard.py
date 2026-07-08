"""
GSTR-1 Dashboard for Modicare - Asc Global Ai
Processes monthly source data and generates GSTR-1 return sections.
"""

# Fix Windows asyncio WinError 10054 — use SelectorEventLoop instead of ProactorEventLoop.
# set_event_loop_policy is deprecated in 3.14+ (SelectorEventLoop becomes default there).
import asyncio, platform, sys
if platform.system() == "Windows" and (3, 8) <= sys.version_info < (3, 14):
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import glob
import tempfile
import xlsxwriter
from concurrent.futures import ThreadPoolExecutor, as_completed

# Default local data folder (same folder as this script)
DEFAULT_FOLDER = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _get_page_icon():
    try:
        import urllib.request as _u
        from PIL import Image as _pil
        import io as _bio
        with _u.urlopen("https://i.ibb.co/LXmWddkt/ASC-New-Logo-Blue.png", timeout=5) as _r:
            return _pil.open(_bio.BytesIO(_r.read())).convert("RGBA")
    except Exception:
        return "📊"

st.set_page_config(
    page_title="ASC GSTR-1 Dashboard",
    page_icon=_get_page_icon(),
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────
# SAAS AUTH GATE — must run before any visible UI
# ─────────────────────────────────────────────────────────────────
from saas_auth import (
    saas_auth_gate, saas_sidebar_nav, show_admin_panel,
    show_profile_page, show_announcements_banner, log_download, get_current_user,
)
saas_auth_gate()   # shows login page + st.stop() if not authenticated

st.markdown("""
<style>
    /* Hide Streamlit chrome */
    header[data-testid="stHeader"],
    div[data-testid="stToolbar"],
    div[data-testid="stDecoration"],
    div[data-testid="stStatusWidget"],
    #MainMenu, footer { display: none !important; }

    /* Zero the CSS variable that drives padding-top on the app container */
    :root { --header-height: 0px !important; }

    /* Remove top padding/margin from every container level */
    .stApp { margin-top: 0 !important; padding-top: 0 !important; }
    section[data-testid="stAppViewContainer"],
    section[data-testid="stAppViewContainer"] > div,
    div[data-testid="stAppViewBlockContainer"],
    div[data-testid="stMainBlockContainer"],
    section.main,
    section.main > div { padding-top: 0 !important; margin-top: 0 !important; }
    div.block-container { padding-top: 0.75rem !important; }

    .main-header {
        background: linear-gradient(90deg, #1e3a5f 0%, #2d6a9f 100%);
        padding: 20px 30px; border-radius: 10px; margin-bottom: 20px;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 26px; }
    .main-header p  { margin: 4px 0 0; font-size: 13px; opacity: 0.85; }
    div[data-testid="stMetric"] { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px; }

    /* ── Sidebar: ASC Blue theme ───────────────────────────────── */
    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] > div:first-child {
        background-color: #1f2f60 !important;
    }

    /* ALL text elements → white */
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div,
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] h4,
    section[data-testid="stSidebar"] h5,
    section[data-testid="stSidebar"] h6,
    section[data-testid="stSidebar"] b,
    section[data-testid="stSidebar"] strong,
    section[data-testid="stSidebar"] li,
    section[data-testid="stSidebar"] a,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] small {
        color: rgba(255,255,255,0.92) !important;
    }

    /* Slightly muted: captions, help text, secondary info */
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] span,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] .st-emotion-cache-fis6aj { color: rgba(255,255,255,0.55) !important; }

    /* Divider */
    section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.15) !important; }

    /* Inputs (return period, folder path) — white box, dark readable text */
    section[data-testid="stSidebar"] input {
        background: #ffffff !important;
        color: #1f2f60 !important;
        border-color: rgba(255,255,255,0.5) !important;
        border-radius: 6px !important;
        caret-color: #1f2f60 !important;
    }
    section[data-testid="stSidebar"] input::placeholder { color: #94adc8 !important; }

    /* BaseWeb input border wrapper */
    section[data-testid="stSidebar"] div[data-baseweb="input"] {
        background: #ffffff !important;
        border-color: rgba(255,255,255,0.4) !important;
        border-radius: 6px !important;
    }

    /* ALL sidebar buttons — base style */
    section[data-testid="stSidebar"] button {
        border-radius: 7px !important;
    }

    /* Logout / secondary button — target by kind attribute on the button itself */
    section[data-testid="stSidebar"] button[kind="secondary"],
    section[data-testid="stSidebar"] button[kind="secondaryFormSubmit"] {
        background-color: rgba(255,255,255,0.15) !important;
        border: 1.5px solid rgba(255,255,255,0.45) !important;
        color: #ffffff !important;
    }
    section[data-testid="stSidebar"] button[kind="secondary"] *,
    section[data-testid="stSidebar"] button[kind="secondaryFormSubmit"] * {
        color: #ffffff !important;
    }
    section[data-testid="stSidebar"] button[kind="secondary"]:hover,
    section[data-testid="stSidebar"] button[kind="secondaryFormSubmit"]:hover {
        background-color: rgba(255,255,255,0.28) !important;
    }

    /* Process & Generate / primary buttons — yellow */
    section[data-testid="stSidebar"] button[kind="primary"],
    section[data-testid="stSidebar"] button[kind="primaryFormSubmit"] {
        background-color: #f9be3e !important;
        color: #1f2f60 !important;
        border: none !important;
        font-weight: 700 !important;
    }
    section[data-testid="stSidebar"] button[kind="primary"] *,
    section[data-testid="stSidebar"] button[kind="primaryFormSubmit"] * {
        color: #1f2f60 !important;
    }
    section[data-testid="stSidebar"] button[kind="primary"]:hover,
    section[data-testid="stSidebar"] button[kind="primaryFormSubmit"]:hover {
        background-color: #e0a830 !important;
    }

    /* Radio button circles */
    section[data-testid="stSidebar"] [data-testid="stRadio"] label { color: rgba(255,255,255,0.9) !important; }

    /* Select/dropdown — general */
    section[data-testid="stSidebar"] div[data-baseweb="select"] { background: rgba(255,255,255,0.10) !important; }
    section[data-testid="stSidebar"] div[data-baseweb="select"] * { color: #ffffff !important; }

    /* Return period month/year selectors — white box with dark text */
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div:first-child {
        background-color: #ffffff !important;
        border: 1.5px solid rgba(255,255,255,0.5) !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div:first-child * {
        color: #1f2f60 !important;
    }
    /* Dropdown menu (the options list) */
    [data-baseweb="popover"] [data-baseweb="menu"] {
        background: #ffffff !important;
    }
    [data-baseweb="popover"] [role="option"] {
        color: #1f2f60 !important;
        background: #ffffff !important;
    }
    [data-baseweb="popover"] [role="option"]:hover,
    [data-baseweb="popover"] [aria-selected="true"] {
        background: #e8ecf7 !important;
        color: #1f2f60 !important;
    }

    /* Hide sidebar scrollbar (Chrome/Safari/Firefox/Edge) */
    section[data-testid="stSidebar"] > div:first-child {
        overflow-y: auto !important;
        scrollbar-width: none !important;
        -ms-overflow-style: none !important;
    }
    section[data-testid="stSidebar"] > div:first-child::-webkit-scrollbar {
        width: 0px !important;
        display: none !important;
    }

    /* ── File uploader: reset so content is visible ────────────── */
    /* The drop zone has a white/light background — reset text to dark */
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        background-color: rgba(255,255,255,0.95) !important;
        border: 2px dashed rgba(31,47,96,0.35) !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] *,
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] * {
        color: #1f2f60 !important;
        fill: #1f2f60 !important;
    }
    /* "Browse files" button inside uploader */
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button,
    section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
        background-color: #1f2f60 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button *,
    section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] * {
        color: #ffffff !important;
    }
    /* Keep button blue on hover — prevent Streamlit default white override */
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover,
    section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover {
        background-color: #2a3d72 !important;
        color: #ffffff !important;
        border: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover *,
    section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover * {
        color: #ffffff !important;
    }
    /* Keep dropzone background white on hover/drag-over */
    section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:hover {
        background-color: rgba(255,255,255,0.95) !important;
        border-color: #1f2f60 !important;
    }

    /* ── Static sidebar — always visible, no collapse ─────────── */
    /* Lock the sidebar open: override any transform Streamlit applies */
    section[data-testid="stSidebar"] {
        transform: none !important;
        min-width: 280px !important;
        width: 280px !important;
    }
    /* Hide both the « collapse and » expand toggle buttons */
    [data-testid="stSidebarNavButton"],
    [data-testid="stSidebarNavCollapseButton"],
    [data-testid="collapsedControl"],
    [data-testid="stSidebarNavButton"] *,
    [data-testid="stSidebarNavCollapseButton"] * {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
    }
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>GSTR-1 Dashboard</h1>
    <p>ASC Global AI &nbsp;|&nbsp; Upload monthly source files → Generate GSTR-1 return data</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def read_csv_safe(file_obj, **kwargs):
    # Try utf-8 first (fastest path for most files).
    # Fall back to latin1 which decodes every byte without raising UnicodeDecodeError.
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    try:
        return pd.read_csv(file_obj, encoding="utf-8", **kwargs)
    except UnicodeDecodeError:
        if hasattr(file_obj, "seek"):
            file_obj.seek(0)
        return pd.read_csv(file_obj, encoding="latin1", **kwargs)

def fmt_inr(val):
    try: return f"₹{val:,.2f}"
    except: return val

def filter_sales_returns(return_df):
    """
    Simplified: ONLY Taken rule applied.
    Lapsed and Not Taken checks are SKIPPED.
    ALL returns → cgst_returns → net against B2CS (Taken rules).
    Invoice Series uses retufm_no (return bill number).
    """
    if return_df is None or return_df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df = return_df.copy()

    # ── Auto-detect header row (file has headers in row 2, index 2) ──
    _known = {"fstate","inv_tot","cgstamt","igstamt","taxslab","retufm_no"}
    cols_lower = {str(c).strip().lower() for c in df.columns}
    if not (cols_lower & _known):
        for i in range(min(5, len(df))):
            row_vals = {str(v).strip().lower() for v in df.iloc[i].dropna().tolist()}
            if row_vals & _known:
                df.columns = [str(v).strip() for v in df.iloc[i].tolist()]
                df = df.iloc[i+1:].reset_index(drop=True)
                break

    df = df.dropna(how="all")

    # ── Normalise column names ────────────────────────────────────
    df.rename(columns={
        "bill_stcode"       : "bill_stcod",
        "ship_stcode"       : "ship_stcod",
        "Tax Slab"          : "taxslab",
        "tax_slab"          : "taxslab",
        "Taxable Value"     : "inv_tot",
        "IGST Amt"          : "igstamt",
        "CGST Amt"          : "cgstamt",
        "SGST Amt"          : "sgstamt",
    }, inplace=True)

    # ── Normalise Return FM No column (various naming conventions) ─
    if "retufm_no" not in df.columns:
        for _c in ["retfm_no","Return_FM_No","Return FM No","ReturnFMNo","return_fm_no","retufmno"]:
            if _c in df.columns:
                df.rename(columns={_c: "retufm_no"}, inplace=True)
                break

    # ── Normalise Return Bill Number column ────────────────────────
    if "Return_Bill_Number" not in df.columns:
        for _c in ["Return_Bill_No","Return Bill Number","Return Bill No",
                   "ReturnBillNo","return_bill_number","Return_BillNo","ReturnBillNumber"]:
            if _c in df.columns:
                df.rename(columns={_c: "Return_Bill_Number"}, inplace=True)
                break

    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df = num(df, TAX_COLS)

    # Ensure required columns exist
    for col, default in [("taxslab",0),("cgstamt",0),("igstamt",0),("fstcode",0)]:
        if col not in df.columns:
            df[col] = default

    # Clean retufm_no and Return_Bill_Number — replace NaN-like values with ""
    _nan_strs = ["nan","NaN","None","none","<NA>"]
    for _sc in ["retufm_no", "Return_Bill_Number"]:
        if _sc in df.columns:
            df[_sc] = df[_sc].fillna("").astype(str).str.strip()
            df[_sc] = df[_sc].replace(_nan_strs, "")

    # ── Filter: Only ASC Remarks = "Taken" rows ──────────────────
    _asc_vals = {"taken", "not taken", "lapsed"}
    asc_col = None
    for c in df.columns:
        c_norm = str(c).strip().lower().replace(" ", "").replace("_", "")
        if "ascremark" in c_norm or ("asc" in c_norm and "rem" in c_norm) or "remark" in c_norm:
            asc_col = c
            break
    if asc_col is None:
        for c in df.columns:
            try:
                sample = set(df[c].astype(str).str.strip().str.lower().head(1000).unique())
                if sample & _asc_vals:
                    asc_col = c
                    break
            except Exception:
                continue

    taken_df = df[df[asc_col].astype(str).str.strip().str.lower() == "taken"].copy() \
               if asc_col is not None else df.copy()

    igst_returns = pd.DataFrame()
    cgst_returns = taken_df   # Only Taken rows → 187,711 rows not 189,509
    lapsed_df    = pd.DataFrame()

    # Debug: warn if key return columns are missing so user knows to check column names
    _missing = [c for c in ["retufm_no","Return_Bill_Number"] if c not in cgst_returns.columns]
    if _missing:
        st.warning(f"⚠️ Sales Return: columns not found → {_missing}. Available: {list(cgst_returns.columns)[:15]}")

    return igst_returns, cgst_returns, lapsed_df


def build_sales_return_summary(taken_df):
    """Direct: cgst_returns already has only Taken rows. Group by state+slab, sum exactly as in sheet, negate."""
    if taken_df is None or (isinstance(taken_df, pd.DataFrame) and taken_df.empty):
        return pd.DataFrame()
    df = taken_df.copy()

    # Ensure numeric columns
    for col in ["inv_tot","igstamt","cgstamt","sgstamt","ugstamt","cessamt","taxslab"]:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["taxslab"] = df["taxslab"].round(0).astype(int)
    df = df[df["taxslab"].isin({0, 3, 5, 12, 18})]
    if df.empty:
        return pd.DataFrame()

    # Use fstate and fstcode exactly as in the sheet
    if "fstate"  not in df.columns: df["fstate"]  = ""
    if "fstcode" not in df.columns: df["fstcode"] = 0
    df["fstate"]  = df["fstate"].astype(str).str.strip()
    df["fstcode"] = pd.to_numeric(df["fstcode"], errors="coerce").fillna(0).astype(int)

    # Group by state + taxslab — sum exactly as sheet, no recalculation
    grp = df.groupby(["fstate","fstcode","taxslab"], as_index=False).agg(
        Taxable_Value=("inv_tot",  "sum"),
        IGST         =("igstamt", "sum"),
        CGST         =("cgstamt", "sum"),
        SGST         =("sgstamt", "sum"),
        UGST         =("ugstamt", "sum"),
        Cess         =("cessamt", "sum"),
    )

    # Negate — returns reduce liability
    for col in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess"]:
        grp[col] = -grp[col].abs()

    grp["Total_Tax"]     = grp["IGST"] + grp["CGST"] + grp["SGST"] + grp["UGST"] + grp["Cess"]
    grp["Invoice_Value"] = grp["Taxable_Value"] + grp["Total_Tax"]

    grp.rename(columns={"fstate":"Supplier State","fstcode":"Supplier State Code","taxslab":"Tax Slab"}, inplace=True)
    grp["Place of Supply"] = grp["Supplier State"]
    grp["POS Code"]        = grp["Supplier State Code"]

    # Add missing slabs as 0 for each state
    states    = grp[["Supplier State","Supplier State Code"]].drop_duplicates()
    all_rows  = pd.DataFrame([{"Supplier State":r["Supplier State"],"Supplier State Code":r["Supplier State Code"],"Tax Slab":s}
                               for _,r in states.iterrows() for s in [0,3,5,12,18]])
    grp = pd.merge(all_rows, grp, on=["Supplier State","Supplier State Code","Tax Slab"], how="left")
    for c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess","Total_Tax","Invoice_Value"]:
        grp[c] = grp[c].fillna(0)
    grp["Place of Supply"] = grp["Supplier State"]
    grp["POS Code"]        = grp["Supplier State Code"]

    return grp.sort_values(["Supplier State","Tax Slab"]).reset_index(drop=True)


def num(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df

TAX_COLS = ["inv_tot","sgstamt","cgstamt","igstamt","ugstamt","cessamt","totval","inv_qty","taxslab","dist_rate"]

# Only load these columns from state CSVs — skips 20+ unused columns, halves memory
_SALES_USECOLS = [
    "gstin","mscname","gstin2",
    "inv_no","inv_date",
    "fstate","fstcode","ship_state","ship_stcod",
    "bill_state","bill_stcod",
    "hsncode","taxslab","inv_qty","inv_tot","totval",
    "igstamt","cgstamt","sgstamt","ugstamt","cessamt",
]

# dtype hints — avoids slow type inference, reduces object column count
_SALES_DTYPE = {
    "fstcode"   : "Int16", "ship_stcod": "Int16", "bill_stcod": "Int16",
    "inv_tot"   : "float64", "totval"   : "float64",
    "igstamt"   : "float32", "cgstamt"  : "float32",
    "sgstamt"   : "float32", "ugstamt"  : "float32", "cessamt"   : "float32",
    "taxslab"   : "float32", "inv_qty"  : "float32",
}

# GST State Code → State Name (for GSTIN-based POS derivation in B2B)
GSTIN_STATE_MAP = {
    "01":"Jammu & Kashmir",  "02":"Himachal Pradesh", "03":"Punjab",
    "04":"Chandigarh",       "05":"Uttarakhand",      "06":"Haryana",
    "07":"Delhi",            "08":"Rajasthan",        "09":"Uttar Pradesh",
    "10":"Bihar",            "11":"Sikkim",           "12":"Arunachal Pradesh",
    "13":"Nagaland",         "14":"Manipur",          "15":"Mizoram",
    "16":"Tripura",          "17":"Meghalaya",        "18":"Assam",
    "19":"West Bengal",      "20":"Jharkhand",        "21":"Odisha",
    "22":"Chhattisgarh",     "23":"Madhya Pradesh",   "24":"Gujarat",
    "25":"Daman and Diu",    "26":"Dadra & Nagar Haveli", "27":"Maharashtra",
    "29":"Karnataka",        "30":"Goa",              "31":"Lakshadweep",
    "32":"Kerala",           "33":"Tamilnadu",        "34":"Puducherry",
    "35":"Andaman & Nicobar","36":"Telangana",        "37":"Andhra Pradesh",
    "38":"Ladakh",
}

# ─────────────────────────────────────────────────────────────────
# PAGE ROUTING — admin / profile / dashboard
# ─────────────────────────────────────────────────────────────────
_saas_page = saas_sidebar_nav()   # renders user info + nav in sidebar
_cur_user  = get_current_user()

if _saas_page == "admin" and _cur_user.get("role") == "admin":
    show_admin_panel(_cur_user)
    st.stop()
elif _saas_page == "profile":
    show_profile_page(_cur_user)
    st.stop()

# Show announcements (only on dashboard)
show_announcements_banner()

# ─────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Source Files")
    import datetime as _dt
    st.markdown("**Return Period**")
    _months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    _years  = list(range(2020, 2031))
    _now    = _dt.datetime.now()
    _col_m, _col_y = st.columns(2)
    with _col_m:
        _sel_month = st.selectbox("Month", _months,
                                  index=_now.month - 1,
                                  label_visibility="collapsed")
    with _col_y:
        _sel_year = st.selectbox("Year", _years,
                                 index=_years.index(_now.year) if _now.year in _years else 0,
                                 label_visibility="collapsed")
    period_label = f"{_sel_month}-{str(_sel_year)[2:]}"
    st.caption(f"📅 {_dt.datetime.strptime(period_label, '%b-%y').strftime('%B %Y')}")
    st.markdown("---")

    load_mode = st.radio(
        "How to load files?",
        ["📂 From Local Folder (Recommended)", "⬆️ Upload Files Manually"],
        index=0
    )

    if load_mode == "📂 From Local Folder (Recommended)":
        st.markdown("**Folder path:**")
        data_folder = st.text_input(
            "Folder containing source files",
            value=DEFAULT_FOLDER,
            label_visibility="collapsed"
        )
        st.caption("Put all monthly files in this folder (or subfolders) and click Process.")
        state_files = None
        stock_file  = None
        return_file = None
        crosscharge_file = None
        assets_file = None
    else:
        data_folder = None
        st.subheader("1. State-wise Sales CSVs")
        state_files = st.file_uploader(
            "Upload all state CSV files", type=["csv"],
            accept_multiple_files=True, key="state_sales"
        )
        st.subheader("2. Stock Transfer CSV")
        stock_file = st.file_uploader("GST_StockTransfer*.csv", type=["csv"], key="stock")
        st.subheader("3. Sales Return Excel")
        return_file = st.file_uploader("GST_SalesReturn_*.xlsx", type=["xlsx","xls"], key="sret")
        st.subheader("4. Cross Charge Invoices Excel")
        crosscharge_file = st.file_uploader("Cross Charge Invoices*.xlsx", type=["xlsx","xls"], key="cc")
        st.subheader("5. Master Assets Sale Excel")
        assets_file = st.file_uploader("Master Assets Sale Details*.xlsx", type=["xlsx","xls"], key="assets")

    process_btn = st.button("🚀 Process & Generate GSTR-1", width="stretch", type="primary")

# ── helper to auto-detect files from folder ──────────────────────
def detect_local_files(folder):
    """Return paths for each file type from a local folder."""
    state_csvs = (
        glob.glob(os.path.join(folder, "**/saledtl*.csv"), recursive=True) +
        glob.glob(os.path.join(folder, "**/_saledtl*.csv"), recursive=True)
    )
    # dedupe
    state_csvs = list(dict.fromkeys(state_csvs))

    def first_match(*patterns):
        for pat in patterns:
            hits = glob.glob(os.path.join(folder, "**", pat), recursive=True)
            if hits: return hits[0]
        return None

    return {
        "state_csvs"  : state_csvs,
        "stock"       : first_match("GST_StockTransfer*.csv","*StockTransfer*.csv"),
        "sales_return": first_match("GST_Sales_Return*.xlsx","GST_SalesReturn*.xlsx","*Sales_Return*.xlsx","*SalesReturn*.xlsx"),
        "cross_charge": first_match("*Cross*Charge*.xlsx","*CrossCharge*.xlsx"),
        "assets"      : first_match("*Assets*Sale*.xlsx","*Asset*Sale*.xlsx"),
    }

# Status badges
if load_mode == "📂 From Local Folder (Recommended)":
    detected = detect_local_files(data_folder) if data_folder and os.path.isdir(data_folder) else {}
    c1,c2,c3,c4,c5 = st.columns(5)
    for col, lbl, val in zip(
        [c1,c2,c3,c4,c5],
        ["State Sales","Stock Transfer","Sales Return","Cross Charge","Assets"],
        [
            f"{len(detected.get('state_csvs',[]))} files" if detected.get('state_csvs') else None,
            os.path.basename(detected.get('stock','') or '') or None,
            os.path.basename(detected.get('sales_return','') or '') or None,
            os.path.basename(detected.get('cross_charge','') or '') or None,
            os.path.basename(detected.get('assets','') or '') or None,
        ]
    ):
        with col:
            st.markdown(f"**{'✅' if val else '⬜'} {lbl}**")
            st.caption(val if val else "Not found")
else:
    detected = {}
    c1,c2,c3,c4,c5 = st.columns(5)
    for col, lbl, up, cnt in zip(
        [c1,c2,c3,c4,c5],
        ["State Sales","Stock Transfer","Sales Return","Cross Charge","Assets"],
        [bool(state_files), bool(stock_file), bool(return_file), bool(crosscharge_file), bool(assets_file)],
        [len(state_files) if state_files else 0, 1 if stock_file else 0,
         1 if return_file else 0, 1 if crosscharge_file else 0, 1 if assets_file else 0],
    ):
        with col:
            st.markdown(f"**{'✅' if up else '⬜'} {lbl}**")
            st.caption(f"{cnt} file(s)" if up else "Not uploaded")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────
# CLASSIFY SALES
# ─────────────────────────────────────────────────────────────────
def classify_sales(df):
    # ── Copy ONLY the columns needed — avoids OOM on large DataFrames ──
    _work_cols = [c for c in
        ["gstin2","fstcode","ship_stcod","ship_state","bill_stcod","bill_state","inv_no","inv_tot"]
        if c in df.columns]
    w = df[_work_cols].copy()   # small copy: ~8 cols × N rows instead of all 43+ cols

    w["gstin2"] = w["gstin2"].astype(str).str.strip().replace({"nan":"","NaN":"","None":""})
    for col in ["fstcode","ship_stcod","inv_tot"]:
        if col in w.columns:
            w[col] = pd.to_numeric(w[col], errors="coerce").fillna(0)
    w["fstcode"]    = w["fstcode"].astype(int)
    w["ship_stcod"] = w["ship_stcod"].astype(int)

    # POS correction: use bill_stcod as Place of Supply where valid
    if "bill_stcod" in w.columns and "bill_state" in w.columns:
        w["bill_stcod"] = pd.to_numeric(w["bill_stcod"], errors="coerce").fillna(0).astype(int)
        valid_bill = w["bill_stcod"] > 0
        w.loc[valid_bill, "ship_stcod"] = w.loc[valid_bill, "bill_stcod"]
        w.loc[valid_bill, "ship_state"]  = w.loc[valid_bill, "bill_state"]

    has_gstin   = w["gstin2"].str.len() == 15
    inter_state = w["fstcode"] != w["ship_stcod"]
    inv_taxable = w.groupby("inv_no")["inv_tot"].transform("sum")
    large_inv   = inv_taxable >= 100000

    gstr1_section = np.where(
        has_gstin, "B2B",
        np.where(inter_state & large_inv, "B2CL", "B2CS")
    )

    # Direct column assignment — most reliable approach under memory pressure.
    # df is a fresh DataFrame from pd.concat so direct assignment is safe.
    with pd.option_context("mode.copy_on_write", False):
        df["gstin2"]        = w["gstin2"].values
        df["fstcode"]       = w["fstcode"].values
        df["ship_stcod"]    = w["ship_stcod"].values
        if "ship_state" in w.columns:
            df["ship_state"] = w["ship_state"].values
        df["gstr1_section"] = gstr1_section
    return df

# ─────────────────────────────────────────────────────────────────
# B2B  – Table 4A
# ─────────────────────────────────────────────────────────────────
def build_b2b(sales_df):
    df = sales_df[sales_df["gstr1_section"] == "B2B"].copy()
    if df.empty: return pd.DataFrame()
    df = num(df, TAX_COLS)

    # For B2B: Place of Supply = buyer's GSTIN state (first 2 digits of gstin2).
    if "gstin2" in df.columns:
        gstin_prefix = df["gstin2"].astype(str).str[:2]
        df["ship_stcod"] = pd.to_numeric(gstin_prefix, errors="coerce").fillna(df["ship_stcod"]).astype(int)
        df["ship_state"]  = gstin_prefix.map(GSTIN_STATE_MAP).fillna(df["ship_state"])

    if "totval" not in df.columns:
        df["totval"] = 0
    line = df.groupby(
        ["gstin","mscname","inv_no","inv_date","fstate","fstcode","gstin2","ship_state","ship_stcod","taxslab"],
        as_index=False
    ).agg(
        Taxable_Value=("inv_tot","sum"),
        IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
        SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
        Total_Invoice_Value=("totval","first"),
    )
    line["Invoice_Value"] = line["Taxable_Value"] + line["IGST"] + line["CGST"] + line["SGST"] + line["UGST"] + line["Cess"]

    line.rename(columns={
        "gstin":"Supplier GSTIN","mscname":"Supplier Name",
        "inv_no":"Invoice No","inv_date":"Invoice Date",
        "fstate":"Supplier State","fstcode":"Supplier State Code",
        "gstin2":"Receiver GSTIN","ship_state":"Place of Supply","ship_stcod":"POS Code",
        "taxslab":"Tax Slab",
        "Total_Invoice_Value":"Total Value",
    }, inplace=True)
    line["retufm_no"]          = ""
    line["Return_Bill_Number"] = ""
    return line

# ─────────────────────────────────────────────────────────────────
# B2CL – Table 5A
# ─────────────────────────────────────────────────────────────────
def build_b2cl(sales_df):
    df = sales_df[sales_df["gstr1_section"] == "B2CL"].copy()
    if df.empty: return pd.DataFrame()
    df = num(df, TAX_COLS)

    if "totval" not in df.columns:
        df["totval"] = 0
    line = df.groupby(
        ["inv_no","inv_date","mscname","gstin","fstate","fstcode","ship_state","ship_stcod","taxslab"],
        as_index=False
    ).agg(
        Taxable_Value=("inv_tot","sum"),
        IGST=("igstamt","sum"), Cess=("cessamt","sum"),
        Total_Invoice_Value=("totval","first"),
    )
    line["Invoice_Value"] = line["Taxable_Value"] + line["IGST"] + line["Cess"]
    line.rename(columns={
        "inv_no":"Invoice No","inv_date":"Invoice Date",
        "mscname":"Supplier Name","gstin":"Supplier GSTIN",
        "fstate":"Supplier State","fstcode":"Supplier State Code",
        "ship_state":"Place of Supply","ship_stcod":"POS Code",
        "taxslab":"Tax Slab",
        "Total_Invoice_Value":"Total Value",
    }, inplace=True)
    line["retufm_no"]          = ""
    line["Return_Bill_Number"] = ""
    return line

# ─────────────────────────────────────────────────────────────────
# B2CS – Table 7
# ─────────────────────────────────────────────────────────────────
def build_b2cs(sales_df=None, cgst_returns=None):
    """
    B2CS with Taken rules:
      Sales: group by (Supplier State + Place of Supply + Tax Slab)
             IGST for inter-state, CGST+SGST for intra-state
      Returns (Taken): CGST invoices — all tax slabs
             Rule: Fstate = Bill State (treat as intra-state)
             → net against intra-state B2CS rows per (Supplier State + Tax Slab)
             → if no matching B2CS sale → add negative row
    """
    # ── 1. B2CS SALES aggregated by state + POS + taxslab ─────────
    sales_grp = pd.DataFrame()
    if sales_df is not None and not sales_df.empty:
        df = sales_df[sales_df["gstr1_section"] == "B2CS"].copy()
        if not df.empty:
            df = num(df, TAX_COLS)
            # Convert float32 → float64 for exact decimal values (same as Combined GSTR-1)
            for _fc in ["inv_tot","igstamt","cgstamt","sgstamt","ugstamt","cessamt"]:
                if _fc in df.columns:
                    df[_fc] = df[_fc].astype("float64")
            # Union Territory state codes — these use CGST+UGST instead of CGST+SGST
            _UT_CODES = {1, 4, 25, 26, 31, 34, 35, 38}
            df["_tt"]     = df["igstamt"] + df["cgstamt"] + df["sgstamt"] + df["ugstamt"]
            is_inter      = df["fstcode"].astype(int) != df["ship_stcod"].astype(int)
            is_ut         = df["fstcode"].astype(int).isin(_UT_CODES)
            df["igstamt"] = np.where(is_inter,                   df["_tt"],   0)
            df["cgstamt"] = np.where(~is_inter,                  df["_tt"]/2, 0)
            df["sgstamt"] = np.where(~is_inter & ~is_ut,         df["_tt"]/2, 0)
            df["ugstamt"] = np.where(~is_inter &  is_ut,         df["_tt"]/2, 0)
            df.drop(columns=["_tt"], inplace=True)
            _gcols = [c for c in ["fstate","fstcode","ship_state","ship_stcod","taxslab"]
                      if c in df.columns]
            sales_grp = df.groupby(_gcols, as_index=False).agg(
                Taxable_Value=("inv_tot", lambda x: int(np.round(x * 100).sum()) / 100),
                IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
                SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
            )

    # ── 2. RETURNS (Taken) — all tax slabs ────────────────────────
    # Rule: fstate = bill_state → treat as intra-state (CGST+SGST)
    if cgst_returns is not None and not cgst_returns.empty:
        ret = num(cgst_returns.copy(), TAX_COLS)
        # Normalise taxslab column
        if "taxslab" not in ret.columns:
            for alt in ["tax_slab","Tax Slab","TAXSLAB","rate","gst_rate"]:
                if alt in ret.columns:
                    ret["taxslab"] = pd.to_numeric(ret[alt], errors="coerce").fillna(0)
                    break
            else:
                ret["taxslab"] = 0
        ret["taxslab"]  = pd.to_numeric(ret["taxslab"], errors="coerce").fillna(0)
        if "fstcode" in ret.columns:
            ret["fstcode"] = pd.to_numeric(ret["fstcode"], errors="coerce").fillna(0).astype(int)
        else:
            ret["fstcode"] = 0

        # All tax slabs processed — no slab filter
        if not ret.empty:
            # Ensure totval exists (Invoice Value = Taxable + All Taxes)
            if "totval" not in ret.columns:
                ret["totval"] = (pd.to_numeric(ret.get("inv_tot",0), errors="coerce").fillna(0) +
                                 pd.to_numeric(ret.get("cgstamt",0), errors="coerce").fillna(0) +
                                 pd.to_numeric(ret.get("sgstamt",0), errors="coerce").fillna(0) +
                                 pd.to_numeric(ret.get("igstamt",0), errors="coerce").fillna(0))
            else:
                ret["totval"] = pd.to_numeric(ret["totval"], errors="coerce").fillna(0)

            # Group returns by (supplier state + taxslab)
            ret_grp = ret.groupby(["fstate","fstcode","taxslab"], as_index=False).agg(
                R_Tax   =("inv_tot","sum"),
                R_CGST  =("cgstamt","sum"),
                R_SGST  =("sgstamt","sum"),
                R_IGST  =("igstamt","sum"),
                R_Cess  =("cessamt","sum"),
                R_InvVal=("totval","sum"),   # ← SUM OF INVOICE VALUE
            )

            new_rows = []
            for _, r in ret_grp.iterrows():
                fstate  = str(r["fstate"]).strip()
                fcode   = int(r["fstcode"])
                tslab   = r["taxslab"]
                r_tax   = abs(r["R_Tax"])
                r_cgst  = abs(r["R_CGST"])
                r_sgst  = abs(r["R_SGST"])
                r_igst  = abs(r["R_IGST"])
                r_cess  = abs(r["R_Cess"])
                r_invval= abs(r["R_InvVal"])

                # Always add a SEPARATE NEGATIVE ROW for the return amount.
                # Rule: fstate = bill_state (intra-state treatment) → all tax as CGST+SGST, IGST=0
                sup_state  = GSTIN_STATE_MAP.get(str(fcode).zfill(2), fstate)
                total_tax  = r_igst + r_cgst + r_sgst + r_cess
                half_tax   = round(total_tax / 2, 0)
                new_rows.append({
                    "fstate"       : fstate,
                    "fstcode"      : fcode,
                    "ship_state"   : sup_state,
                    "ship_stcod"   : fcode,
                    "taxslab"      : tslab,
                    "Taxable_Value": round(-r_tax, 0),
                    "IGST"         : 0,
                    "CGST"         : -half_tax,
                    "SGST"         : -half_tax,
                    "UGST"         : 0,
                    "Cess"         : round(-r_cess, 0),
                    "Invoice_Value": round(-r_invval, 0),
                    "Source"       : "Return",
                })
            if new_rows:
                sales_grp = pd.concat([sales_grp, pd.DataFrame(new_rows)], ignore_index=True)

    if sales_grp.empty:
        return pd.DataFrame()

    # Mark sales rows if not already tagged
    if "Source" not in sales_grp.columns:
        sales_grp["Source"] = "Sales"
    else:
        sales_grp["Source"] = sales_grp["Source"].fillna("Sales")

    # ── 3. Final aggregation — include Source to keep Return rows SEPARATE ──
    # Without Source in groupby, negative return rows merge with positive sales rows
    _final_gcols = [c for c in ["fstate","fstcode","ship_state","ship_stcod","taxslab","Source"]
                    if c in sales_grp.columns]
    grp = sales_grp.groupby(_final_gcols, as_index=False).agg(
        Taxable_Value=("Taxable_Value","sum"),
        IGST=("IGST","sum"), CGST=("CGST","sum"),
        SGST=("SGST","sum"), UGST=("UGST","sum"), Cess=("Cess","sum"),
    )
    grp["Total_Tax"]     = grp["IGST"]+grp["CGST"]+grp["SGST"]+grp["UGST"]+grp["Cess"]
    grp["Invoice_Value"] = grp["Taxable_Value"] + grp["Total_Tax"]

    # Keep exact decimal values — no rounding applied
    for _rc in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess","Total_Tax","Invoice_Value"]:
        if _rc in grp.columns:
            grp[_rc] = pd.to_numeric(grp[_rc], errors="coerce").fillna(0)

    grp.rename(columns={
        "fstate":"Supplier State","fstcode":"Supplier State Code",
        "ship_state":"Place of Supply","ship_stcod":"POS Code","taxslab":"Tax Slab",
    }, inplace=True)
    if "Place of Supply" not in grp.columns:
        grp["Place of Supply"] = grp.get("Supplier State","")
    if "POS Code" not in grp.columns:
        grp["POS Code"] = grp.get("Supplier State Code","")

    col_order = ["Source","Supplier State","Supplier State Code","Place of Supply","POS Code",
                 "Tax Slab","Taxable_Value","IGST","CGST","SGST","UGST","Cess",
                 "Total_Tax","Invoice_Value"]
    return grp[[c for c in col_order if c in grp.columns]]

# ─────────────────────────────────────────────────────────────────
# CDNR – Table 9B (Credit Notes to Registered)
# ─────────────────────────────────────────────────────────────────
def build_cdnr(return_df):
    if return_df is None or return_df.empty: return pd.DataFrame()
    df = return_df.copy()
    df["gstin2"] = df["gstin2"].astype(str).str.strip().replace({"nan":"","NaN":"","None":""})
    df = df[df["gstin2"].str.len() == 15].copy()
    if df.empty: return pd.DataFrame()
    df = num(df, TAX_COLS)

    # Normalise column name differences between return file and sales file
    ship_col = "ship_stcode" if "ship_stcode" in df.columns else "ship_stcod"

    key_cols = [c for c in ["gstin","mscname","gstin2","retufm_no","retufm_dt","RETURNTYPE","fstate","fstcode","taxslab",ship_col] if c in df.columns]
    grp = df.groupby(key_cols, as_index=False).agg(
        Taxable_Value=("inv_tot","sum"),
        IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
        SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
    )
    # Negate monetary values — returns reduce tax liability
    for _c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess"]:
        grp[_c] = -grp[_c].abs()
    grp["Note_Value"] = grp["Taxable_Value"]+grp["IGST"]+grp["CGST"]+grp["SGST"]+grp["UGST"]+grp["Cess"]
    grp.rename(columns={
        "gstin":"Supplier GSTIN","mscname":"Supplier Name",
        "gstin2":"Receiver GSTIN","retufm_no":"Note No","retufm_dt":"Note Date",
        "RETURNTYPE":"Note Type","fstate":"Supplier State","fstcode":"Supplier State Code",
        "taxslab":"Tax Slab", ship_col:"Place of Supply",
    }, inplace=True)
    # Invoice No = Return Bill No — reorder to place it before Note No
    grp["Invoice No"] = grp["Note No"] if "Note No" in grp.columns else ""
    cols = list(grp.columns)
    if "Invoice No" in cols and "Note No" in cols:
        cols.remove("Invoice No")
        cols.insert(cols.index("Note No"), "Invoice No")
        grp = grp[cols]
    return grp

# ─────────────────────────────────────────────────────────────────
# CDNUR – Table 9C (Credit Notes to Unregistered)
# ─────────────────────────────────────────────────────────────────
def build_cdnur(return_df):
    if return_df is None or return_df.empty: return pd.DataFrame()
    df = return_df.copy()
    df["gstin2"] = df["gstin2"].astype(str).str.strip().replace({"nan":"","NaN":"","None":""})
    df = df[df["gstin2"].str.len() != 15].copy()
    if df.empty: return pd.DataFrame()
    df = num(df, TAX_COLS)

    ship_col = "ship_stcode" if "ship_stcode" in df.columns else "ship_stcod"
    key_cols = [c for c in ["gstin","mscname","retufm_no","retufm_dt","RETURNTYPE","fstate","fstcode","taxslab",ship_col] if c in df.columns]
    grp = df.groupby(key_cols, as_index=False).agg(
        Taxable_Value=("inv_tot","sum"),
        IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
        SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
    )
    # Negate monetary values — returns reduce tax liability
    for _c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess"]:
        grp[_c] = -grp[_c].abs()
    grp["Note_Value"] = grp["Taxable_Value"]+grp["IGST"]+grp["CGST"]+grp["SGST"]+grp["UGST"]+grp["Cess"]
    grp.rename(columns={
        "gstin":"Supplier GSTIN","mscname":"Supplier Name",
        "retufm_no":"Note No","retufm_dt":"Note Date","RETURNTYPE":"Note Type",
        "fstate":"Supplier State","fstcode":"Supplier State Code",
        "taxslab":"Tax Slab", ship_col:"Place of Supply",
    }, inplace=True)
    # Invoice No = Return Bill No — reorder to place it before Note No
    grp["Invoice No"] = grp["Note No"] if "Note No" in grp.columns else ""
    cols = list(grp.columns)
    if "Invoice No" in cols and "Note No" in cols:
        cols.remove("Invoice No")
        cols.insert(cols.index("Note No"), "Invoice No")
        grp = grp[cols]
    return grp

# ─────────────────────────────────────────────────────────────────
# STOCK TRANSFER – Table 6A
# ─────────────────────────────────────────────────────────────────
def build_stock_transfer(stock_df):
    if stock_df is None or stock_df.empty: return pd.DataFrame()
    df = stock_df.copy()
    df = num(df, TAX_COLS + ["freight"])

    if "totval" not in df.columns:
        df["totval"] = 0
    grp = df.groupby(
        [c for c in ["gstin","mscname","inv_no","inv_date","fstate","fstcode","gstin2","ship_state","ship_stcod","hsncode","taxslab"] if c in df.columns],
        as_index=False
    ).agg(
        Total_Qty=("inv_qty","sum"),
        Taxable_Value=("inv_tot","sum"),
        IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
        SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
        Total_Invoice_Value=("totval","first"),
    )
    grp["Invoice_Value"] = grp["Taxable_Value"]+grp["IGST"]+grp["CGST"]+grp["SGST"]+grp["UGST"]+grp["Cess"]
    grp.rename(columns={
        "gstin":"Consignor GSTIN","mscname":"From Warehouse",
        "inv_no":"Transfer Invoice No","inv_date":"Invoice Date",
        "fstate":"From State","fstcode":"From State Code",
        "gstin2":"Consignee GSTIN","ship_state":"To State","ship_stcod":"To State Code",
        "hsncode":"HSN Code","taxslab":"Tax Slab",
        "Total_Invoice_Value":"Total Value",
    }, inplace=True)
    grp["retufm_no"]          = ""
    grp["Return_Bill_Number"] = ""
    return grp

# ─────────────────────────────────────────────────────────────────
# CROSS CHARGE
# ─────────────────────────────────────────────────────────────────
def build_cross_charge(raw_df):
    if raw_df is None or raw_df.empty: return pd.DataFrame()
    df = raw_df.copy()

    # ── Extract From State from title row 0 ──────────────────────
    # Title example: "CROSS CHARGE INVOICING FROM ASSAM GSTIN"
    title_text = " ".join(df.iloc[0].dropna().astype(str).tolist()).upper()
    from_state = from_state_code = ""
    for code, name in GSTIN_STATE_MAP.items():
        if name.upper() in title_text:
            from_state      = name
            from_state_code = int(code) if code.isdigit() else code
            break

    # Also try to find a GSTIN in the title for From GSTIN
    import re as _re
    _gstin_pat  = _re.compile(r'\b\d{2}[A-Z]{5}\d{4}[A-Z][\dA-Z]Z[\dA-Z]\b')
    from_gstin  = ""
    for val in df.iloc[0].astype(str).tolist():
        m = _gstin_pat.search(val.strip())
        if m:
            from_gstin = m.group()
            if not from_state:
                sc = from_gstin[:2]
                from_state      = GSTIN_STATE_MAP.get(sc, "")
                from_state_code = int(sc) if sc.isdigit() else sc
            break

    # Row 1 (index=1) is the actual column header
    df.columns = [str(v).strip() for v in df.iloc[1].tolist()]
    df = df.iloc[2:].reset_index(drop=True)
    df = df.dropna(how="all")

    # Rename all known column variants to standard names
    df.rename(columns={
        "From State"  : "From State",
        "From GSTIN"  : "From GSTIN",
        "To STATES"   : "To State",
        "STATES"      : "To State",
        "State Code"  : "To State Code",
        "To State Code": "To State Code",
        "To GST No."  : "Receiver GSTIN",
        "GST No."     : "Receiver GSTIN",
        "Invoice No." : "Invoice No",
        "Invoice Date": "Invoice Date",
        "Monthly"     : "Taxable Value",
        "GST @18%"    : "GST (18%)",
        "Invoice Amt.": "Invoice Value",
        "Remarks"     : "Remarks",
    }, inplace=True)

    for col in ["Taxable Value","GST (18%)","Invoice Value"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Remove total / blank rows
    state_col = "To State" if "To State" in df.columns else None
    if state_col:
        df = df[~df[state_col].astype(str).str.strip().str.upper().isin(["NAN","TOTAL",""])]

    # Apply From State from title (overrides blank column if present)
    if "From State" not in df.columns or df["From State"].astype(str).str.strip().eq("").all():
        df["From State"] = from_state
    if "From State Code" not in df.columns or df["From State Code"].astype(str).str.strip().eq("").all():
        df["From State Code"] = from_state_code
    if from_gstin and ("From GSTIN" not in df.columns or df["From GSTIN"].astype(str).str.strip().eq("").all()):
        df["From GSTIN"] = from_gstin

    # Derive To State Code from Receiver GSTIN if missing
    if "To State Code" not in df.columns and "Receiver GSTIN" in df.columns:
        df["To State Code"] = df["Receiver GSTIN"].astype(str).str[:2].apply(
            lambda x: int(x) if x.isdigit() else x
        )

    if "HSN Code" not in df.columns:
        df["HSN Code"] = 998399
    if "Tax Slab" not in df.columns:
        df["Tax Slab"] = 18

    col_order = ["HSN Code","From State","From State Code","From GSTIN",
                 "To State","To State Code","Receiver GSTIN",
                 "Address","Pin Code","Invoice No","Invoice Date",
                 "Taxable Value","GST (18%)","Invoice Value","Remarks","Tax Slab"]
    return df[[c for c in col_order if c in df.columns]]

# ─────────────────────────────────────────────────────────────────
# ASSETS SALE
# ─────────────────────────────────────────────────────────────────
def build_assets(assets_df):
    if assets_df is None or assets_df.empty: return pd.DataFrame()
    df = assets_df.copy()

    # ── Auto-detect header row ────────────────────────────────────
    _known_new  = {"fstate","fstcode","gstin2","inv_no","inv_tot","taxslab"}
    _known_old  = {"invoice no.", "taxable amount", "consignor name", "s.no."}
    _known_exp  = {"supplier state","consignee gstin","invoice no","taxable amount","hsn"}

    cols_lower = {str(c).strip().lower() for c in df.columns}
    if not (cols_lower & (_known_new | _known_old | _known_exp)):
        for i in range(min(3, len(df))):
            row_vals = {str(v).strip().lower() for v in df.iloc[i].dropna().tolist()}
            if row_vals & (_known_new | _known_old | _known_exp):
                df.columns = [str(v).strip() for v in df.iloc[i].tolist()]
                df = df.iloc[i+1:].reset_index(drop=True)
                break

    df = df.dropna(how="all")

    # ── Rename all known column variants ─────────────────────────
    df.rename(columns={
        # Raw CSV-style columns
        "fstate"      : "Supplier State",
        "fstcode"     : "Supplier State Code",
        "gstin2"      : "Consignee GSTIN",
        "bill_state"  : "Consignee State",
        "inv_no"      : "Invoice No",
        "inv_date"    : "Invoice Date",
        "taxslab"     : "Tax Slab",
        "inv_tot"     : "Taxable Amount",
        "igstamt"     : "IGST",
        "cgstamt"     : "CGST",
        "sgstamt"     : "SGST",
        "ugstamt"     : "UGST",
        "cessamt"     : "Cess",
        # Formatted Excel columns
        "Invoice No." : "Invoice No",
        "Invoice Date": "Invoice Date",
        "Invoice Value": "Invoice Value",
        # Fix misspelled Quantity column
        "Qnantity"    : "Quantity",
        "Qunatity"    : "Quantity",
        "Quantity"    : "Quantity",
        # Other common variants
        "TYPE"        : "Type",
        "Category"    : "Category",
        "HSN"         : "HSN Code",
        "UQC"         : "UQC",
    }, inplace=True)

    # ── Numeric columns ───────────────────────────────────────────
    for col in ["Taxable Amount","CGST","SGST","IGST","UGST","Cess","Invoice Value","Tax Slab"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── Round Quantity to 2 decimal places ───────────────────────
    if "Quantity" in df.columns:
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").round(2).fillna(0)

    # ── Derive Invoice Value if missing ──────────────────────────
    if "Invoice Value" not in df.columns:
        tax_cols = ["IGST","CGST","SGST","UGST","Cess"]
        tax_sum  = sum(df[c] for c in tax_cols if c in df.columns)
        df["Invoice Value"] = df.get("Taxable Amount", 0) + tax_sum

    # ── Derive Supplier State from Consignor GSTIN if missing ─────
    if "Supplier State" not in df.columns or df["Supplier State"].astype(str).str.strip().eq("").all():
        _gstin_col = next(
            (c for c in df.columns if "consignor" in c.lower() and "gst" in c.lower()),
            next((c for c in df.columns if "gstin" in c.lower() or "gst no" in c.lower()), None)
        )
        if _gstin_col:
            prefix = df[_gstin_col].astype(str).str[:2]
            df["Supplier State Code"] = prefix.apply(lambda x: int(x) if x.isdigit() else x)
            df["Supplier State"]      = prefix.map(GSTIN_STATE_MAP).fillna("")
        else:
            df["Supplier State Code"] = ""
            df["Supplier State"]      = ""
    elif "Supplier State Code" not in df.columns:
        df["Supplier State Code"] = ""

    if "Tax Slab" not in df.columns:
        df["Tax Slab"] = ""

    # ── Remove blank / total rows ─────────────────────────────────
    inv_col = next((c for c in ["Invoice No","Invoice No."] if c in df.columns), None)
    if inv_col:
        df = df[~df[inv_col].astype(str).str.strip().str.upper().isin(["NAN","","TOTAL"])]

    # ── Classify into GSTR1 sections ─────────────────────────────
    # Category = "B2B" → B2B tab | Category = "B2C" → B2CS (B2C) tab
    if "Category" in df.columns:
        is_b2b = df["Category"].astype(str).str.strip().str.upper() == "B2B"
        df["GSTR1_Section"] = np.where(is_b2b, "B2B", "B2CS")
    else:
        df["GSTR1_Section"] = "B2B"

    # ── Final column order ────────────────────────────────────────
    col_order = ["GSTR1_Section","Supplier State","Supplier State Code",
                 "Consignee GSTIN","Consignee State",
                 "Invoice No","Invoice Date","Type","Category","HSN Code","UQC","Quantity",
                 "Tax Slab","Taxable Amount","IGST","CGST","SGST","UGST","Cess","Invoice Value"]
    present = [c for c in col_order if c in df.columns]
    rest    = [c for c in df.columns if c not in present]
    result  = df[present + rest]

    # Remove duplicate columns and 'nan' named columns
    result  = result.loc[:, ~result.columns.duplicated()]
    result  = result[[c for c in result.columns if str(c).strip().lower() != "nan"]]
    return result

# ─────────────────────────────────────────────────────────────────
# HSN SUMMARY – Table 12
# ─────────────────────────────────────────────────────────────────
def build_hsn(sales_df, cc_df=None, stock_df=None, cgst_returns=None, assets_df=None):
    # HSN Summary: Sales + Cross Charge + Stock Transfer + Returns + Assets
    _hsn_cols = ["fstate","fstcode","gstr1_section","hsncode","proddesc","taxslab",
                 "inv_qty","inv_tot","igstamt","cgstamt","sgstamt","ugstamt","cessamt"]
    frames = []
    if sales_df is not None and not sales_df.empty:
        tmp = sales_df[[c for c in _hsn_cols if c in sales_df.columns]].copy()
        if "gstr1_section" in tmp.columns:
            tmp["Type"] = np.where(tmp["gstr1_section"] == "B2B", "B2B", "B2C")
        else:
            tmp["Type"] = "B2C"
        tmp["source"]     = "Sales"
        tmp["uqc"]        = "PCS-PIECES"
        # From State (Supplier) = fstate for sales rows
        tmp["from_state"]  = tmp["fstate"].astype(str).str.strip()  if "fstate"  in tmp.columns else ""
        tmp["from_stcode"] = pd.to_numeric(tmp["fstcode"], errors="coerce").fillna(0).astype(int) if "fstcode" in tmp.columns else 0
        frames.append(tmp)

    # Add Stock Transfer rows
    if stock_df is not None and not stock_df.empty:
        st_tmp = stock_df[[c for c in _hsn_cols if c in stock_df.columns]].copy()
        st_tmp["Type"]       = "Stock Transfer"
        st_tmp["source"]     = "StockTransfer"
        st_tmp["uqc"]        = "PCS-PIECES"
        # From State (Supplier) = fstate | From State Code = fstcode from stock transfer CSV
        st_tmp["from_state"]  = st_tmp["fstate"].astype(str).str.strip()  if "fstate"  in st_tmp.columns else ""
        st_tmp["from_stcode"] = pd.to_numeric(st_tmp["fstcode"], errors="coerce").fillna(0).astype(int) if "fstcode" in st_tmp.columns else 0
        frames.append(st_tmp)

    # Add Assets Sales rows — map Assets columns to HSN format
    if assets_df is not None and not assets_df.empty and "HSN Code" in assets_df.columns:
        _col_map = {
            "Supplier State"     : "fstate",
            "Supplier State Code": "fstcode",
            "HSN Code"           : "hsncode",
            "Tax Slab"           : "taxslab",
            "Quantity"           : "inv_qty",
            "UQC"                : "uqc",
            "Taxable Amount"     : "inv_tot",
            "IGST"               : "igstamt",
            "CGST"               : "cgstamt",
            "SGST"               : "sgstamt",
            "UGST"               : "ugstamt",
            "Cess"               : "cessamt",
            "GSTR1_Section"      : "gstr1_section",
        }
        at_cols = {dst: assets_df[src] for src, dst in _col_map.items() if src in assets_df.columns}
        if "hsncode" in at_cols and "inv_tot" in at_cols:
            at_tmp = pd.DataFrame(at_cols).copy()
            at_tmp["Type"] = at_tmp.get("gstr1_section",
                pd.Series("B2C", index=at_tmp.index)).apply(
                    lambda x: "B2B" if str(x).upper() == "B2B" else "B2C")
            at_tmp["source"] = "Assets"
            # From State (Supplier) = fstate | From State Code = fstcode from Asset Sale sheet
            at_tmp["from_state"]  = at_tmp["fstate"].astype(str).str.strip()  if "fstate"  in at_tmp.columns else ""
            at_tmp["from_stcode"] = pd.to_numeric(at_tmp["fstcode"], errors="coerce").fillna(0).astype(int) if "fstcode" in at_tmp.columns else 0
            frames.append(at_tmp)

    # Add Cross Charge rows — show both From State and To State
    if cc_df is not None and not cc_df.empty:
        tv_col      = "Taxable Value" if "Taxable Value" in cc_df.columns else None
        gst_col     = "GST (18%)"    if "GST (18%)"    in cc_df.columns else None
        from_st_col = next((c for c in ["From State"]                   if c in cc_df.columns), None)
        from_sc_col = next((c for c in ["From State Code"]               if c in cc_df.columns), None)
        to_st_col   = next((c for c in ["To State","State"]              if c in cc_df.columns), None)
        to_sc_col   = next((c for c in ["To State Code","State Code"]    if c in cc_df.columns), None)
        if tv_col:
            cc_tmp = pd.DataFrame({
                # POS (To State) as Supplier State — groups CC under receiver state
                "fstate"      : cc_df[to_st_col].astype(str).str.strip().str.title() if to_st_col else "",
                "fstcode"     : pd.to_numeric(cc_df[to_sc_col], errors="coerce").fillna(0).astype(int) if to_sc_col else 0,
                # From State as additional columns
                "from_state"  : cc_df[from_st_col].astype(str).str.strip().str.title() if from_st_col else "",
                "from_stcode" : pd.to_numeric(cc_df[from_sc_col], errors="coerce").fillna(0).astype(int) if from_sc_col else 0,
                "hsncode"     : 998399,
                "taxslab"     : 18,
                "inv_qty"     : 0,
                "inv_tot"     : pd.to_numeric(cc_df[tv_col],  errors="coerce").fillna(0),
                "igstamt"     : pd.to_numeric(cc_df[gst_col], errors="coerce").fillna(0) if gst_col else 0,
                "cgstamt"     : 0, "sgstamt": 0, "ugstamt": 0, "cessamt": 0,
            })
            cc_tmp["Type"]   = "Cross Charge"
            cc_tmp["source"] = "CrossCharge"
            frames.append(cc_tmp)

    # Add Return rows — direct from sheet, no state filtering, no overrides
    if cgst_returns is not None and not cgst_returns.empty:
        ret = cgst_returns.copy()
        # Ensure numeric columns
        for _c in ["taxslab","inv_tot","cgstamt","sgstamt","igstamt","ugstamt","cessamt","inv_qty"]:
            if _c not in ret.columns: ret[_c] = 0
            ret[_c] = pd.to_numeric(ret[_c], errors="coerce").fillna(0)
        if "hsncode" in ret.columns:
            _need = ["fstate","fstcode","bill_state","bill_stcod",
                     "hsncode","taxslab","inv_qty","inv_tot",
                     "igstamt","cgstamt","sgstamt","ugstamt","cessamt"]
            ret_tmp = ret[[c for c in _need if c in ret.columns]].copy()

            # Direct: negate values exactly as in sheet
            for _c in ["inv_tot","inv_qty","igstamt","cgstamt","sgstamt","ugstamt","cessamt"]:
                if _c not in ret_tmp.columns: ret_tmp[_c] = 0
                ret_tmp[_c] = -pd.to_numeric(ret_tmp[_c], errors="coerce").fillna(0).abs()
            ret_tmp["uqc"] = "PCS"

            # From State (Supplier) = fstate  |  From State Code = fstcode
            ret_tmp["from_state"]  = ret_tmp["fstate"].astype(str).str.strip() if "fstate" in ret_tmp.columns else ""
            ret_tmp["from_stcode"] = pd.to_numeric(ret_tmp["fstcode"], errors="coerce").fillna(0).astype(int) if "fstcode" in ret_tmp.columns else 0

            # Place of Supply (POS) = bill_state  |  POS Code = bill_stcod
            ret_tmp["fstate"]  = ret_tmp["bill_state"].astype(str).str.strip() if "bill_state" in ret_tmp.columns else ret_tmp.get("fstate", "")
            ret_tmp["fstcode"] = pd.to_numeric(ret_tmp["bill_stcod"], errors="coerce").fillna(0).astype(int) if "bill_stcod" in ret_tmp.columns else ret_tmp.get("fstcode", 0)

            # Drop bill_state/bill_stcod — already mapped above
            ret_tmp.drop(columns=[c for c in ["bill_state","bill_stcod"] if c in ret_tmp.columns], inplace=True)

            ret_tmp["Type"]   = "Return"
            ret_tmp["source"] = "Return"
            frames.append(ret_tmp)

    if not frames: return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)
    df = num(df, ["inv_qty","inv_tot","igstamt","cgstamt","sgstamt","ugstamt","cessamt","taxslab"])

    # Normalise key groupby columns
    df["fstate"]   = df["fstate"].astype(str).str.strip().str.title()
    df["fstcode"]  = pd.to_numeric(df["fstcode"], errors="coerce").fillna(0).astype(int)
    df["hsncode"]  = pd.to_numeric(df["hsncode"], errors="coerce").fillna(0).astype(int)
    df["taxslab"]  = pd.to_numeric(df["taxslab"], errors="coerce").fillna(0)
    df["Type"]     = df["Type"].astype(str).str.strip()
    if "source" not in df.columns:
        df["source"] = "Sales"
    df["source"]   = df["source"].astype(str).str.strip()

    # Normalise from_state/from_stcode (only CC rows have these)
    if "from_state" not in df.columns:
        df["from_state"]  = ""
    if "from_stcode" not in df.columns:
        df["from_stcode"] = 0
    df["from_state"]  = df["from_state"].astype(str).str.strip().replace({"nan":"","None":"","NaN":""})
    df["from_stcode"] = pd.to_numeric(df["from_stcode"], errors="coerce").fillna(0).astype(int)

    # Normalise UQC (Unit Quantity Code — available from Assets/Returns)
    if "uqc" not in df.columns:
        df["uqc"] = ""
    df["uqc"] = df["uqc"].astype(str).str.strip().replace({"nan":"","None":""})

    # HSN Summary groupby — include UQC (first value per group)
    grp_cols = [c for c in ["fstate","fstcode","from_state","from_stcode","Type","source","hsncode","taxslab"]
                if c in df.columns]
    grp = df.groupby(grp_cols, as_index=False).agg(
        UQC          =("uqc","first"),
        Total_Qty    =("inv_qty","sum"),
        Taxable_Value=("inv_tot","sum"),
        IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
        SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
    )
    grp["Total_Tax"]     = grp["IGST"]+grp["CGST"]+grp["SGST"]+grp["UGST"]+grp["Cess"]
    grp["Invoice_Value"] = grp["Taxable_Value"]+grp["Total_Tax"]
    # Clear from_stcode where from_state is empty (non-CC rows)
    if "from_state" in grp.columns and "from_stcode" in grp.columns:
        grp["from_stcode"] = grp["from_stcode"].astype(object)   # allow mixed types
        grp.loc[grp["from_state"].astype(str).str.strip().eq(""), "from_stcode"] = ""

    grp.rename(columns={
        "fstate"      :"Place of Supply (POS)",
        "fstcode"     :"POS Code",
        "from_state"  :"From State (Supplier)",
        "from_stcode" :"From State Code",
        "hsncode"     :"HSN Code",
        "taxslab"     :"Tax Slab",
        "source"      :"Source",
    }, inplace=True)

    col_order = [
        "From State (Supplier)","From State Code",
        "Place of Supply (POS)","POS Code",
        "Type","Source",
        "HSN Code","UQC","Tax Slab",
        "Total_Qty",
        "Taxable_Value","IGST","CGST","SGST","UGST","Cess","Total_Tax","Invoice_Value",
    ]
    grp = grp[[c for c in col_order if c in grp.columns]]


    # Sort by Type first so B2B, B2C, Stock Transfer, Assets, CC, Return are grouped together
    type_order = {"B2B":1,"B2C":2,"Stock Transfer":3,"Cross Charge":4,"Assets":5,"Return":6}
    if "Type" in grp.columns:
        grp["_type_order"] = grp["Type"].map(type_order).fillna(99)
        sort_cols = [c for c in ["_type_order","Place of Supply (POS)","HSN Code","Tax Slab"]
                     if c in grp.columns]
        grp = grp.sort_values(sort_cols).drop(columns=["_type_order"]).reset_index(drop=True)
    else:
        sort_cols = [c for c in ["Place of Supply (POS)","HSN Code","Tax Slab"] if c in grp.columns]
        grp = grp.sort_values(sort_cols).reset_index(drop=True)
    return grp

# ─────────────────────────────────────────────────────────────────
# DOCUMENT SUMMARY – Table 13  (with gap/cancelled detection)
# ─────────────────────────────────────────────────────────────────
def build_doc_summary(sales_df, return_df=None, stock_df=None, cc_df=None, assets_df=None):
    rows          = []
    cancelled_rows = []

    def detect_gaps(inv_series_raw, doc_type, prefix_state_map=None, source_remark=None, series_length=None):
        """Vectorised gap detection — no Python for-loops over rows."""
        if prefix_state_map is None:
            prefix_state_map = {}

        # Unique invoice numbers as a Series
        inv_s = (inv_series_raw.dropna().astype(str).str.strip()
                 .replace("", np.nan).dropna())
        inv_s = pd.Series(inv_s.unique(), name="inv")

        # Split into prefix + numeric suffix using vectorised str.extract
        parsed = inv_s.str.extract(r'^(.*?)(\d+)$', expand=True)
        parsed.columns = ["prefix", "num_str"]
        parsed["inv"]  = inv_s.values

        std   = parsed.dropna(subset=["num_str"])
        other = parsed[parsed["num_str"].isna()]

        # Non-standard invoices (no trailing digits)
        if not other.empty:
            rows.append({
                "Document Type": doc_type, "Supplier State": "", "Supplier State Code": "",
                "Series": "Non-Standard", "From No": "-", "To No": "-",
                "Total Issued": len(other), "Cancelled": 0, "Net Issued": len(other),
                "Source Remarks": source_remark or "",
            })

        if std.empty:
            return

        std["num"] = pd.to_numeric(std["num_str"], errors="coerce")
        std["pad"] = std["num_str"].str.len()
        std = std.dropna(subset=["num"])
        std["num"] = std["num"].astype(int)

        # Exact width of each invoice number as it actually appears in the
        # source data — used instead of a group-wide max so a short number
        # (e.g. 64474) is never padded with a leading zero borrowed from a
        # longer number elsewhere in the same series (e.g. 134429).
        pad_lookup = (std.drop_duplicates(subset=["prefix", "num"])
                         .set_index(["prefix", "num"])["pad"])

        # Group by prefix — fully vectorised aggregation
        grp = std.groupby("prefix", sort=True).agg(
            min_n=("num", "min"), max_n=("num", "max"),
            present=("num", "nunique"),
        ).reset_index()

        grp["total_rng"] = grp["max_n"] - grp["min_n"] + 1
        grp["cancelled"] = grp["total_rng"] - grp["present"]

        for _, r in grp.iterrows():
            pfx      = r["prefix"]
            min_n    = int(r["min_n"])
            max_n    = int(r["max_n"])
            min_pad  = int(pad_lookup.get((pfx, min_n), len(str(min_n))))
            max_pad  = int(pad_lookup.get((pfx, max_n), len(str(max_n))))
            st_name, st_code = prefix_state_map.get(pfx, ("", ""))

            _series_display = (pfx[:series_length] if pfx and series_length else (pfx if pfx else "Numeric"))
            rows.append({
                "Document Type"      : doc_type,
                "Supplier State"     : st_name,
                "Supplier State Code": st_code,
                "Series"             : _series_display,
                "From No"            : f"{pfx}{str(min_n).zfill(min_pad)}",
                "To No"              : f"{pfx}{str(max_n).zfill(max_pad)}",
                "Total Issued"       : int(r["total_rng"]),
                "Cancelled"          : int(r["cancelled"]),
                "Net Issued"         : int(r["present"]),
                "Source Remarks"     : source_remark or "",
            })

            # Cancelled detail — only when gaps exist and series is manageable
            if r["cancelled"] > 0 and r["total_rng"] <= 50000:
                present_set = set(
                    std.loc[std["prefix"] == pfx, "num"].tolist()
                )
                missing = np.setdiff1d(
                    np.arange(min_n, max_n + 1), list(present_set)
                )
                for n in missing:
                    n = int(n)
                    cancelled_rows.append({
                        "Document Type"       : doc_type,
                        "Supplier State"      : st_name,
                        "Supplier State Code" : st_code,
                        "Series"              : pfx if pfx else "Numeric",
                        "Cancelled Invoice No": f"{pfx}{n}",
                    })

    if sales_df is not None and not sales_df.empty and "inv_no" in sales_df.columns:
        # Build prefix→(state,code) map — vectorised, no iterrows
        prefix_state_map = {}
        if "fstate" in sales_df.columns and "fstcode" in sales_df.columns:
            tmp = (sales_df[["inv_no","fstate","fstcode"]]
                   .drop_duplicates(subset=["inv_no"]).copy())
            tmp["prefix"] = tmp["inv_no"].astype(str).str.extract(r'^(.*?)(\d+)$')[0]
            tmp = tmp.dropna(subset=["prefix"])
            tmp = tmp.drop_duplicates(subset=["prefix"])
            prefix_state_map = dict(
                zip(tmp["prefix"], zip(tmp["fstate"], tmp["fstcode"]))
            )
        detect_gaps(sales_df["inv_no"], "Invoices for Outward Supply", prefix_state_map)

    if return_df is not None and not return_df.empty and "retufm_no" in return_df.columns:
        notes = return_df.drop_duplicates(subset=["retufm_no"])
        rows.append({
            "Document Type": "Credit / Debit Notes",
            "Series"       : "-",
            "From No"      : "-",
            "To No"        : "-",
            "Total Issued" : len(notes),
            "Cancelled"    : 0,
            "Net Issued"   : len(notes),
        })

    # ── Sales Return document series using Return Bill No from GST_Sales_Return.xlsx ──
    if return_df is not None and not return_df.empty:
        _rf = return_df.copy()
        # Auto-detect header row
        _known_r = {"fstate","inv_tot","cgstamt","igstamt","taxslab","retufm_no","return_bill_number","totval"}
        _cols_r  = {str(c).strip().lower() for c in _rf.columns}
        if not (_cols_r & _known_r):
            for i in range(min(10, len(_rf))):
                _rv = {str(v).strip().lower() for v in _rf.iloc[i].dropna().tolist()}
                if _rv & _known_r:
                    _rf.columns = [str(v).strip() for v in _rf.iloc[i].tolist()]
                    _rf = _rf.iloc[i+1:].reset_index(drop=True)
                    break
        _rf = _rf.dropna(how="all")

        # Find Return Bill Number column
        _rbn_col = None
        for _c in ["Return_Bill_Number","retufm_no","Return Bill Number","ReturnBillNo","return_bill_number"]:
            if _c in _rf.columns:
                _rbn_col = _c
                break

        if _rbn_col:
            # Build prefix→state map from fstate/fstcode
            _ret_prefix_map = {}
            if "fstate" in _rf.columns and "fstcode" in _rf.columns:
                _tmp = (_rf[[_rbn_col,"fstate","fstcode"]]
                        .drop_duplicates(subset=[_rbn_col]).copy())
                _tmp["prefix"] = _tmp[_rbn_col].astype(str).str.extract(r'^(.*?)(\d+)$')[0]
                _tmp = _tmp.dropna(subset=["prefix"]).drop_duplicates(subset=["prefix"])
                _ret_prefix_map = dict(zip(_tmp["prefix"], zip(_tmp["fstate"], _tmp["fstcode"])))
            # Clean and deduplicate return bill numbers
            _rbn_series = (_rf[_rbn_col].astype(str).str.strip()
                           .replace({"nan":"","NaN":"","None":"","none":""})
                           .dropna())
            _rbn_series = _rbn_series[_rbn_series != ""]
            if not _rbn_series.empty:
                detect_gaps(_rbn_series, "Credit note", _ret_prefix_map, source_remark="Sales Return", series_length=3)

    if stock_df is not None and not stock_df.empty and "inv_no" in stock_df.columns:
        st_prefix_map = {}
        if "fstate" in stock_df.columns and "fstcode" in stock_df.columns:
            _st = (stock_df[["inv_no","fstate","fstcode"]]
                   .drop_duplicates(subset=["inv_no"]).copy())
            _st["prefix"] = _st["inv_no"].astype(str).str.extract(r'^(.*?)(\d+)$')[0]
            _st = _st.dropna(subset=["prefix"]).drop_duplicates(subset=["prefix"])
            st_prefix_map = dict(zip(_st["prefix"], zip(_st["fstate"], _st["fstcode"])))
        detect_gaps(stock_df["inv_no"], "Invoices for Outward Supply", st_prefix_map, source_remark="Stock Transfer (Delivery Challan)", series_length=3)

    # Assets Sale document series
    # Total Issued = inv_tot (sum per invoice), Cancelled = 0, Net Issued = 1 per invoice
    if assets_df is not None and not assets_df.empty:
        _af = assets_df.copy()
        # Header detection
        _known_a = {"inv_no","invoice no","invoice no.","taxable amount","supplier state","fstate"}
        _cols_a  = {str(c).strip().lower() for c in _af.columns}
        if not (_cols_a & _known_a):
            for i in range(min(3, len(_af))):
                _rv = {str(v).strip().lower() for v in _af.iloc[i].dropna().tolist()}
                if _rv & _known_a:
                    _af.columns = [str(v).strip() for v in _af.iloc[i].tolist()]
                    _af = _af.iloc[i+1:].reset_index(drop=True)
                    break
        _af = _af.dropna(how="all")
        # Normalise column names
        _af.rename(columns={"inv_no":"Invoice No","Invoice No.":"Invoice No",
                             "fstate":"Supplier State","fstcode":"Supplier State Code",
                             "inv_tot":"inv_tot"}, inplace=True)
        _inv_col = next((c for c in ["Invoice No","inv_no","InvoiceNo"] if c in _af.columns), None)
        _st_col  = next((c for c in ["Supplier State","fstate"] if c in _af.columns), None)
        _stc_col = next((c for c in ["Supplier State Code","fstcode"] if c in _af.columns), None)
        _tot_col = next((c for c in ["inv_tot","Taxable Amount","Taxable_Amount"] if c in _af.columns), None)

        if _inv_col:
            # One row per state — From No = first invoice, To No = last invoice, Total=1, Net=1
            import re as _re3
            _st_grp_cols = [c for c in [_st_col, _stc_col] if c]
            if _st_grp_cols:
                for _keys, _grp in _af.groupby(_st_grp_cols):
                    _keys = _keys if isinstance(_keys, tuple) else (_keys,)
                    _st_name = str(_keys[0]).strip() if len(_keys) > 0 else ""
                    _st_code = _keys[1]              if len(_keys) > 1 else ""
                    _inv_list = (_grp[_inv_col].astype(str).str.strip()
                                 .replace({"nan":"","None":"","none":""})
                                 .dropna())
                    _inv_list = [v for v in _inv_list.tolist() if v]
                    if not _inv_list:
                        continue
                    _first_inv = _inv_list[0]
                    _last_inv  = _inv_list[-1]
                    _m   = _re3.match(r'^(.*?)(\d+)$', _first_inv)
                    _pfx = _m.group(1) if _m else _first_inv
                    rows.append({
                        "Document Type"      : "Invoices for Outward Supply",
                        "Source Remarks"     : "Assets Sale",
                        "Supplier State"     : _st_name,
                        "Supplier State Code": _st_code,
                        "Series"             : _pfx,
                        "From No"            : _first_inv,
                        "To No"              : _last_inv,
                        "Total Issued"       : 1,
                        "Cancelled"          : 0,
                        "Net Issued"         : 1,
                    })
            else:
                _inv_list = [v for v in _af[_inv_col].astype(str).str.strip().tolist()
                             if v and v.lower() not in ["nan","none",""]]
                if _inv_list:
                    _m = _re3.match(r'^(.*?)(\d+)$', _inv_list[0])
                    rows.append({"Document Type":"Invoices for Outward Supply","Source Remarks":"Assets Sale",
                                 "Series":_m.group(1) if _m else "",
                                 "From No":_inv_list[0],"To No":_inv_list[-1],
                                 "Total Issued":1,"Cancelled":0,"Net Issued":1})
        else:
            rows.append({"Document Type":"Invoices for Outward Supply","Source Remarks":"Assets Sale",
                         "Series":"-","From No":"-","To No":"-","Total Issued":0,"Cancelled":0,"Net Issued":0})

    # Cross Charge document series
    # Supplier State = From State (from title), Total/Net Issued = Invoice Value (monetary)
    if cc_df is not None and not cc_df.empty:
        _cf_raw = cc_df.copy()
        # ── Extract From State from title row 0 ──────────────────────
        _from_state = ""
        _from_code  = 0
        try:
            _title_txt = " ".join(_cf_raw.iloc[0].dropna().astype(str).tolist()).upper()
            for _sc, _sn in GSTIN_STATE_MAP.items():
                if _sn.upper() in _title_txt:
                    _from_state = _sn
                    _from_code  = int(_sc) if _sc.isdigit() else 0
                    break
        except Exception:
            pass

        # ── Detect header row ──────────────────────────────────────
        _cf = _cf_raw.copy()
        _known_c = {"invoice no","invoice no.","invoice amt.","states","gst no."}
        _cols_c  = {str(c).strip().lower() for c in _cf.columns}
        if not (_cols_c & _known_c):
            for i in range(min(3, len(_cf))):
                _rv = {str(v).strip().lower() for v in _cf.iloc[i].dropna().tolist()}
                if _rv & _known_c:
                    _cf.columns = [str(v).strip() for v in _cf.iloc[i].tolist()]
                    _cf = _cf.iloc[i+1:].reset_index(drop=True)
                    break
        _cf = _cf.dropna(how="all")
        _cf.rename(columns={"Invoice No.":"Invoice No","inv_no":"Invoice No"}, inplace=True)

        _inv_col  = next((c for c in ["Invoice No","inv_no","InvoiceNo"] if c in _cf.columns), None)
        _val_col  = next((c for c in ["Invoice Amt.","Invoice Amt","Invoice Value","invoice_val"] if c in _cf.columns), None)

        if _inv_col:
            # Single row — From No = first, To No = last, Total/Net = count of invoices
            import re as _re2
            _valid_inv = (_cf[_inv_col].astype(str).str.strip()
                          .replace({"nan":"","None":"","none":""})
                          .dropna())
            _valid_inv = _valid_inv[_valid_inv != ""].tolist()
            if _valid_inv:
                _first_inv = _valid_inv[0]
                _last_inv  = _valid_inv[-1]
                _count     = len(_valid_inv)
                _m   = _re2.match(r'^(.*?)(\d+)$', _first_inv)
                _pfx = _m.group(1) if _m else _first_inv
                rows.append({
                    "Document Type"      : "Invoices for Outward Supply",
                    "Source Remarks"     : "Cross Charge",
                    "Supplier State"     : _from_state,
                    "Supplier State Code": _from_code,
                    "Series"             : _pfx,
                    "From No"            : _first_inv,
                    "To No"              : _last_inv,
                    "Total Issued"       : _count,
                    "Cancelled"          : 0,
                    "Net Issued"         : _count,
                })
        else:
            rows.append({"Document Type":"Invoices for Outward Supply","Source Remarks":"Cross Charge",
                         "Series":"-","From No":"-","To No":"-","Total Issued":0,"Cancelled":0,"Net Issued":0})

    doc_df       = pd.DataFrame(rows)        if rows        else pd.DataFrame()
    cancelled_df = pd.DataFrame(cancelled_rows) if cancelled_rows else pd.DataFrame()
    return doc_df, cancelled_df

# ─────────────────────────────────────────────────────────────────
# TAX SUMMARY
# ─────────────────────────────────────────────────────────────────
def build_tax_summary(sales_df, return_df=None, stock_df=None):
    _tax_need = ["fstate","fstcode","taxslab","inv_tot","igstamt","cgstamt","sgstamt","ugstamt","cessamt"]
    frames = []
    for df, lbl in [(sales_df,"Sales"),(return_df,"Returns (Credit Notes)"),(stock_df,"Stock Transfer")]:
        if df is not None and not df.empty:
            cols = [c for c in _tax_need if c in df.columns]
            tmp = df[cols].copy()   # select only needed columns before copy
            tmp["_src"] = lbl
            frames.append(tmp)
    if not frames: return pd.DataFrame()
    all_df = pd.concat(frames, ignore_index=True)
    all_df = num(all_df, ["inv_tot","igstamt","cgstamt","sgstamt","ugstamt","cessamt","taxslab"])
    grp_cols = [c for c in ["_src","fstate","fstcode","taxslab"] if c in all_df.columns]
    grp = all_df.groupby(grp_cols, as_index=False).agg(
        Taxable_Value=("inv_tot","sum"),
        IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
        SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
    )
    grp["Total_Tax"]     = grp["IGST"]+grp["CGST"]+grp["SGST"]+grp["UGST"]+grp["Cess"]
    grp["Invoice_Value"] = grp["Taxable_Value"]+grp["Total_Tax"]
    grp.rename(columns={
        "_src":"Source","fstate":"Supplier State",
        "fstcode":"Supplier State Code","taxslab":"Tax Slab",
    }, inplace=True)
    return grp

# ─────────────────────────────────────────────────────────────────
# COMBINED GSTR-1 MASTER SHEET
# Creates a single flat table with all transactions tagged by section
# ─────────────────────────────────────────────────────────────────
def build_combined(sales_df, return_df=None, stock_df=None, cc_df=None, assets_df=None, cgst_returns=None):  # noqa: cc_df/assets_df kept for API compatibility
    """Invoice-level combined sheet using single groupby agg — fast & clean."""
    parts = []

    # ── Sales (B2B / B2CL / B2CS) ────────────────────────────────
    if sales_df is not None and not sales_df.empty and "inv_no" in sales_df.columns:
        # Select only the columns needed — avoids copying the full 1.87M-row DataFrame
        _need = [c for c in ["inv_no","taxslab","gstr1_section","gstin","mscname",
                              "fstate","fstcode","gstin2","inv_date","ship_state","ship_stcod",
                              "inv_tot","totval","igstamt","cgstamt","sgstamt","ugstamt","cessamt"]
                 if c in sales_df.columns]
        df = num(sales_df[_need].copy(), TAX_COLS)
        # Convert float32 → float64 before aggregation to avoid precision accumulation errors
        for _fc in ["inv_tot","igstamt","cgstamt","sgstamt","ugstamt","cessamt"]:
            if _fc in df.columns:
                df[_fc] = df[_fc].astype("float64")
        if "totval" not in df.columns:
            df["totval"] = 0
        agg = df.groupby(["inv_no","taxslab"], as_index=False).agg(
            GSTR1_Section=("gstr1_section","first"),
            Supplier_GSTIN=("gstin","first"),
            Supplier_Name=("mscname","first"),
            Supplier_State=("fstate","first"),
            Supplier_State_Code=("fstcode","first"),
            Receiver_GSTIN=("gstin2","first"),
            Invoice_Date=("inv_date","first"),
            Place_of_Supply=("ship_state","first"),
            POS_Code=("ship_stcod","first"),
            Taxable_Value=("inv_tot","sum"),
            IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
            SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
            Total_Invoice_Value=("totval","first"),
        ).rename(columns={"inv_no":"Invoice_No","taxslab":"Tax Slab"})
        agg["Invoice_Value"] = agg["Taxable_Value"]+agg["IGST"]+agg["CGST"]+agg["SGST"]+agg["UGST"]+agg["Cess"]
        agg["Source"] = "State Sales"
        agg["Note_Type"] = ""
        parts.append(agg)

    # ── Returns (CDNR / CDNUR) ───────────────────────────────────
    if return_df is not None and not return_df.empty:
        df = return_df.copy()
        df["gstin2"] = df["gstin2"].astype(str).str.strip().replace({"nan":"","NaN":"","None":""})
        df = num(df, TAX_COLS)
        ship_col = "ship_stcode" if "ship_stcode" in df.columns else "ship_stcod"
        df["sec"] = np.where(df["gstin2"].str.len() == 15, "CDNR", "CDNUR")
        note_key = "retufm_no" if "retufm_no" in df.columns else "inv_no"
        date_col = "retufm_dt" if "retufm_dt" in df.columns else "inv_date"
        rtype    = "RETURNTYPE" if "RETURNTYPE" in df.columns else None
        agg_dict = {
            "GSTR1_Section" : ("sec","first"),
            "Supplier_GSTIN": ("gstin","first"),
            "Supplier_Name" : ("mscname","first"),
            "Receiver_GSTIN": ("gstin2","first"),
            "Invoice_Date"  : (date_col,"first"),
            "Place_of_Supply": (ship_col,"first"),
            "Taxable_Value"  : ("inv_tot","sum"),
            "IGST": ("igstamt","sum"), "CGST": ("cgstamt","sum"),
            "SGST": ("sgstamt","sum"), "UGST": ("ugstamt","sum"), "Cess": ("cessamt","sum"),
        }
        if rtype: agg_dict["Note_Type"] = (rtype,"first")
        grp_keys = [note_key, "taxslab"] if "taxslab" in df.columns else [note_key]
        agg = df.groupby(grp_keys, as_index=False).agg(**agg_dict).rename(
            columns={note_key:"Invoice_No","taxslab":"Tax Slab"}
        )
        # Negate return values in Combined sheet
        for _c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess"]:
            if _c in agg.columns: agg[_c] = -agg[_c].abs()
        agg["Invoice_Value"] = agg["Taxable_Value"]+agg["IGST"]+agg["CGST"]+agg["SGST"]+agg["UGST"]+agg["Cess"]
        agg["Source"] = "Sales Return"
        agg["POS_Code"] = ""
        if "Note_Type" not in agg.columns: agg["Note_Type"] = ""
        parts.append(agg)

    # ── Stock Transfer ───────────────────────────────────────────
    if stock_df is not None and not stock_df.empty and "inv_no" in stock_df.columns:
        df = num(stock_df.copy(), TAX_COLS)
        if "totval" not in df.columns:
            df["totval"] = 0
        agg = df.groupby(["inv_no","taxslab"], as_index=False).agg(
            Supplier_GSTIN      =("gstin",     "first"),
            Supplier_Name       =("mscname",   "first"),
            Supplier_State      =("fstate",    "first"),
            Supplier_State_Code =("fstcode",   "first"),
            Receiver_GSTIN      =("gstin2",    "first"),
            Invoice_Date        =("inv_date",  "first"),
            Place_of_Supply     =("ship_state","first"),
            POS_Code            =("ship_stcod","first"),
            Taxable_Value       =("inv_tot",   "sum"),
            IGST=("igstamt","sum"), CGST=("cgstamt","sum"),
            SGST=("sgstamt","sum"), UGST=("ugstamt","sum"), Cess=("cessamt","sum"),
            Total_Invoice_Value =("totval",    "first"),
        ).rename(columns={"inv_no":"Invoice_No","taxslab":"Tax Slab"})
        agg["Invoice_Value"] = agg["Taxable_Value"]+agg["IGST"]+agg["CGST"]+agg["SGST"]+agg["UGST"]+agg["Cess"]
        agg["GSTR1_Section"] = "Stock Transfer"
        agg["Source"] = "Stock Transfer"
        agg["Note_Type"] = ""
        parts.append(agg)

    # ── Sales Returns (Taken) — CGST returns as negative rows ────
    if cgst_returns is not None and not cgst_returns.empty:
        ret = cgst_returns.copy()
        # Ensure required columns exist
        for col in ["inv_tot","cgstamt","sgstamt","igstamt","ugstamt","cessamt","taxslab"]:
            if col not in ret.columns:
                ret[col] = 0
            ret[col] = pd.to_numeric(ret[col], errors="coerce").fillna(0)

        # Clean string columns from GST_Sales_Return.xlsx — replace NaN-like with ""
        _nan_vals = ["nan","NaN","None","none","<NA>"]
        for _sc in ["retufm_no", "Return_Bill_Number"]:
            if _sc in ret.columns:
                ret[_sc] = (ret[_sc].fillna("").astype(str).str.strip()
                            .replace(_nan_vals, ""))

        date_col = "retufm_dt" if "retufm_dt" in ret.columns else ("inv_date" if "inv_date" in ret.columns else None)

        # Deduplicate column names before groupby
        ret = ret.loc[:, ~ret.columns.duplicated()].copy()

        # Include retufm_no and Return_Bill_Number directly as groupby keys
        # so they appear as filled columns in the output without any extra aggregation
        grp_keys = [k for k in ["Return_Bill_Number","retufm_no","fstate","fstcode","taxslab"]
                    if k and k in ret.columns]

        # Need at least one identifier key
        if not any(k in grp_keys for k in ["Return_Bill_Number","retufm_no"]):
            for _fb in ["inv_no"]:
                if _fb in ret.columns:
                    grp_keys = [_fb] + [k for k in grp_keys if k != _fb]
                    break

        if grp_keys:
            agg_dict = {
                "Taxable_Value":("inv_tot","sum"),
                "IGST"         :("igstamt","sum"), "CGST":("cgstamt","sum"),
                "SGST"         :("sgstamt","sum"), "UGST":("ugstamt","sum"),
                "Cess"         :("cessamt","sum"),
            }
            if date_col and date_col in ret.columns:
                agg_dict["Invoice_Date"] = (date_col, "first")
            # Place of Supply = bill_state | POS Code = bill_stcod from return sheet
            if "bill_state" in ret.columns:
                agg_dict["Place_of_Supply"] = ("bill_state", "first")
            if "bill_stcod" in ret.columns:
                agg_dict["POS_Code"] = ("bill_stcod", "first")

            agg_ret = ret.groupby(grp_keys, as_index=False).agg(**agg_dict)
            for _c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess"]:
                agg_ret[_c] = -agg_ret[_c].abs()
            agg_ret["Invoice_Value"] = (agg_ret["Taxable_Value"] + agg_ret["IGST"] +
                                        agg_ret["CGST"] + agg_ret["SGST"] +
                                        agg_ret["UGST"] + agg_ret["Cess"])

            # ── Total_Invoice_Value: compound key (retufm_no + Return_Bill_Number + inv_no) ──
            # Step 1: get first totval per unique (retufm_no, Return_Bill_Number, inv_no)
            # Step 2: sum those first-values per the main groupby keys → correct total per row
            if "totval" in ret.columns:
                ret["totval"] = pd.to_numeric(ret["totval"], errors="coerce").fillna(0)
                _tv_key_cols = [k for k in ["Return_Bill_Number","retufm_no","inv_no","taxslab"]
                                if k in ret.columns]
                if _tv_key_cols:
                    # first() per unique invoice line → avoids row duplication
                    _tv_dedup = (ret.groupby(_tv_key_cols, as_index=False)["totval"]
                                 .first())
                    # sum deduped totval per main grp_keys (merge keys = overlap)
                    _merge_keys = [k for k in grp_keys if k in _tv_dedup.columns]
                    if _merge_keys:
                        _tv_sum = (_tv_dedup.groupby(_merge_keys, as_index=False)["totval"]
                                   .sum()
                                   .rename(columns={"totval":"Total_Invoice_Value"}))
                        _tv_sum["Total_Invoice_Value"] = -_tv_sum["Total_Invoice_Value"].abs()
                        agg_ret = agg_ret.merge(_tv_sum, on=_merge_keys, how="left")
                        agg_ret["Total_Invoice_Value"] = agg_ret["Total_Invoice_Value"].fillna(0)

            # Rename keys to output column names
            agg_ret.rename(columns={
                "Return_Bill_Number": "Invoice_No",
                "fstate"            : "Supplier_State",
                "fstcode"           : "Supplier_State_Code",
                "taxslab"           : "Tax Slab",
            }, inplace=True)

            # Ensure Return_Bill_Number column exists (copy of Invoice_No)
            if "Return_Bill_Number" not in agg_ret.columns and "Invoice_No" in agg_ret.columns:
                agg_ret["Return_Bill_Number"] = agg_ret["Invoice_No"]

            agg_ret["GSTR1_Section"] = "Sales Return (Taken)"
            agg_ret["Source"]        = "Sales Return"
            agg_ret["Note_Type"]     = "CGST Return"
            parts.append(agg_ret)

    if not parts:
        return pd.DataFrame()

    combined = pd.concat(parts, ignore_index=True)

    # Add Total_Tax column = sum of all tax components
    for _tc in ["IGST","CGST","SGST","UGST","Cess"]:
        if _tc not in combined.columns:
            combined[_tc] = 0
        combined[_tc] = pd.to_numeric(combined[_tc], errors="coerce").fillna(0)
    combined["Total_Tax"] = combined["IGST"] + combined["CGST"] + combined["SGST"] + combined["UGST"] + combined["Cess"]

    # Fill Total_Invoice_Value: use totval sum where available, else fall back to Invoice_Value
    if "Total_Invoice_Value" not in combined.columns:
        combined["Total_Invoice_Value"] = pd.to_numeric(
            combined.get("Invoice_Value", 0), errors="coerce").fillna(0)
    else:
        iv  = pd.to_numeric(combined["Invoice_Value"],       errors="coerce").fillna(0)
        tiv = pd.to_numeric(combined["Total_Invoice_Value"], errors="coerce").fillna(0)
        combined["Total_Invoice_Value"] = tiv.where(tiv != 0, iv)
    # Rename Total_Invoice_Value → "Total Value" (shows sum of totval from source file)
    combined.rename(columns={"Total_Invoice_Value": "Total Value"}, inplace=True)

    col_order = ["GSTR1_Section","Source","Supplier_GSTIN","Supplier_Name",
                 "Supplier_State","Supplier_State_Code","Receiver_GSTIN",
                 "Invoice_No","retufm_no","Return_Bill_Number","Invoice_Date","Place_of_Supply","POS_Code",
                 "Taxable_Value","IGST","CGST","SGST","UGST","Cess",
                 "Invoice_Value","Total Value","Tax Slab","Note_Type"]
    return combined[[c for c in col_order if c in combined.columns]]

# ─────────────────────────────────────────────────────────────────
# EXCEL WRITER
# ─────────────────────────────────────────────────────────────────
def write_excel(sheets_dict: dict, period: str) -> bytes:
    _nan_strs = {"nan","NaN","None","none","<NA>"}
    _tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    _tmp.close()
    try:
        # constant_memory=True: writes each cell directly to disk — no shared string
        # table buffered in RAM, so large sheets with many unique strings won't OOM.
        with xlsxwriter.Workbook(_tmp.name, {
            "nan_inf_to_errors"  : True,
            "strings_to_numbers" : False,
            "constant_memory"    : True,
        }) as wb:
            hdr   = wb.add_format({"bold":True,"bg_color":"#1e3a5f","font_color":"white","border":1,"align":"center","valign":"vcenter","text_wrap":True})
            num_f = wb.add_format({"num_format":"#,##0.00","border":1})
            txt   = wb.add_format({"border":1})
            ttl   = wb.add_format({"bold":True,"font_size":12,"bg_color":"#d6e4f0","border":1})

            for sname, df in sheets_dict.items():
                ws = wb.add_worksheet(sname[:31])
                if df is None or (isinstance(df, pd.DataFrame) and df.empty):
                    ws.write(0, 0, f"No data – {sname}", ttl)
                    continue

                cols    = list(df.columns)
                add_rem = "Remarks" not in df.columns
                ncols   = len(cols) + (1 if add_rem else 0)

                # constant_memory does not support merge_range — write title as single cell
                ws.set_row(0, 22)
                for ci, col in enumerate(cols):
                    ws.set_column(ci, ci, max(16, len(str(col)) + 4))
                if add_rem:
                    ws.set_column(ncols - 1, ncols - 1, 20)

                # Row 0: title
                ws.write(0, 0, f"{sname}  |  Period: {period}", ttl)
                # Row 1: headers
                for ci, col in enumerate(cols):
                    ws.write(1, ci, col, hdr)
                if add_rem:
                    ws.write(1, len(cols), "Remarks", hdr)

                num_cols = set(df.select_dtypes(include=[np.number]).columns)

                # Pre-convert each column to a list once — then write row by row
                col_lists = []
                for col in cols:
                    series = df.iloc[:, cols.index(col)]
                    if col in num_cols:
                        col_lists.append(pd.to_numeric(series, errors="coerce").fillna(0).tolist())
                    else:
                        col_lists.append(
                            series.astype(str).str.strip().replace(_nan_strs, "").tolist()
                        )

                nrows = len(df)
                for ri in range(nrows):
                    for ci, col in enumerate(cols):
                        val = col_lists[ci][ri]
                        if col in num_cols:
                            ws.write_number(ri + 2, ci, float(val) if val != "" else 0, num_f)
                        else:
                            ws.write_string(ri + 2, ci, str(val), txt)
                    if add_rem:
                        ws.write_string(ri + 2, len(cols), "", txt)

        with open(_tmp.name, "rb") as _f:
            return _f.read()
    finally:
        try:
            os.unlink(_tmp.name)
        except Exception:
            pass

# ─────────────────────────────────────────────────────────────────
# MAIN: PROCESS ON BUTTON CLICK
# ─────────────────────────────────────────────────────────────────
if process_btn:
    using_local = (load_mode == "📂 From Local Folder (Recommended)")

    if using_local:
        if not data_folder or not os.path.isdir(data_folder):
            st.error(f"Folder not found: {data_folder}")
            st.stop()
        loc = detect_local_files(data_folder)
        if not loc["state_csvs"] and not loc["stock"] and not loc["sales_return"]:
            st.error("No source files detected in the folder. Check the folder path.")
            st.stop()
    else:
        if not any([state_files, stock_file, return_file, crosscharge_file, assets_file]):
            st.error("Please upload at least one source file.")
            st.stop()

    prog      = st.progress(0)
    pct_box   = st.empty()   # large percentage number
    status    = st.empty()

    def update_progress(pct, msg=""):
        prog.progress(pct)
        pct_box.markdown(
            f"<div style='text-align:center;font-size:40px;font-weight:bold;"
            f"color:#1e3a5f;margin:-8px 0 4px'>{pct}%</div>",
            unsafe_allow_html=True
        )
        if msg:
            status.info(msg)

    # ── Load state sales — parallel reads ────────────────────────
    update_progress(5, "📂 Loading state-wise sales files in parallel…")
    sales_df = pd.DataFrame()

    def _read_csv_path(path):
        try:
            return read_csv_safe(path, low_memory=False, usecols=lambda c: c in _SALES_USECOLS, dtype=_SALES_DTYPE)
        except Exception as e:
            return (os.path.basename(path), str(e))

    if using_local and loc["state_csvs"]:
        workers = min(8, len(loc["state_csvs"]))
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futures = {ex.submit(_read_csv_path, p): p for p in loc["state_csvs"]}
            frames  = []
            for fut in as_completed(futures):
                res = fut.result()
                if isinstance(res, tuple):
                    st.warning(f"Skipped {res[0]}: {res[1]}")
                else:
                    frames.append(res)
        if frames:
            sales_df = pd.concat(frames, ignore_index=True)
            del frames          # free individual CSV frames immediately
            import gc; gc.collect()
            sales_df = classify_sales(sales_df)
    elif state_files:
        frames = []
        for f in state_files:
            try: frames.append(read_csv_safe(f, low_memory=False, usecols=lambda c: c in _SALES_USECOLS, dtype=_SALES_DTYPE))
            except Exception as e: st.warning(f"Skipped {f.name}: {e}")
        if frames:
            sales_df = pd.concat(frames, ignore_index=True)
            sales_df = classify_sales(sales_df)
    update_progress(20, f"✅ 20% — Loaded {len(sales_df):,} sales rows")

    # ── Load stock transfer ───────────────────────────────────────
    _ST_COLS = {"gstin","mscname","gstin2","inv_no","inv_date","fstate","fstcode",
                "ship_state","ship_stcod","hsncode","taxslab","inv_qty","inv_tot",
                "totval","igstamt","cgstamt","sgstamt","ugstamt","cessamt","freight"}
    stock_df = pd.DataFrame()
    if using_local and loc["stock"]:
        try: stock_df = read_csv_safe(loc["stock"], low_memory=False, usecols=lambda c: c in _ST_COLS)
        except Exception as e: st.warning(f"Stock Transfer error: {e}")
    elif stock_file:
        try: stock_df = read_csv_safe(stock_file, low_memory=False, usecols=lambda c: c in _ST_COLS)
        except Exception as e: st.warning(f"Stock Transfer error: {e}")
    update_progress(35, "📂 35% — Loading Stock Transfer…")

    # ── Load sales return ─────────────────────────────────────────
    update_progress(40, "📂 40% — Loading Sales Return data…")
    return_df = pd.DataFrame()
    if using_local and loc["sales_return"]:
        try: return_df = pd.read_excel(loc["sales_return"], engine="calamine", header=None)
        except Exception as e: st.warning(f"Sales Return error: {e}")
    elif return_file:
        try: return_df = pd.read_excel(return_file, engine="calamine", header=None)
        except Exception as e: st.warning(f"Sales Return error: {e}")

    # ── Filter: Lapsed / IGST (Not Taken) / CGST (Taken) ─────────
    igst_returns, cgst_returns, lapsed_df = filter_sales_returns(return_df)
    # Use cgst_returns (already header-detected + normalised) for the summary
    # so ASC Remarks column is found by its proper name, not raw column index
    sales_return_summary_df = build_sales_return_summary(cgst_returns if not cgst_returns.empty else return_df)

    update_progress(55, "📂 55% — Loading Cross Charge…")

    # ── Load cross charge ─────────────────────────────────────────
    cc_raw = pd.DataFrame()
    if using_local and loc["cross_charge"]:
        try: cc_raw = pd.read_excel(loc["cross_charge"], header=None, engine="openpyxl")
        except Exception as e: st.warning(f"Cross Charge error: {e}")
    elif crosscharge_file:
        try: cc_raw = pd.read_excel(crosscharge_file, header=None, engine="openpyxl")
        except Exception as e: st.warning(f"Cross Charge error: {e}")
    update_progress(62, "📂 62% — Loading Assets…")

    # ── Load assets ───────────────────────────────────────────────
    assets_raw = pd.DataFrame()
    if using_local and loc["assets"]:
        try: assets_raw = pd.read_excel(loc["assets"], engine="openpyxl", header=None)
        except Exception as e: st.warning(f"Assets error: {e}")
    elif assets_file:
        try: assets_raw = pd.read_excel(assets_file, engine="openpyxl", header=None)
        except Exception as e: st.warning(f"Assets error: {e}")
    update_progress(68, "⚙️ 68% — Generating GSTR-1 sections in parallel…")

    # ── Generate all sections in parallel ────────────────────────
    _igst = igst_returns if not igst_returns.empty else None   # Not Taken → CDNR/CDNUR
    _cgst = cgst_returns if not cgst_returns.empty else None   # Taken → net B2CS/HSN
    _ret  = _igst                                              # Combined uses IGST returns only
    _stk  = stock_df  if not stock_df.empty  else None

    # b2cs excluded from parallel — runs sequentially after to avoid threading
    # conflict between cgst_returns (returns DataFrame) and sales_df access
    _tasks = {
        "b2b"     : lambda: build_b2b(sales_df),
        "b2cl"    : lambda: build_b2cl(sales_df),
        "cdnr"    : lambda: build_cdnr(igst_returns),
        "cdnur"   : lambda: build_cdnur(igst_returns),
        "st"      : lambda: build_stock_transfer(stock_df),
        "cc"      : lambda: build_cross_charge(cc_raw),
        "assets"  : lambda: build_assets(assets_raw),
        "hsn"     : lambda: pd.DataFrame(),
        "doc"     : lambda: build_doc_summary(sales_df, return_df, stock_df, cc_raw, assets_raw),
        "tax"     : lambda: build_tax_summary(sales_df, _ret, _stk),
    }
    # max_workers=2: each build_* holds its own copy of a large frame, so
    # fewer concurrent tasks = lower peak memory (identical output, same tasks)
    _results = {}
    with ThreadPoolExecutor(max_workers=2) as ex:
        fut_map = {ex.submit(fn): name for name, fn in _tasks.items()}
        for fut in as_completed(fut_map):
            _results[fut_map[fut]] = fut.result()

    # Free raw input frames — every consumer ran inside the executor above.
    # _tasks/fut_map lambdas close over these frames, so drop them first.
    del _tasks, fut_map, cc_raw, assets_raw, return_df
    import gc; gc.collect()

    b2b_df      = _results["b2b"]
    b2cl_df     = _results["b2cl"]
    # Build b2cs SEQUENTIALLY after parallel batch — ensures cgst_returns is fully available
    b2cs_df     = build_b2cs(sales_df=sales_df, cgst_returns=None)
    cdnr_df     = _results["cdnr"]
    cdnur_df    = _results["cdnur"]
    st_df       = _results["st"]
    cc_df       = _results["cc"]
    assets_df   = _results["assets"]
    doc_df, cancelled_df = _results["doc"]
    tax_df      = _results["tax"]
    # Build HSN and Combined after parallel batch so they can use cc_df
    hsn_df      = build_hsn(sales_df, cc_df=cc_df,
                             stock_df=stock_df if not stock_df.empty else None,
                             cgst_returns=_cgst,
                             assets_df=assets_df if not assets_df.empty else None)
    try:
        combined_df = build_combined(sales_df, _ret, _stk, cc_df, assets_df, cgst_returns=_cgst)
    except Exception as _ce:
        st.warning(f"⚠️ Combined GSTR-1 build error: {_ce}")
        combined_df = pd.DataFrame()

    # Free return/stock frames — build_hsn/build_combined were their last consumers.
    # (sales_df stays: it is stored in session_state for the display section.)
    del igst_returns, cgst_returns, stock_df, _igst, _cgst, _ret, _stk
    gc.collect()
    update_progress(90, "⚡ 90% — Building Excel file…")

    # ── Merge sales return summary into B2CS ─────────────────────
    b2cs_excel_df = b2cs_df.copy() if b2cs_df is not None and not b2cs_df.empty else pd.DataFrame()
    if sales_return_summary_df is not None and not sales_return_summary_df.empty:
        ret_rows = sales_return_summary_df.copy()
        ret_rows["Source"] = "Return"
        b2cs_excel_df = pd.concat([b2cs_excel_df, ret_rows], ignore_index=True)

    # ── Pre-compute Excel bytes during Process so download is instant ──
    _sheets = {
        "COMBINED GSTR-1"      : combined_df,
        "B2B"                  : b2b_df,
        "B2CL"                 : b2cl_df,
        "B2CS"                 : b2cs_excel_df,
        "CDNR (Registered)"    : cdnr_df,
        "CDNUR (Unregistered)" : cdnur_df,
        "Stock Transfer"       : st_df,
        "Cross Charge"         : cc_df,
        "Assets Sale"          : assets_df,
        "HSN Summary"          : hsn_df,
        "Document Series"      : doc_df,
        "Cancelled Invoices"   : cancelled_df,
        "Tax Summary"          : tax_df,
    }
    _excel_bytes = write_excel(_sheets, period_label)
    _fname = f"GSTR1_Modicare_{period_label.replace(' ','_').replace('-','')}.xlsx"

    # ── Save all results + pre-built Excel to session state ──
    st.session_state["gstr1_results"] = {
        "period_label" : period_label,
        "sales_df"     : sales_df,
        "b2b_df"          : b2b_df,    "b2cl_df"        : b2cl_df,    "b2cs_df"    : b2cs_df,
        "sales_return_df" : sales_return_summary_df,
        "cdnr_df"         : cdnr_df,   "cdnur_df"       : cdnur_df,   "st_df"      : st_df,
        "cc_df"        : cc_df,     "assets_df"   : assets_df,  "hsn_df"     : hsn_df,
        "doc_df"       : doc_df,    "cancelled_df": cancelled_df,
        "tax_df"       : tax_df,    "combined_df" : combined_df,
        "excel_bytes"  : _excel_bytes,
        "excel_fname"  : _fname,
        "n_sheets"     : len(_sheets),
    }
    update_progress(100, "✅ 100% — All GSTR-1 sections generated!")
    pct_box.markdown(
        "<div style='text-align:center;font-size:40px;font-weight:bold;color:#16a34a'>100% ✅</div>",
        unsafe_allow_html=True
    )
    # Clear uploaded file objects from Streamlit widget state after processing.
    # Without this, Streamlit deepcopies the large file buffers on every rerun → OOM.
    import gc
    for _wkey in ["state_sales", "stock", "sret", "cc", "assets"]:
        st.session_state.pop(_wkey, None)
    gc.collect()

# ─────────────────────────────────────────────────────────────────
# DISPLAY — runs from session_state so widgets persist across reruns
# ─────────────────────────────────────────────────────────────────
if "gstr1_results" in st.session_state:
    r            = st.session_state["gstr1_results"]
    period_label = r["period_label"]
    sales_df     = r["sales_df"]
    b2b_df       = r["b2b_df"];    b2cl_df      = r["b2cl_df"];   b2cs_df    = r["b2cs_df"]
    cdnr_df      = r["cdnr_df"];   cdnur_df     = r["cdnur_df"];  st_df      = r["st_df"]
    cc_df        = r["cc_df"];     assets_df    = r["assets_df"]; hsn_df     = r["hsn_df"]
    doc_df       = r["doc_df"];    cancelled_df = r["cancelled_df"]
    tax_df       = r["tax_df"];    combined_df  = r["combined_df"]
    excel_bytes  = r.get("excel_bytes")
    excel_fname  = r.get("excel_fname", "GSTR1_Modicare.xlsx")
    n_sheets     = r.get("n_sheets", 13)

    # ── Merge Assets into B2B / B2CL / B2CS ──────────────────────
    if assets_df is not None and not assets_df.empty and "GSTR1_Section" in assets_df.columns:

        # Map Assets column names → GSTR-1 section column names
        _A_RENAME = {
            "Supplier State"     : "Supplier State",
            "Supplier State Code": "Supplier State Code",
            "Consignee GSTIN"    : "Receiver GSTIN",
            "Consignee State"    : "Place of Supply",
            "Invoice No"         : "Invoice No",
            "Invoice Date"       : "Invoice Date",
            "Tax Slab"           : "Tax Slab",
            "Taxable Amount"     : "Taxable_Value",
            "IGST"               : "IGST",
            "CGST"               : "CGST",
            "SGST"               : "SGST",
            "UGST"               : "UGST",
            "Cess"               : "Cess",
            "Invoice Value"      : "Invoice_Value",
            "HSN Code"           : "HSN Code",
            "Quantity"           : "Quantity",
            "UQC"                : "UQC",
            "Type"               : "Type",
        }

        def _map_assets(section):
            rows = assets_df[assets_df["GSTR1_Section"] == section].copy()
            if rows.empty: return pd.DataFrame()
            rows = rows.rename(columns=_A_RENAME)
            rows["Source"] = "Assets"
            return rows

        a_b2b  = _map_assets("B2B")
        a_b2cl = _map_assets("B2CL")
        a_b2cs = _map_assets("B2CS")

        def _safe_concat(base_df, new_df):
            """Concat two DataFrames safely — deduplicates column names first."""
            if new_df.empty: return base_df
            # Remove duplicate columns from new_df
            new_df = new_df.loc[:, ~new_df.columns.duplicated()]
            if not base_df.empty:
                base_df = base_df.loc[:, ~base_df.columns.duplicated()]
            return pd.concat([base_df, new_df], ignore_index=True)

        if not a_b2b.empty:
            # Fix POS Code for B2B Assets — derive from Receiver GSTIN first 2 digits
            if "POS Code" not in a_b2b.columns or a_b2b["POS Code"].isna().all():
                if "Receiver GSTIN" in a_b2b.columns:
                    a_b2b["POS Code"] = (a_b2b["Receiver GSTIN"].astype(str).str[:2]
                                         .apply(lambda x: int(x) if x.isdigit() else x))
                    a_b2b["Place of Supply"] = (a_b2b["Receiver GSTIN"].astype(str).str[:2]
                                                .map(GSTIN_STATE_MAP).fillna(""))
            b2b_df  = _safe_concat(b2b_df, a_b2b)
        if not a_b2cl.empty:
            b2cl_df = _safe_concat(b2cl_df, a_b2cl)

        # B2CS: fix POS Code + HSN Code, then aggregate
        if not a_b2cs.empty:
            # Fix POS Code — derive from Supplier State Code (intra-state for B2CS)
            if "POS Code" not in a_b2cs.columns or a_b2cs["POS Code"].isna().all():
                a_b2cs["POS Code"] = a_b2cs.get("Supplier State Code", "")
            if "Place of Supply" not in a_b2cs.columns or a_b2cs["Place of Supply"].isna().all():
                a_b2cs["Place of Supply"] = a_b2cs.get("Supplier State", "")

            for _c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess","Invoice_Value"]:
                if _c in a_b2cs.columns:
                    a_b2cs[_c] = pd.to_numeric(a_b2cs[_c], errors="coerce").fillna(0)

            # Include HSN Code in groupby so it's preserved in B2CS output
            _gcols = [c for c in ["Supplier State","Supplier State Code",
                                  "Place of Supply","POS Code","HSN Code","Tax Slab"]
                      if c in a_b2cs.columns]
            if _gcols:
                _agg_cols = [c for c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess","Invoice_Value"]
                             if c in a_b2cs.columns]
                a_b2cs_agg = a_b2cs.groupby(_gcols, as_index=False)[_agg_cols].sum()
                a_b2cs_agg["Total_Tax"] = (a_b2cs_agg.get("IGST",  pd.Series(0, index=a_b2cs_agg.index)) +
                                           a_b2cs_agg.get("CGST",  pd.Series(0, index=a_b2cs_agg.index)) +
                                           a_b2cs_agg.get("SGST",  pd.Series(0, index=a_b2cs_agg.index)) +
                                           a_b2cs_agg.get("UGST",  pd.Series(0, index=a_b2cs_agg.index)) +
                                           a_b2cs_agg.get("Cess",  pd.Series(0, index=a_b2cs_agg.index)))
                a_b2cs_agg["Source"] = "Assets"
                b2cs_df = _safe_concat(b2cs_df, a_b2cs_agg)

    # ── Summary metrics ───────────────────────────────────────────
    st.markdown("## 📈 GSTR-1 Summary — " + period_label)
    total_taxable = sales_df["inv_tot"].sum() if not sales_df.empty and "inv_tot" in sales_df.columns else 0
    total_igst    = sales_df["igstamt"].sum() if not sales_df.empty and "igstamt" in sales_df.columns else 0
    total_cgst    = sales_df["cgstamt"].sum() if not sales_df.empty and "cgstamt" in sales_df.columns else 0
    total_sgst    = sales_df["sgstamt"].sum() if not sales_df.empty and "sgstamt" in sales_df.columns else 0
    total_tax     = total_igst + total_cgst + total_sgst

    m1,m2,m3,m4,m5,m6 = st.columns(6)
    with m1: st.metric("B2B Invoices",   f"{len(b2b_df):,}"  if not b2b_df.empty else "0")
    with m2: st.metric("B2CS Rate Rows", f"{len(b2cs_df):,}" if not b2cs_df.empty else "0")
    with m3: st.metric("Credit Notes",   f"{len(cdnr_df)+len(cdnur_df):,}")
    with m4: st.metric("Taxable Value",  fmt_inr(total_taxable))
    with m5: st.metric("Total GST",      fmt_inr(total_tax))
    with m6: st.metric("Invoice Value",  fmt_inr(total_taxable + total_tax))

    # ── Section tabs ──────────────────────────────────────────────
    st.markdown("---")
    tabs = st.tabs([
        "📋 B2B","📋 B2CL","📋 B2CS",
        "↩️ CDNR","↩️ CDNUR",
        "🔄 Stock Transfer","⚡ Cross Charge","🏭 Assets",
        "📦 HSN Summary","📄 Document Series","💰 Tax Summary","🗂️ Combined"
    ])

    def show(tab, df, title, note="", tab_key=""):
        with tab:
            st.subheader(title)
            if note: st.caption(note)
            if df is None or (isinstance(df, pd.DataFrame) and df.empty):
                st.info("No records in this section for the uploaded data.")
            else:
                filtered = df.copy()

                # ── Tax Slab filter (shown when Tax Slab column exists) ──
                tax_col = "Tax Slab"
                if tax_col in df.columns:
                    available_rates = sorted(
                        df[tax_col].dropna().unique().tolist(),
                        key=lambda x: float(x) if str(x).replace('.','').isdigit() else 999
                    )
                    sel_rates = st.multiselect(
                        "Filter by Tax Slab",
                        options=available_rates,
                        default=available_rates,
                        key=f"tax_filter_{tab_key}",
                        format_func=lambda x: f"{x}%",
                    )
                    if sel_rates:
                        filtered = filtered[filtered[tax_col].isin(sel_rates)]

                # Coerce mixed-type object columns to string
                display_df = filtered.copy()
                for col in display_df.select_dtypes(include="object").columns:
                    display_df[col] = display_df[col].astype(str).replace("nan", "")
                st.dataframe(display_df, width='stretch', hide_index=True, height=450)
                st.caption(f"Rows: {len(filtered):,}")

    show(tabs[0],  b2b_df,    "Table 4A – B2B Invoices",              "Buyer has valid 15-digit GSTIN",              "b2b")
    show(tabs[1],  b2cl_df,   "Table 5A – B2CL Invoices",             "Inter-state, unregistered buyer, invoice ≥ ₹1,00,000", "b2cl")
    show(tabs[2],  b2cs_df,   "Table 7 – B2CS Summary",               "Consumer sales — rate-wise & state-wise summary",       "b2cs")
    # ── CDNR — detail + tax slab summary ────────────────────────
    with tabs[3]:
        st.subheader("Table 9B – Credit Notes (Registered)")
        st.caption("Returns from registered buyers")
        if cdnr_df is None or cdnr_df.empty:
            st.info("No records in this section for the uploaded data.")
        else:
            # Tax Slab filter
            tax_col = "Tax Slab"
            filtered_cdnr = cdnr_df.copy()
            if tax_col in cdnr_df.columns:
                av_rates = sorted(cdnr_df[tax_col].dropna().unique().tolist(),
                                  key=lambda x: float(x) if str(x).replace('.','').isdigit() else 999)
                sel_cdnr = st.multiselect("Filter by Tax Slab", options=av_rates,
                                          default=av_rates, key="tax_filter_cdnr",
                                          format_func=lambda x: f"{x}%")
                if sel_cdnr:
                    filtered_cdnr = filtered_cdnr[filtered_cdnr[tax_col].isin(sel_cdnr)]

            # Detail table
            disp = filtered_cdnr.copy()
            for col in disp.select_dtypes(include="object").columns:
                disp[col] = disp[col].astype(str).replace("nan", "")
            st.dataframe(disp, width="stretch", hide_index=True)
            st.caption(f"Rows: {len(filtered_cdnr):,}")

            # ── Tax Slab Summary (B2B credit note impact) ────────
            st.markdown("---")
            st.markdown("##### 📊 Tax Slab-wise Summary — B2B Credit Notes")
            num_cols = [c for c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess","Note_Value"]
                        if c in filtered_cdnr.columns]
            if tax_col in filtered_cdnr.columns and num_cols:
                summary = filtered_cdnr.groupby(tax_col, as_index=False)[num_cols].sum()
                summary[tax_col] = summary[tax_col].apply(lambda x: f"{x}%")
                # Add totals row
                totals = {tax_col: "TOTAL"}
                for c in num_cols:
                    totals[c] = summary[c].sum()
                summary = pd.concat([summary, pd.DataFrame([totals])], ignore_index=True)
                st.dataframe(summary, width="stretch", hide_index=True)

    # ── CDNUR — detail + tax slab summary ───────────────────────
    with tabs[4]:
        st.subheader("Table 9C – Credit Notes (Unregistered)")
        st.caption("Returns from unregistered buyers")
        if cdnur_df is None or cdnur_df.empty:
            st.info("No records in this section for the uploaded data.")
        else:
            tax_col = "Tax Slab"
            filtered_cdnur = cdnur_df.copy()
            if tax_col in cdnur_df.columns:
                av_rates = sorted(cdnur_df[tax_col].dropna().unique().tolist(),
                                  key=lambda x: float(x) if str(x).replace('.','').isdigit() else 999)
                sel_cdnur = st.multiselect("Filter by Tax Slab", options=av_rates,
                                           default=av_rates, key="tax_filter_cdnur",
                                           format_func=lambda x: f"{x}%")
                if sel_cdnur:
                    filtered_cdnur = filtered_cdnur[filtered_cdnur[tax_col].isin(sel_cdnur)]

            disp = filtered_cdnur.copy()
            for col in disp.select_dtypes(include="object").columns:
                disp[col] = disp[col].astype(str).replace("nan", "")
            st.dataframe(disp, width="stretch", hide_index=True)
            st.caption(f"Rows: {len(filtered_cdnur):,}")

            st.markdown("---")
            st.markdown("##### 📊 Tax Slab-wise Summary — Credit Notes (Unregistered)")
            num_cols = [c for c in ["Taxable_Value","IGST","CGST","SGST","UGST","Cess","Note_Value"]
                        if c in filtered_cdnur.columns]
            if tax_col in filtered_cdnur.columns and num_cols:
                summary = filtered_cdnur.groupby(tax_col, as_index=False)[num_cols].sum()
                summary[tax_col] = summary[tax_col].apply(lambda x: f"{x}%")
                totals = {tax_col: "TOTAL"}
                for c in num_cols:
                    totals[c] = summary[c].sum()
                summary = pd.concat([summary, pd.DataFrame([totals])], ignore_index=True)
                st.dataframe(summary, width="stretch", hide_index=True)
    show(tabs[5],  st_df,     "Table 6A – Stock Transfers",           "Warehouse-to-warehouse movements",                      "st")
    show(tabs[6],  cc_df,     "Cross Charge Invoices",                "Employee cost cross charges (included in B2B)",          "cc")
    with tabs[7]:
        st.subheader("Master Assets Sales")
        st.caption("Asset / scrap sale invoices — merged into B2B, B2CL, B2CS sections")
        if assets_df is None or assets_df.empty:
            st.info("No asset records found.")
        else:
            sec_col = "GSTR1_Section"
            n_b2b  = len(assets_df[assets_df[sec_col] == "B2B"])  if sec_col in assets_df.columns else 0
            n_b2cl = len(assets_df[assets_df[sec_col] == "B2CL"]) if sec_col in assets_df.columns else 0
            n_b2cs = len(assets_df[assets_df[sec_col] == "B2CS"]) if sec_col in assets_df.columns else 0
            amt    = assets_df["Taxable Amount"].sum() if "Taxable Amount" in assets_df.columns else 0

            c1, c2, c3, c4 = st.columns(4)
            with c1: st.metric("Total Records", f"{len(assets_df):,}")
            with c2: st.metric("Moved to B2B",  f"{n_b2b:,}")
            with c3: st.metric("Moved to B2CL", f"{n_b2cl:,}")
            with c4: st.metric("Moved to B2CS", f"{n_b2cs:,}")
            st.success(f"✅ All {len(assets_df):,} asset records (Taxable: ₹{amt:,.2f}) have been merged into B2B / B2CL / B2CS tabs.")
            st.markdown("---")
            # Show full raw data for reference
            d = assets_df.copy()
            for col in d.select_dtypes(include="object").columns:
                d[col] = d[col].astype(str).replace("nan", "")
            st.dataframe(d, width="stretch", hide_index=True)
            st.caption(f"Rows: {len(d):,}  |  Reference only — data is already in B2B/B2CL/B2CS")

    # ── HSN Summary — wrapped in @st.fragment so only this section reruns ──
    @st.fragment
    def _hsn_tab(hsn_df):
        st.subheader("Table 12 – HSN Summary")
        st.caption("State-wise & HSN-wise outward supply summary with quantity")
        if hsn_df is None or hsn_df.empty:
            st.info("No records in this section for the uploaded data.")
            return

        # Use POS column for state grouping (renamed from fstate)
        state_col  = next((c for c in ["Place of Supply (POS)","Supplier State"]
                           if c in hsn_df.columns), None)

        # Filter out invalid state names (numeric, empty, "0.0", "nan" etc.)
        if state_col:
            import re as _re
            valid_mask = (
                hsn_df[state_col].astype(str).str.strip()
                .apply(lambda s: bool(s) and not _re.match(r'^[\d\.]+$', s)
                       and s.lower() not in ["nan","none","","0","0.0"])
            )
            hsn_df = hsn_df[valid_mask].copy()

        all_states = sorted(hsn_df[state_col].dropna().unique().tolist()) if state_col else []

        sel_state = st.selectbox(
            "Select State",
            options=["All States"] + all_states,
            index=0,
            key="hsn_state_filter",
        )


        def _show_df(df):
            d = df.copy()
            for col in d.select_dtypes(include="object").columns:
                d[col] = d[col].astype(str).replace({"nan":"","NaN":"","None":"","<NA>":""})
            st.dataframe(d, width="stretch", hide_index=True)
            st.caption(f"Rows: {len(d):,}")

        def _show_hsn_state(state_data):
            type_col   = "Type"   if "Type"   in state_data.columns else None
            source_col = "Source" if "Source" in state_data.columns else None

            # Split by Type AND Source to separate Assets from Sales
            def _filter(typ=None, src=None):
                d = state_data.copy()
                if type_col and typ:
                    d = d[d[type_col] == typ]
                if source_col and src:
                    d = d[d[source_col] == src]
                return d

            b2b_data = _filter("B2B")        if type_col else pd.DataFrame()
            b2c_data = _filter("B2C")        if type_col else state_data.copy()
            cc_data  = _filter("Cross Charge") if type_col else pd.DataFrame()
            st_data  = _filter("Stock Transfer") if type_col else pd.DataFrame()
            # Assets rows — Source = "Assets" (B2B or B2C)
            asset_data = state_data[state_data[source_col] == "Assets"].copy() if source_col else pd.DataFrame()

            def _tv(d): return d["Taxable_Value"].sum() if "Taxable_Value" in d.columns else 0
            def _qty(d): return d["Total_Qty"].sum()    if "Total_Qty"     in d.columns else 0

            b2b_val   = _tv(b2b_data);   b2c_val  = _tv(b2c_data)
            cc_val    = _tv(cc_data);    st_val   = _tv(st_data)
            asset_val = _tv(asset_data); b2b_qty  = _qty(b2b_data)
            b2c_qty   = _qty(b2c_data);  st_qty   = _qty(st_data)
            asset_qty = _qty(asset_data)

            t_b2b, t_b2c, t_st, t_asset = st.tabs([
                f"🏢 B2B  |  Qty: {b2b_qty:,.0f}  |  Taxable: ₹{b2b_val:,.2f}  |  CC: ₹{cc_val:,.2f}",
                f"🛒 B2C  |  Qty: {b2c_qty:,.0f}  |  Taxable: ₹{b2c_val:,.2f}",
                f"🔄 Stock Transfer  |  Qty: {st_qty:,.0f}  |  Taxable: ₹{st_val:,.2f}",
                f"🏭 Assets  |  Qty: {asset_qty:,.0f}  |  Taxable: ₹{asset_val:,.2f}",
            ])

            # Columns to remove from non-CC rows (only CC has From State info)
            _cc_only_cols = ["From State (Supplier)", "From State Code"]

            def _clean_sales(df):
                d = df.drop(columns=[c for c in _cc_only_cols if c in df.columns], errors="ignore")
                d = d.rename(columns={"Place of Supply (POS)": "Supplier State",
                                      "POS Code": "Supplier State Code"})
                return d

            with t_b2b:
                if b2b_data.empty and cc_data.empty:
                    st.info("No B2B HSN data for this state.")
                else:
                    if not b2b_data.empty:
                        st.markdown("##### HSN Sales (B2B)")
                        _show_df(_clean_sales(b2b_data))
                    if not cc_data.empty:
                        st.markdown("##### Cross Charge (HSN 998399)")
                        _show_df(cc_data)
            with t_b2c:
                if b2c_data.empty:
                    st.info("No B2C HSN data for this state.")
                else:
                    st.markdown("##### HSN Sales (B2C)")
                    _show_df(_clean_sales(b2c_data))
            with t_asset:
                if asset_data.empty:
                    st.info("No Asset Sale HSN data for this state.")
                else:
                    st.markdown("##### Asset Sales (HSN-wise)")
                    _show_df(_clean_sales(asset_data))
            with t_st:
                if st_data.empty:
                    st.info("No Stock Transfer HSN data for this state.")
                else:
                    st.markdown("##### HSN Stock Transfer")
                    _show_df(_clean_sales(st_data))

        if sel_state != "All States" and state_col:
            filtered   = hsn_df[hsn_df[state_col] == sel_state].copy()
            state_code = filtered["Supplier State Code"].iloc[0] if "Supplier State Code" in filtered.columns and not filtered.empty else ""
            st.markdown(f"#### 🏭 {sel_state} &nbsp;&nbsp; `Code: {state_code}`")
            if filtered.empty:
                st.warning(f"⚠️ No HSN data present for **{sel_state}** in this period.")
            else:
                _show_hsn_state(filtered)
            base = filtered
        else:
            if not all_states:
                st.warning("⚠️ No HSN data present for any state in this period.")
                base = hsn_df
            else:
                for state in all_states:
                    state_data = hsn_df[hsn_df[state_col] == state].copy() if state_col else hsn_df.copy()
                    state_code = state_data["Supplier State Code"].iloc[0] if "Supplier State Code" in state_data.columns and not state_data.empty else ""
                    total_qty  = state_data["Total_Qty"].sum()     if "Total_Qty"     in state_data.columns else 0
                    total_val  = state_data["Taxable_Value"].sum() if "Taxable_Value" in state_data.columns else 0
                    # CC value from aggregated hsn_df rows (Type="Cross Charge")
                    cc_rows = state_data[state_data["Type"] == "Cross Charge"] if "Type" in state_data.columns else pd.DataFrame()
                    cc_val  = cc_rows["Taxable_Value"].sum() if "Taxable_Value" in cc_rows.columns else 0
                    with st.expander(
                        f"🏭 {state}  (Code: {state_code})  |  Qty: {total_qty:,.0f}  |  Taxable: ₹{total_val:,.2f}  |  CC: ₹{cc_val:,.2f}",
                        expanded=False,
                    ):
                        _show_hsn_state(state_data)
                base = hsn_df

        st.markdown("---")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.metric("States",        len(all_states) if sel_state == "All States" else 1)
        with c2: st.metric("HSN Rows",      f"{len(base):,}")
        with c3: st.metric("Total Qty",     f"{base['Total_Qty'].sum():,.0f}"    if "Total_Qty"     in base.columns else "-")
        with c4: st.metric("Total Taxable", fmt_inr(base["Taxable_Value"].sum()) if "Taxable_Value" in base.columns else "-")

    with tabs[8]:
        _hsn_tab(hsn_df)

    with tabs[9]:
        st.subheader("Table 13 – Document Series")
        st.caption("Series-wise invoice count — gaps in sequence are treated as Cancelled")
        if doc_df is None or doc_df.empty:
            st.info("No records in this section for the uploaded data.")
        else:
            display_doc = doc_df.copy()
            for col in display_doc.select_dtypes(include="object").columns:
                display_doc[col] = display_doc[col].astype(str).replace("nan", "")
            st.dataframe(display_doc, width="stretch", hide_index=True)
            total_cancelled = int(doc_df["Cancelled"].sum()) if "Cancelled" in doc_df.columns else 0
            st.caption(f"Series: {len(doc_df):,}  |  Total Cancelled (gaps): {total_cancelled:,}")
        if cancelled_df is not None and not cancelled_df.empty:
            with st.expander(f"🚫 Cancelled Invoice Detail — {len(cancelled_df):,} missing invoice numbers", expanded=False):
                display_c = cancelled_df.copy()
                for col in display_c.select_dtypes(include="object").columns:
                    display_c[col] = display_c[col].astype(str).replace("nan", "")
                st.dataframe(display_c, width="stretch", hide_index=True)
        else:
            st.success("✅ No gaps found — all invoice series are complete.")

    show(tabs[10], tax_df,     "Tax Liability Summary",        "Rate-wise tax across all sections",                            "tax")
    # Combined GSTR-1 — limit display to avoid protobuf serialization error on large DataFrames
    with tabs[11]:
        st.subheader("🗂️ Combined GSTR-1 Master")
        st.caption("All invoices in one sheet — use for final GSTR-1 filing review")
        if combined_df is None or (isinstance(combined_df, pd.DataFrame) and combined_df.empty):
            st.info("No records in this section for the uploaded data.")
        else:
            _DISPLAY_LIMIT = 10000
            _total_rows = len(combined_df)
            _disp = combined_df.head(_DISPLAY_LIMIT).copy()
            for col in _disp.select_dtypes(include="object").columns:
                _disp[col] = _disp[col].astype(str).replace("nan", "")
            if _total_rows > _DISPLAY_LIMIT:
                st.warning(f"Showing first {_DISPLAY_LIMIT:,} of {_total_rows:,} rows. Download the Excel file below for the complete data.")
            st.dataframe(_disp, width='stretch', hide_index=True, height=450)
            st.caption(f"Rows displayed: {min(_DISPLAY_LIMIT, _total_rows):,} / {_total_rows:,}")

    # ── Combined GSTR-1 dedicated download button (always visible inside tab) ──
    with tabs[11]:
        st.markdown("---")
        if combined_df is not None and not combined_df.empty:
            _nan_s = ["nan","NaN","None","none","<NA>"]
            _fn = f"Combined_GSTR1_{period_label}.xlsx"
            _tmp2 = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            _tmp2.close()
            try:
                with xlsxwriter.Workbook(_tmp2.name, {"strings_to_numbers": False}) as _wb:
                    _ws  = _wb.add_worksheet("Combined GSTR-1")
                    _hf  = _wb.add_format({"bold":True,"bg_color":"#1e3a5f","font_color":"white","border":1})
                    _nf  = _wb.add_format({"num_format":"#,##0.00","border":1})
                    _tf  = _wb.add_format({"border":1})
                    _ttf = _wb.add_format({"bold":True,"font_size":12,"bg_color":"#d6e4f0","border":1})
                    _cls = list(combined_df.columns)
                    _ws.merge_range(0,0,0,max(len(_cls)-1,0),f"Combined GSTR-1  |  Period: {period_label}",_ttf)
                    for _ci,_c in enumerate(_cls):
                        _ws.write(1,_ci,_c,_hf)
                        _ws.set_column(_ci,_ci,max(16,len(str(_c))+4))
                    _ncols = set(combined_df.select_dtypes(include=[np.number]).columns)
                    for _ci,_c in enumerate(_cls):
                        _s = combined_df.iloc[:,_ci]
                        if _c in _ncols:
                            _ws.write_column(2,_ci,pd.to_numeric(_s,errors="coerce").fillna(0).tolist(),_nf)
                        else:
                            _ws.write_column(2,_ci,_s.astype(object).where(_s.notna(),"").astype(str).str.strip().replace(_nan_s,"").tolist(),_tf)
                with open(_tmp2.name, "rb") as _f:
                    _combined_bytes = _f.read()
            finally:
                os.unlink(_tmp2.name)
            st.download_button(
                label="📥 Download Combined GSTR-1 Excel",
                data=_combined_bytes,
                file_name=_fn,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.caption(f"Rows: {len(combined_df):,} | Columns: {len(combined_df.columns)}")
        else:
            st.warning("⚠️ Combined GSTR-1 has no data. Re-process to enable download.")

    # ── DOWNLOAD — HTML anchor link (no page refresh) ────────────
    st.markdown("---")
    st.subheader("⬇️ Download GSTR-1 Output Files")
    dc1, dc2 = st.columns([1, 2])
    with dc1:
        if excel_bytes:
            if st.download_button(
                label="📥 Download GSTR-1 Excel (All Sections)",
                data=excel_bytes,
                file_name=excel_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            ):
                log_download(get_current_user().get("username", ""), excel_fname)
        else:
            st.warning("Re-process data to enable download.")
    with dc2:
        cancelled_count = len(cancelled_df) if cancelled_df is not None and not cancelled_df.empty else 0
        st.success(f"✅ **{excel_fname}** ready — {n_sheets} sheets including **COMBINED GSTR-1** + **Cancelled Invoices** ({cancelled_count:,} gaps).")

st.markdown("---")
# st.caption("ASC Consulting Pvt. Ltd. | Modicare GSTR-1 Dashboard v2.0 | Internal use only")

