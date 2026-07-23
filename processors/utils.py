"""Shared utilities for ITC reconciliation."""
import re
import pandas as pd
import numpy as np

_MONTH_ABBR = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

def normalize_gstin(val):
    if pd.isna(val):
        return ""
    return str(val).strip().upper().replace(" ", "")


def normalize_invoice(val):
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    s = re.sub(r"\s+", "", s)
    # Remove common separators for matching but keep a canonical form
    return s


def normalize_invoice_for_match(val):
    """Aggressive normalization for cross-source invoice matching."""
    s = normalize_invoice(val)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def make_con_key(gstin, invoice_no):
    return normalize_gstin(gstin) + normalize_invoice(invoice_no)


def make_con_key_match(gstin, invoice_no):
    """CON key used for reconciliation matching (strips special chars from invoice)."""
    return normalize_gstin(gstin) + normalize_invoice_for_match(invoice_no)


def validate_gstin(gstin):
    g = normalize_gstin(gstin)
    if not g:
        return False, "Missing GSTIN"
    if len(g) != 15:
        return False, f"GSTIN length {len(g)} (expected 15)"
    if not re.match(r"^[0-9]{2}[A-Z0-9]{13}$", g):
        return False, "Invalid GSTIN format"
    return True, ""


def validate_invoice(inv):
    inv = normalize_invoice(inv)
    if not inv:
        return False, "Missing Invoice No"
    if len(inv) > 16:
        return False, f"Invoice length {len(inv)} (max 16)"
    return True, ""


def extract_gstin_from_narration(narration):
    if pd.isna(narration):
        return ""
    matches = re.findall(r"[0-9]{2}[A-Z0-9]{13}", str(narration).upper())
    return matches[0] if matches else ""


def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


def sum_tax_columns(df, prefixes=("Input IGST", "Input CGST", "Input SGST", "Input Cess")):
    igst = cgst = sgst = cess = 0.0
    for col in df.columns:
        col_str = str(col)
        vals = safe_numeric(df[col])
        if "IGST" in col_str and "Input" in col_str:
            igst += vals
        elif "CGST" in col_str and "Input" in col_str:
            cgst += vals
        elif "SGST" in col_str and "Input" in col_str:
            sgst += vals
        elif "Cess" in col_str and "Input" in col_str:
            cess += vals
    return igst, cgst, sgst, cess


def tax_match(books_tax, g2b_tax, tolerance=1.0):
    diff = abs(float(books_tax or 0) - float(g2b_tax or 0))
    return diff <= tolerance, diff


def total_itc(row, igst_col="IGST", cgst_col="CGST", sgst_col="SGST", cess_col="Cess"):
    igst = float(row.get(igst_col, 0) or 0)
    cgst = float(row.get(cgst_col, 0) or 0)
    sgst = float(row.get(sgst_col, 0) or 0)
    cess = float(row.get(cess_col, 0) or 0)
    return igst + cgst + sgst + cess


def gstin_state_code(gstin):
    g = normalize_gstin(gstin)
    return g[:2] if len(g) >= 2 else ""


GST_STATE_CODE_TO_NAME = {
    "01": "Jammu and Kashmir",
    "02": "Himachal Pradesh",
    "03": "Punjab",
    "04": "Chandigarh",
    "05": "Uttarakhand",
    "06": "Haryana",
    "07": "Delhi",
    "08": "Rajasthan",
    "09": "Uttar Pradesh",
    "10": "Bihar",
    "11": "Sikkim",
    "12": "Arunachal Pradesh",
    "13": "Nagaland",
    "14": "Manipur",
    "15": "Mizoram",
    "16": "Tripura",
    "17": "Meghalaya",
    "18": "Assam",
    "19": "West Bengal",
    "20": "Jharkhand",
    "21": "Odisha",
    "22": "Chhattisgarh",
    "23": "Madhya Pradesh",
    "24": "Gujarat",
    "25": "Daman and Diu",
    "26": "Dadra and Nagar Haveli",
    "27": "Maharashtra",
    "28": "Andhra Pradesh",
    "29": "Karnataka",
    "30": "Goa",
    "31": "Lakshadweep",
    "32": "Kerala",
    "33": "Tamil Nadu",
    "34": "Puducherry",
    "35": "Andaman and Nicobar Islands",
    "36": "Telangana",
    "37": "Andhra Pradesh",
    "38": "Ladakh",
    "97": "Other Territory",
    "99": "Centre Jurisdiction",
}


VOUCHER_STATE_ALIASES = {
    "ANDHRA PRADESH": "Andhra Pradesh",
    "ARUNACHAL PRADESH": "Arunachal Pradesh",
    "ASSAM": "Assam",
    "BIHAR": "Bihar",
    "CHANDIGARH": "Chandigarh",
    "CHHATTISGARH": "Chhattisgarh",
    "DELHI": "Delhi",
    "GOA": "Goa",
    "GUJARAT": "Gujarat",
    "HARYANA": "Haryana",
    "HIMACHAL PRADESH": "Himachal Pradesh",
    "HIMACHAL": "Himachal Pradesh",
    "JHARKHAND": "Jharkhand",
    "KARNATAKA": "Karnataka",
    "KERALA": "Kerala",
    "LADAKH": "Ladakh",
    "MADHYA PRADESH": "Madhya Pradesh",
    "MADHYA": "Madhya Pradesh",
    "MAHARASHTRA": "Maharashtra",
    "MANIPUR": "Manipur",
    "MEGHALAYA": "Meghalaya",
    "MIZORAM": "Mizoram",
    "NAGALAND": "Nagaland",
    "ODISHA": "Odisha",
    "PUNJAB": "Punjab",
    "RAJASTHAN": "Rajasthan",
    "SIKKIM": "Sikkim",
    "TAMIL NADU": "Tamil Nadu",
    "TAMILNADU": "Tamil Nadu",
    "TELANGANA": "Telangana",
    "TRIPURA": "Tripura",
    "UTTAR PRADESH": "Uttar Pradesh",
    "UTTAR": "Uttar Pradesh",
    "UTTARAKHAND": "Uttarakhand",
    "WEST BENGAL": "West Bengal",
    "WESTBENGAL": "West Bengal",
    "MP": "Madhya Pradesh",
    "UP": "Uttar Pradesh",
    "AP": "Andhra Pradesh",
    "BI": "Bihar",
    "HI": "Himachal Pradesh",
    "MA": "Manipur",
    "MI": "Mizoram",
    "TR": "Tripura",
}


def state_name_from_code(code):
    c = str(code or "").strip()
    if not c:
        return ""
    if c.isdigit():
        c = c.zfill(2)[:2]
    return GST_STATE_CODE_TO_NAME.get(c, "")


def state_name_from_gstin(gstin):
    return state_name_from_code(gstin_state_code(gstin))


def state_name_from_voucher_type(voucher_type):
    """Derive state name from JV-* voucher types and common aliases."""
    if pd.isna(voucher_type):
        return ""
    vt = str(voucher_type).strip().upper()
    if not vt:
        return ""
    if vt.startswith("JV-"):
        body = vt[3:]
        if body.endswith("-PUR"):
            body = body[:-4]
        elif "-PUR" in body:
            body = body.split("-PUR", 1)[0]
        body = body.strip()
        if body in VOUCHER_STATE_ALIASES:
            return VOUCHER_STATE_ALIASES[body]
        for key, val in VOUCHER_STATE_ALIASES.items():
            if key in body:
                return val
        return ""
    for key, val in VOUCHER_STATE_ALIASES.items():
        if key in vt:
            return val
    return ""


def validate_gst_head(gstin, ship_state_code, igst, cgst, sgst, tolerance=1.0):
    """Check IGST for inter-state vs CGST+SGST for intra-state supply."""
    igst, cgst, sgst = float(igst or 0), float(cgst or 0), float(sgst or 0)
    sup_state = gstin_state_code(gstin)
    ship = str(ship_state_code or "").strip().zfill(2) if ship_state_code else ""
    if not sup_state or not ship:
        return True, ""
    is_intra = sup_state == ship
    if is_intra:
        if igst > tolerance and (cgst > tolerance or sgst > tolerance):
            return False, "IGST should not apply for intra-state supply"
        if (cgst > tolerance or sgst > tolerance) and abs(cgst - sgst) > tolerance:
            return False, "CGST/SGST mismatch for intra-state"
    else:
        if (cgst > tolerance or sgst > tolerance) and igst <= tolerance:
            return False, "CGST/SGST should not apply for inter-state supply"
        if igst <= tolerance and (cgst > tolerance or sgst > tolerance):
            return False, "IGST missing for inter-state supply"
    return True, ""


def sum_expense_tax_columns(df):
    """Sum tax from Total Input columns (AP/AQ/AR) or Input columns."""
    igst = cgst = sgst = cess = pd.Series(0.0, index=df.index)
    if "Total Input IGST" in df.columns:
        igst = safe_numeric(df["Total Input IGST"])
    else:
        for col in df.columns:
            if "Input" in str(col) and "IGST" in str(col):
                igst += safe_numeric(df[col])
    if "Total Input CGST" in df.columns:
        cgst = safe_numeric(df["Total Input CGST"])
    else:
        for col in df.columns:
            if "Input" in str(col) and "CGST" in str(col):
                cgst += safe_numeric(df[col])
    if "Total Input SGST" in df.columns:
        sgst = safe_numeric(df["Total Input SGST"])
    else:
        for col in df.columns:
            if "Input" in str(col) and "SGST" in str(col):
                sgst += safe_numeric(df[col])
    return igst, cgst, sgst, cess


def is_ineligible_debit(rate_val):
    r = str(rate_val or "").lower()
    return "ineligible" in r and "debit" in r


def is_ineligible_credit(rate_val):
    r = str(rate_val or "").lower()
    return "ineligible" in r and "credit" in r


def format_data_received_month(value):
    """Format processing month as Mon-YY (e.g. May-26)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, str):
        s = value.strip()
        if not s or s.lower() in ("nan", "none"):
            return ""
        m = re.match(r"^([A-Za-z]{3})-(\d{2})$", s)
        if m:
            return f"{m.group(1).capitalize()}-{m.group(2)}"
    try:
        if hasattr(value, "to_timestamp"):
            dt = value.to_timestamp()
        else:
            dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return ""
        return f"{_MONTH_ABBR[dt.month - 1]}-{dt.strftime('%y')}"
    except Exception:
        return ""


def _month_from_filename(name):
    """Parse MMYYYY or month name from uploaded file name."""
    if not name:
        return None
    text = str(name)
    m = re.search(r"(?<!\d)(0[1-9]|1[0-2])(20\d{2})(?!\d)", text)
    if m:
        return pd.Timestamp(year=int(m.group(2)), month=int(m.group(1)), day=1)
    m = re.search(
        r"\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
        r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
        r"[\s.\-]*(?:'?)?(\d{2}|\d{4})\b",
        text,
        re.I,
    )
    if m:
        month_token = m.group(1)[:3].capitalize()
        if month_token in _MONTH_ABBR:
            year = int(m.group(2))
            if year < 100:
                year += 2000
            return pd.Timestamp(year=year, month=_MONTH_ABBR.index(month_token) + 1, day=1)
    return None


def _collect_processing_dates(mrr_result=None, expense_result=None):
    dates = []
    if mrr_result:
        for key in ("pivot", "combined"):
            df = mrr_result.get(key)
            if df is None or df.empty:
                continue
            for col in ("Invoice Date", "sup_dt"):
                if col in df.columns:
                    dates.extend(df[col].tolist())
    if expense_result:
        df = expense_result.get("all_expenses")
        if df is not None and not df.empty:
            for col in ("Voucher Date", "Bill Date"):
                if col in df.columns:
                    dates.extend(df[col].tolist())
    return dates


def infer_data_received_month(
    month_filter=None,
    files=None,
    mrr_result=None,
    expense_result=None,
):
    """Derive DATA RECEIVED IN MONTH label (Mon-YY) for Step 3 books sheets."""
    if month_filter:
        label = format_data_received_month(month_filter)
        if label:
            return label

    files = files or {}
    for key in ("mrr", "mrr_return", "expenses", "books"):
        file_obj = files.get(key)
        name = getattr(file_obj, "name", "") if file_obj is not None else ""
        dt = _month_from_filename(name)
        if dt is not None:
            return format_data_received_month(dt)

    raw_dates = _collect_processing_dates(mrr_result, expense_result)
    if raw_dates:
        parsed = pd.to_datetime(raw_dates, dayfirst=True, errors="coerce")
        valid = parsed.dropna()
        if len(valid):
            mode = valid.dt.to_period("M").mode()
            if len(mode):
                return format_data_received_month(mode.iloc[0])

    return ""
