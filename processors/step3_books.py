"""Step 3: Post Step 1 + Step 2 data into Books — Eligible ITC + Ineligible ITC."""
import pandas as pd
from .utils import (
    safe_numeric,
    is_ineligible_debit,
    state_name_from_gstin,
    state_name_from_code,
    state_name_from_voucher_type,
)


def _resolve_state(row, ship_state=None):
    """State for Books: voucher type, then ship state, then GSTIN state code."""
    state = state_name_from_voucher_type(row.get("Voucher Type", ""))
    if state:
        return state
    if ship_state is not None and not pd.isna(ship_state):
        ship = str(ship_state).strip()
        if ship and ship.lower() not in ("nan", "none"):
            if ship.isdigit():
                return state_name_from_code(ship) or ship
            return ship
    for col in ("GSTIN/UIN", "Supplier GSTIN", "GSTIN"):
        gstin = row.get(col, "")
        if not pd.isna(gstin) and str(gstin).strip():
            state = state_name_from_gstin(gstin)
            if state:
                return state
    return ""


def _init_remark_columns(books):
    for col in [
        "Remarks", "REMARKS1", "add remarks",
        "DATA RECEIVED IN MONTH", "transaction id", "StateBOOKS",
    ]:
        if col not in books.columns:
            books[col] = ""
    return books


def _format_books_date(val):
    """Format Date_Books as dd/mm/yyyy."""
    if pd.isna(val) or str(val).strip().lower() in ("", "nan", "none"):
        return ""
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return val
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return val


def _aggregate_books(books_rows, group_keys=None):
    books = pd.DataFrame(books_rows)
    if books.empty:
        return books

    for col in ["SGST BOOKS", "CGST BOOKS", "IGST BOOKS", "Cess tax"]:
        if col in books.columns:
            books[col] = safe_numeric(books[col])

    if group_keys is None:
        books["Total ITC As per Books"] = (
            books["SGST BOOKS"] + books["CGST BOOKS"] +
            books["IGST BOOKS"] + books.get("Cess tax", 0)
        )
        return _init_remark_columns(books)

    text_cols = [
        "Category", "State", "StateBOOKS", "Supplier name", "GSTIN", "Invoice No.",
        "Voucher No.", "Vch No.", "Rate", "Rate Label", "Source", "Voucher Type", "Narration",
        "Vch Date", "Date_Books", "Entered By", "Voucher Ref.", "Zone",
        "Remarks", "REMARKS1", "add remarks", "DATA RECEIVED IN MONTH", "transaction id",
    ]
    text_cols = [c for c in text_cols if c in books.columns]
    agg_map = {c: "sum" for c in ["SGST BOOKS", "CGST BOOKS", "IGST BOOKS", "Cess tax"]}
    for c in text_cols:
        agg_map[c] = "first"
    group_keys = list(group_keys or ["CON"])
    for c in books.columns:
        if c not in agg_map and c not in group_keys:
            agg_map[c] = "first"
    books = books.groupby(group_keys, as_index=False).agg(agg_map)
    books["Total ITC As per Books"] = (
        books["SGST BOOKS"] + books["CGST BOOKS"] +
        books["IGST BOOKS"] + books.get("Cess tax", 0)
    )
    return _init_remark_columns(books)


ELIGIBLE_EXPORT_COLS = [
    "CON", "S.NO.", "Category", "ZONE", "State", "Voucher Type", "Vch Date",
    "Vch No.", "Voucher Ref./ Original Invoice No. as per books", "Narration",
    "Entered By", "Supplier name", "GSTIN", "Invoice No.", "Date_Books", "Rate",
    "SGST BOOKS", "CGST BOOKS", "IGST BOOKS", "Cess tax", "Total ITC As per Books",
    "StateBOOKS", "Remarks", "REMARKS1", "add remarks", "DATA RECEIVED IN MONTH",
    "MONTH OF ITC TAKEN", "Reclaim", "transaction id",
]

INELIGIBLE_EXPORT_COLS = [
    "CON", "Zone", "State Receiver", "Voucher Type", "Vch Date", "Vch No.",
    "Voucher Ref.", "Narration", "Entered By", "Supplier name", "GSTIN",
    "InvoiceNO_Books", "Date_Books", "SGST BOOKS", "CGST BOOKS", "IGST BOOKS",
    "Cess tax", "Total ITC As per Books", "State BOOKS", "Remarks", "Remarks 1",
    "add remarks", "DATA RECEIVED IN MONTH", "transaction id",
]


def _is_blank_date(val):
    return pd.isna(val) or str(val).strip().lower() in ("", "nan", "none")


def _resolve_bill_date(row):
    """Bill Date from master expenses; fall back to Voucher Date when Bill Date is blank."""
    for col in row.index:
        low = str(col).strip().lower().replace(" ", "")
        if low in ("billdate", "billdt"):
            val = row.get(col, "")
            if not _is_blank_date(val):
                return val
    if "Bill Date" in row.index:
        val = row.get("Bill Date", "")
        if not _is_blank_date(val):
            return val
    for col in ("Voucher Date", "Vch Date", "Invoice Date"):
        if col in row.index:
            val = row.get(col, "")
            if not _is_blank_date(val):
                return val
    return ""


def format_eligible_books(books):
    """Align Eligible ITC output to Books Final (25-26 sheet) column layout."""
    if books is None or books.empty:
        return pd.DataFrame(columns=ELIGIBLE_EXPORT_COLS)
    out = books.copy()
    out["Vch No."] = out["Voucher No."] if "Voucher No." in out.columns else out.get("Vch No.", "")
    out["Voucher Ref./ Original Invoice No. as per books"] = (
        out["Voucher Ref."] if "Voucher Ref." in out.columns else ""
    )
    if "StateBOOKS" not in out.columns:
        out["StateBOOKS"] = out["State"] if "State" in out.columns else ""
    if "Date_Books" in out.columns:
        out["Date_Books"] = out["Date_Books"].apply(_format_books_date)
    elif "Bill Date" in out.columns:
        out["Date_Books"] = out["Bill Date"].apply(_format_books_date)
    elif "Invoice Date" in out.columns:
        out["Date_Books"] = out["Invoice Date"].apply(_format_books_date)
    else:
        out["Date_Books"] = ""
    if "Vch Date" in out.columns:
        out["Vch Date"] = out["Vch Date"].apply(_format_books_date)
    for col in ELIGIBLE_EXPORT_COLS:
        if col not in out.columns:
            out[col] = ""
    return out[ELIGIBLE_EXPORT_COLS]


S2_COMBINED_REMARKS = frozenset({"Expense", "Ineligible", "Purchase"})

S3_COMBINED_SOURCES = frozenset({"Expense", "MRR", "MRR Return"})
S3_COMBINED_CATEGORIES = frozenset({"Expense", "Purchase", "Purchase Return"})
S3_INELIGIBLE_SOURCE = "Expense-Ineligible"

BOOKS_COMBINED_CATEGORIES = {
    "Expense", "Ineligible", "Purchase", "Purchase Return",
}

S3_SOURCE_CATEGORY = {
    "Expense": "Expense",
    "MRR": "Purchase",
    "MRR Return": "Purchase Return",
}


def _category_for_source(source):
    return S3_SOURCE_CATEGORY.get(str(source or "").strip(), "")


def _enforce_combined_source_category(df):
    """S3_Books_Combined: Source and Category always align (Expense→Expense, MRR→Purchase, etc.)."""
    if df is None or df.empty or "Source" not in df.columns:
        return df
    out = df.copy()
    if "Category" not in out.columns:
        out["Category"] = ""
    for source, category in S3_SOURCE_CATEGORY.items():
        mask = out["Source"].astype(str).str.strip() == source
        if mask.any():
            out.loc[mask, "Category"] = category
    return out


ALLOWED_INELIGIBLE_RATE_LABELS = frozenset({
    "input igst-ineligible debit",
    "input gst-ineligible debit",
})


def _normalize_rate_label(val):
    return str(val or "").strip().lower()


def _is_allowed_ineligible_rate_label(val):
    return is_ineligible_debit(val)


def _normalize_asc_remark(val):
    if pd.isna(val):
        return ""
    s = str(val or "").strip().lower()
    if s in ("", "nan", "none"):
        return ""
    if s in ("expense", "expesne") or (s.startswith("expes") and "inelig" not in s):
        return "Expense"
    if s in ("ineligible", "ineligble") or "inelig" in s:
        return "Ineligible"
    if s == "purchase":
        return "Purchase"
    return str(val or "").strip()


def _asc_remark_value(row):
    """Use Asc Remarks from master sheet; fall back to computed Remark when blank."""
    for col in row.index:
        if str(col).strip().lower().replace(" ", "") in ("ascremarks", "ascremark"):
            val = row.get(col, "")
            if not pd.isna(val) and str(val).strip().lower() not in ("", "nan", "none"):
                return val
    if "Asc Remarks" in row.index:
        val = row.get("Asc Remarks", "")
        if not pd.isna(val) and str(val).strip().lower() not in ("", "nan", "none"):
            return val
    return row.get("Remark", "")


def _is_asc_expense(val):
    return _normalize_asc_remark(val) == "Expense"


def _is_asc_purchase(val):
    return _normalize_asc_remark(val) == "Purchase"


def _is_asc_ineligible(val):
    return _normalize_asc_remark(val) == "Ineligible"


def _is_blank_field(val):
    return pd.isna(val) or str(val).strip().lower() in ("", "nan", "none")


def _fill_data_received_month(df, data_month):
    """Ensure every export row has DATA RECEIVED IN MONTH (e.g. May-26)."""
    if df is None or df.empty or not str(data_month or "").strip():
        return df
    out = df.copy()
    col = "DATA RECEIVED IN MONTH"
    if col not in out.columns:
        out[col] = data_month
        return out
    blank = out[col].apply(_is_blank_field)
    if blank.any():
        out.loc[blank, col] = data_month
    return out


def _skip_incomplete_s2_row(row):
    """Skip master expense total/summary lines with no invoice or voucher identity."""
    if not _is_blank_field(row.get("con", "")):
        return False
    gstin = row.get("GSTIN/UIN", "")
    bill_no = row.get("Bill No", "")
    voucher_type = row.get("Voucher Type", "")
    if not _is_blank_field(gstin) or not _is_blank_field(bill_no):
        return False
    return _is_blank_field(voucher_type)


def _combined_s2_expense_rows(all_expenses, data_month, mrr_cons=None):
    """S2 rows for Combined — Master sheet Asc Remarks = Expense (incl. Expesne typo)."""
    if all_expenses is None or all_expenses.empty:
        return []
    mrr_cons = mrr_cons or set()
    rows = []
    for _, row in all_expenses.iterrows():
        if _skip_incomplete_s2_row(row):
            continue
        asc = _asc_remark_value(row)
        if not _is_asc_expense(asc):
            continue
        if row["con"] in mrr_cons:
            continue
        rows.append(_row_from_s2_combined(row, data_month, source="Expense"))
    return rows


def _combined_s2_purchase_rows(all_expenses, data_month, mrr_cons=None):
    """S2 Purchase rows — Asc Remarks = Purchase, Source stays Expense (master sheet)."""
    if all_expenses is None or all_expenses.empty:
        return []
    rows = []
    for _, row in all_expenses.iterrows():
        if _skip_incomplete_s2_row(row):
            continue
        asc = _asc_remark_value(row)
        if not _is_asc_purchase(asc):
            continue
        rows.append(_row_from_s2_combined(row, data_month, source="Expense"))
    return rows


def _combined_s2_ineligible_rows(all_expenses, data_month):
    """Ineligible export — Asc Remarks = Ineligible and Rate Label is debit only."""
    if all_expenses is None or all_expenses.empty:
        return []
    rows = []
    for _, row in all_expenses.iterrows():
        asc = _asc_remark_value(row)
        if not _is_asc_ineligible(asc):
            continue
        if not is_ineligible_debit(row.get("Rate Label", "")):
            continue
        rows.append(_row_from_s2_ineligible(row, data_month))
    return rows


def _row_from_s2_combined(row, data_month, source):
    remark = _normalize_asc_remark(_asc_remark_value(row)) or str(row.get("Remark", ""))
    category = _category_for_source(source) or remark
    state = _resolve_state(row)
    bill_date = _resolve_bill_date(row)
    return {
        "CON": row["con"],
        "Category": category,
        "State": state,
        "StateBOOKS": state,
        "Supplier name": row.get("Particulars", ""),
        "GSTIN": row.get("GSTIN/UIN", ""),
        "Invoice No.": row.get("Bill No", ""),
        "Voucher No.": row.get("Voucher No.", ""),
        "Vch Date": row.get("Voucher Date", ""),
        "Date_Books": _format_books_date(bill_date),
        "Voucher Type": row.get("Voucher Type", ""),
        "Narration": row.get("Voucher Narration", ""),
        "Voucher Ref.": row.get("Voucher Ref.", ""),
        "Entered By": row.get("Created by", ""),
        "Rate": row.get("Tax Rate", row.get("Rate Label", "")),
        "Rate Label": row.get("Rate Label", ""),
        "SGST BOOKS": float(row.get("SGST", 0) or 0),
        "CGST BOOKS": float(row.get("CGST", 0) or 0),
        "IGST BOOKS": float(row.get("IGST", 0) or 0),
        "Cess tax": float(row.get("Cess", 0) or 0),
        "Remarks": remark,
        "Source": source,
        "DATA RECEIVED IN MONTH": data_month,
    }


def _row_from_s2_eligible(row, data_month):
    return _row_from_s2_combined(row, data_month, source="Expense")


def _row_from_s2_ineligible(row, data_month):
    state = _resolve_state(row)
    return {
        "CON": row["con"],
        "Category": "Ineligible",
        "State": state,
        "StateBOOKS": state,
        "Supplier name": row.get("Particulars", ""),
        "GSTIN": row.get("GSTIN/UIN", ""),
        "Invoice No.": row.get("Bill No", ""),
        "Voucher No.": row.get("Voucher No.", ""),
        "Vch Date": row.get("Voucher Date", ""),
        "Date_Books": _format_books_date(_resolve_bill_date(row)),
        "Voucher Type": row.get("Voucher Type", ""),
        "Narration": row.get("Voucher Narration", ""),
        "Voucher Ref.": row.get("Voucher Ref.", ""),
        "Entered By": row.get("Created by", ""),
        "Rate": row.get("Rate Label", ""),
        "Rate Label": row.get("Rate Label", ""),
        "SGST BOOKS": float(row.get("SGST", 0) or 0),
        "CGST BOOKS": float(row.get("CGST", 0) or 0),
        "IGST BOOKS": float(row.get("IGST", 0) or 0),
        "Cess tax": float(row.get("Cess", 0) or 0),
        "Remarks": "Ineligible",
        "Source": "Expense-Ineligible",
        "DATA RECEIVED IN MONTH": data_month,
    }


def _row_from_mrr_purchase(row, data_month):
    ship_state = row.get("Ship State", "")
    state = _resolve_state(row, ship_state=ship_state)
    return {
        "CON": row["con"],
        "Category": "Purchase",
        "State": state,
        "StateBOOKS": state,
        "Supplier name": row.get("Supplier Name", ""),
        "GSTIN": row.get("Supplier GSTIN", ""),
        "Invoice No.": row.get("Invoice No.", ""),
        "Voucher No.": row.get("Voucher No.", ""),
        "Vch Date": _format_books_date(row.get("Voucher Date", "")),
        "Date_Books": _format_books_date(row.get("Invoice Date", "")),
        "Rate": row.get("Tax Slab", ""),
        "SGST BOOKS": float(row.get("SGST", 0) or 0),
        "CGST BOOKS": float(row.get("CGST", 0) or 0),
        "IGST BOOKS": float(row.get("IGST", 0) or 0),
        "Cess tax": float(row.get("Cess", 0) or 0),
        "Remarks": "MRR",
        "Source": "MRR",
        "DATA RECEIVED IN MONTH": data_month,
    }


def _row_from_mrr_purchase_return(row, data_month):
    ship_state = row.get("Ship State", "")
    state = _resolve_state(row, ship_state=ship_state)
    return {
        "CON": row["con"],
        "Category": "Purchase Return",
        "State": state,
        "StateBOOKS": state,
        "Supplier name": row.get("Supplier Name", ""),
        "GSTIN": row.get("Supplier GSTIN", ""),
        "Invoice No.": row.get("Invoice No.", ""),
        "Voucher No.": row.get("Voucher No.", ""),
        "Vch Date": _format_books_date(row.get("Voucher Date", "")),
        "Date_Books": _format_books_date(row.get("Invoice Date", "")),
        "Rate": row.get("Tax Slab", ""),
        "SGST BOOKS": float(row.get("SGST", 0) or 0),
        "CGST BOOKS": float(row.get("CGST", 0) or 0),
        "IGST BOOKS": float(row.get("IGST", 0) or 0),
        "Cess tax": float(row.get("Cess", 0) or 0),
        "Remarks": "MRR Return",
        "Source": "MRR Return",
        "DATA RECEIVED IN MONTH": data_month,
    }


def _combined_mrr_rows(mrr_pivot, data_month):
    """All Step 1 MRR pivot rows — Purchase (MRR) and Purchase Return (MRR Return)."""
    if mrr_pivot is None or mrr_pivot.empty:
        return []
    rows = []
    for _, row in mrr_pivot.iterrows():
        if str(row.get("Category", "")) == "Purchase Return":
            rows.append(_row_from_mrr_purchase_return(row, data_month))
        else:
            rows.append(_row_from_mrr_purchase(row, data_month))
    return rows


def _filter_combined_books(books):
    """Combined sheet: S2 (Expense/Ineligible remarks) + full MRR + MRR Return."""
    if books is None or books.empty:
        return books
    if "Source" not in books.columns:
        return books
    return books[books["Source"].isin(["Expense", "Expense-Ineligible", "MRR", "MRR Return"])].copy()


def _filter_combined_remarks(books):
    """S2: Remark Expense/Ineligible only. MRR rows use MRR / MRR Return remarks."""
    if books is None or books.empty or "Remarks" not in books.columns:
        return books
    mrr_mask = books["Source"].isin(["MRR", "MRR Return"])
    s2_mask = books["Source"].isin(["Expense", "Expense-Ineligible"]) & (
        books["Remarks"].astype(str).isin(S2_COMBINED_REMARKS)
    )
    return books[mrr_mask | s2_mask].copy()


def _filter_s2_expense_books(books):
    """S3_Books_Combined uses Step 2 expenses only (not MRR / Step 1)."""
    if books is None or books.empty:
        return books
    if "Source" not in books.columns:
        return books
    return books[books["Source"].isin(["Expense", "Expense-Ineligible"])].copy()


def _filter_combined_categories(books):
    """S3_Books_Combined shows only these Category values (excludes ISD, RCM, Import, etc.)."""
    if books is None or books.empty:
        return books
    return books[books["Category"].astype(str).isin(BOOKS_COMBINED_CATEGORIES)].copy()


def format_unified_books(books):
    """Single S3 export — Books Final layout + ITC Type + Source."""
    books = _filter_combined_books(books)
    books = _filter_combined_categories(books)
    books = _filter_combined_remarks(books)
    if books is None or books.empty:
        cols = ["CON", "ITC Type", "Source"] + [c for c in ELIGIBLE_EXPORT_COLS if c != "CON"]
        return pd.DataFrame(columns=cols)
    out = books.copy()
    if "Source" in out.columns:
        itc_type = out["Source"].apply(
            lambda s: "Ineligible" if str(s) == "Expense-Ineligible" else "Eligible"
        )
        source = out["Source"]
    else:
        itc_type = pd.Series(["Eligible"] * len(out), index=out.index)
        source = pd.Series(["Existing Books"] * len(out), index=out.index)
    formatted = format_eligible_books(out)
    formatted.insert(1, "ITC Type", itc_type.values)
    formatted.insert(2, "Source", source.values)
    return formatted


def split_s3_export_sheets(all_formatted):
    """Split unified Step 3 output into Combined (Expense + MRR + MRR Return) and ineligible sheets."""
    empty_cols = ["CON", "ITC Type", "Source"] + [c for c in ELIGIBLE_EXPORT_COLS if c != "CON"]
    if all_formatted is None or all_formatted.empty or "Source" not in all_formatted.columns:
        empty = pd.DataFrame(columns=empty_cols)
        return empty.copy(), empty.copy()
    combined = all_formatted[all_formatted["Source"].isin(S3_COMBINED_SOURCES)].copy()
    if not combined.empty and "Category" in combined.columns:
        combined = combined[combined["Category"].astype(str).isin(S3_COMBINED_CATEGORIES)].copy()
    combined = _enforce_combined_source_category(combined)
    ineligible = all_formatted[all_formatted["Source"] == S3_INELIGIBLE_SOURCE].copy()
    if not ineligible.empty and "Rate" in ineligible.columns:
        ineligible = ineligible[ineligible["Rate"].apply(is_ineligible_debit)].copy()
    return combined, ineligible


def format_combined_books(eligible_df, ineligible_df):
    """Backward-compatible wrapper — builds unified sheet from compiled books rows."""
    parts = [df for df in (eligible_df, ineligible_df) if df is not None and not df.empty]
    if not parts:
        return format_unified_books(pd.DataFrame())
    return format_unified_books(pd.concat(parts, ignore_index=True))


def format_ineligible_books(books):
    """Align Ineligible ITC output to Books Final (Ineligible sheet) column layout."""
    if books is None or books.empty:
        return pd.DataFrame(columns=INELIGIBLE_EXPORT_COLS)
    out = books.copy()
    out["State Receiver"] = out.get("State", "")
    out["State BOOKS"] = out.get("StateBOOKS", out.get("State", ""))
    out["InvoiceNO_Books"] = out.get("Invoice No.", "")
    out["Vch No."] = out.get("Voucher No.", out.get("Vch No.", ""))
    out["Remarks 1"] = out.get("REMARKS1", "")
    for col in INELIGIBLE_EXPORT_COLS:
        if col not in out.columns:
            out[col] = ""
    return out[INELIGIBLE_EXPORT_COLS]


def build_books_from_sources(
    mrr_pivot, eligible_expenses, ineligible_expenses=None,
    all_expenses=None, data_month="",
):
    """Step 3: compile Eligible ITC (MRR + expenses) and Ineligible ITC sheets."""
    eligible_rows = []
    ineligible_rows = []
    mrr_cons = set()

    if mrr_pivot is not None and not mrr_pivot.empty:
        for _, row in mrr_pivot.iterrows():
            con = row["con"]
            mrr_cons.add(con)
            ship_state = row.get("Ship State", "")
            eligible_rows.append({
                "CON": con,
                "Category": row.get("Category", "Purchase"),
                "State": ship_state,
                "StateBOOKS": ship_state,
                "Supplier name": row.get("Supplier Name", ""),
                "GSTIN": row.get("Supplier GSTIN", ""),
                "Invoice No.": row.get("Invoice No.", ""),
                "Voucher No.": row.get("Voucher No.", ""),
                "Rate": row.get("Tax Slab", ""),
                "SGST BOOKS": float(row.get("SGST", 0) or 0),
                "CGST BOOKS": float(row.get("CGST", 0) or 0),
                "IGST BOOKS": float(row.get("IGST", 0) or 0),
                "Cess tax": float(row.get("Cess", 0) or 0),
                "Source": "MRR",
                "DATA RECEIVED IN MONTH": data_month,
            })

    if eligible_expenses is not None and not eligible_expenses.empty:
        for _, row in eligible_expenses.iterrows():
            con = row["con"]
            if con in mrr_cons:
                continue
            eligible_rows.append(_row_from_s2_eligible(row, data_month))

    if ineligible_expenses is not None and not ineligible_expenses.empty:
        for _, row in ineligible_expenses.iterrows():
            if not _is_allowed_ineligible_rate_label(row.get("Rate Label", "")):
                continue
            ineligible_rows.append(_row_from_s2_ineligible(row, data_month))

    books_eligible = _aggregate_books(eligible_rows)
    books_ineligible = _aggregate_books(ineligible_rows)

    books_s2_eligible = _aggregate_books(
        _combined_s2_expense_rows(all_expenses, data_month, mrr_cons),
        group_keys=None,
    )
    books_combined_ineligible = _aggregate_books(
        _combined_s2_ineligible_rows(all_expenses, data_month),
        group_keys=None,
    )
    books_mrr = _aggregate_books(
        _combined_mrr_rows(mrr_pivot, data_month), group_keys=["CON", "Category"]
    )
    combined_parts = [books_s2_eligible, books_combined_ineligible, books_mrr]
    books_s2 = pd.concat(
        [part for part in combined_parts if len(part)], ignore_index=True
    )

    books = pd.concat(
        [books_eligible, books_ineligible], ignore_index=True
    ) if len(books_ineligible) else books_eligible

    all_formatted = format_unified_books(books_s2)
    books_combined, books_ineligible_combined = split_s3_export_sheets(all_formatted)
    if data_month:
        books_combined = _fill_data_received_month(books_combined, data_month)
        books_ineligible_combined = _fill_data_received_month(books_ineligible_combined, data_month)

    summary = {
        "total_rows": len(books),
        "eligible_rows": len(books_eligible),
        "ineligible_rows": len(books_ineligible),
        "mrr_rows": (
            len(books_eligible[books_eligible["Source"] == "MRR"])
            if len(books_eligible) and "Source" in books_eligible.columns else 0
        ),
        "expense_rows": (
            len(books_eligible[books_eligible["Source"] == "Expense"])
            if len(books_eligible) and "Source" in books_eligible.columns else 0
        ),
        "total_itc": float(books["Total ITC As per Books"].sum()) if len(books) else 0,
        "eligible_itc": float(books_eligible["Total ITC As per Books"].sum()) if len(books_eligible) else 0,
        "ineligible_itc": float(books_ineligible["Total ITC As per Books"].sum()) if len(books_ineligible) else 0,
        "total_igst": float(books["IGST BOOKS"].sum()) if len(books) else 0,
        "total_cgst": float(books["CGST BOOKS"].sum()) if len(books) else 0,
        "total_sgst": float(books["SGST BOOKS"].sum()) if len(books) else 0,
        "combined_rows": len(books_combined),
        "ineligible_combined_rows": len(books_ineligible_combined),
        "s2_combined_source_rows": len(books_s2),
        "excluded_combined_rows": int(len(books_s2) - len(_filter_combined_categories(books_s2))),
    }

    return {
        "books": books,
        "books_eligible": format_eligible_books(books_eligible),
        "books_ineligible": format_ineligible_books(books_ineligible),
        "books_combined": books_combined,
        "ineligible": books_ineligible_combined,
        "summary": summary,
    }


def load_existing_books(file_obj, month_filter=None):
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name="25-26", header=2)
    df.columns = [str(c).strip() for c in df.columns]

    if "Total ITC As per Books" not in df.columns:
        for col in ["SGST BOOKS", "CGST BOOKS", "IGST BOOKS", "Cess tax"]:
            if col in df.columns:
                df[col] = safe_numeric(df[col])
        df["Total ITC As per Books"] = (
            df.get("SGST BOOKS", 0) + df.get("CGST BOOKS", 0) +
            df.get("IGST BOOKS", 0) + df.get("Cess tax", 0)
        )

    if month_filter and "MONTH OF ITC TAKEN" in df.columns:
        df["_month"] = pd.to_datetime(df["MONTH OF ITC TAKEN"], errors="coerce")
        target = pd.to_datetime(month_filter)
        filtered = df[df["_month"].dt.to_period("M") == target.to_period("M")]
        if len(filtered) > 0:
            df = filtered.drop(columns=["_month"])
    elif month_filter and "Vch Date" in df.columns:
        df["_month"] = pd.to_datetime(df["Vch Date"], errors="coerce")
        target = pd.to_datetime(month_filter)
        filtered = df[df["_month"].dt.to_period("M") == target.to_period("M")]
        if len(filtered) > 0:
            df = filtered.drop(columns=["_month"])

    return df
