"""Step 2: Process Master Expenses — Working (All Type) sheet."""
import re
import pandas as pd
from .utils import (
    normalize_gstin, normalize_invoice, make_con_key,
    validate_gstin, validate_invoice, extract_gstin_from_narration,
    safe_numeric, validate_gst_head, sum_expense_tax_columns,
    is_ineligible_debit, is_ineligible_credit,
)


_GSTIN_RE = re.compile(r"[0-9]{2}[A-Z0-9]{13}")


def _clean_gstin_uin(val):
    """Extract valid 15-char GSTIN from dirty expense cells (Tally/Excel exports)."""
    if pd.isna(val):
        return ""
    s = str(val).strip().upper()
    if not s or s in ("NAN", "NONE"):
        return ""
    for artifact in ("_X000D_", "_X000A_", "_X0009_", "\r", "\n", "\t"):
        s = s.replace(artifact, "")
    s = s.replace(" ", "")
    if re.fullmatch(r"[0-9]{2}[A-Z0-9]{13}", s):
        return s
    if s.startswith("GSTIN"):
        s = s[5:]
    elif s.startswith("GST"):
        s = s[3:]
    s = s.lstrip("#-_/\\|")
    if re.fullmatch(r"[0-9]{2}[A-Z0-9]{13}", s):
        return s
    matches = _GSTIN_RE.findall(s)
    if matches:
        return matches[0]
    return normalize_gstin(val)


def _find_working_sheet(xl):
    for name in xl.sheet_names:
        low = name.lower()
        if "working" in low and "all type" in low:
            return name
    for name in xl.sheet_names:
        if "working" in name.lower():
            return name
    return "Working (All Type)"


def _detect_header_row(file_obj, sheet_name):
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    preview = pd.read_excel(file_obj, sheet_name=sheet_name, header=None, nrows=15)
    for i in range(len(preview)):
        cells = [str(v).strip().lower() for v in preview.iloc[i].tolist() if pd.notna(v)]
        if any("voucher type" in c for c in cells):
            return i
    return 0


def read_expenses(file_obj):
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    xl = pd.ExcelFile(file_obj)
    sheet = _find_working_sheet(xl)
    header_row = _detect_header_row(file_obj, sheet)
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def _row_val(row, *names):
    """Read cell value with flexible column name matching."""
    for name in names:
        if name in row.index and pd.notna(row.get(name)):
            val = str(row[name]).strip()
            if val.lower() not in ("", "nan", "none"):
                return val
    targets = {n.lower() for n in names}
    for col in row.index:
        if str(col).strip().lower() in targets:
            val = row[col]
            if pd.notna(val):
                val = str(val).strip()
                if val.lower() not in ("", "nan", "none"):
                    return val
    return ""


def _rate_value(row):
    """Column AO — Rate field used for Ineligible / RCM classification."""
    val = _row_val(row, "Rate")
    return val if val else ""


def _normalize_voucher_type(vt):
    vt = str(vt or "").strip().upper()
    for ch in ("\u2010", "\u2011", "\u2012", "\u2013", "\u2014", "\u2212"):
        vt = vt.replace(ch, "-")
    return vt


def assign_remark(row):
    """Remark options: Purchase, Expense, ISD, Import, RCM, Ineligible."""
    rate_label = str(row.get("Rate Label", "") or "")
    if is_ineligible_debit(rate_label) or row.get("Category") == "Ineligible":
        return "Ineligible"
    return row["Category"]


def categorize_entry(row):
    """Step 2.1: Remark column — per Modicare Working (All Type) spec."""
    vt = _normalize_voucher_type(_row_val(row, "Voucher Type"))
    particulars = _row_val(row, "Particulars").upper()
    rate_val = _rate_value(row).upper()

    if "-PUR" in vt:
        return "Purchase"
    if "-ISD" in vt:
        return "ISD"
    if "CUSTOM" in particulars:
        return "Import"
    if "INELIGIBLE" in rate_val:
        return "Ineligible"
    if "RCM" in rate_val:
        return "RCM"
    return "Expense"


def _format_date_only(val):
    """Strip time from Excel datetime — show date only (dd/mm/yyyy)."""
    if pd.isna(val) or str(val).strip().lower() in ("", "nan", "none"):
        return val
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return val
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return val


def _standardize_asc_remarks_column(df):
    """Ensure master Asc Remarks column is available under a consistent name."""
    if "Asc Remarks" in df.columns:
        return df
    for col in df.columns:
        if str(col).strip().lower().replace(" ", "") in ("ascremarks", "ascremark"):
            return df.rename(columns={col: "Asc Remarks"})
    return df


def _is_blank_cell(val):
    return pd.isna(val) or str(val).strip().lower() in ("", "nan", "none")


def _skip_expense_output_row(row):
    """Drop footer totals and journal summary lines from S2 export."""
    gstin = row.get("GSTIN/UIN", "")
    bill_no = row.get("Bill No", "")
    particulars = row.get("Particulars", "")
    voucher_type = row.get("Voucher Type", "")
    voucher_no = row.get("Voucher No.", "")

    if not _is_blank_cell(gstin) or not _is_blank_cell(bill_no) or not _is_blank_cell(particulars):
        return False

    if _is_blank_cell(voucher_type) and _is_blank_cell(voucher_no):
        return True

    if not _is_blank_cell(voucher_type) and _is_blank_cell(gstin) and _is_blank_cell(bill_no):
        return True

    return False


def process_step2(expense_file):
    raw = read_expenses(expense_file)
    df = raw.copy()
    df = _standardize_asc_remarks_column(df)

    # Drop pre-filled Category/Remark from source file — always recalculate
    for old_col in ("Category", "Remark", "category", "remark"):
        if old_col in df.columns:
            df = df.drop(columns=[old_col])

    df["Category"] = df.apply(categorize_entry, axis=1)
    df["Rate Label"] = df.apply(_rate_value, axis=1)
    df["Remark"] = df.apply(assign_remark, axis=1)

    if "GSTIN/UIN" not in df.columns:
        df["GSTIN/UIN"] = ""
    df["GSTIN/UIN"] = df["GSTIN/UIN"].apply(_clean_gstin_uin)
    missing_gstin = df["GSTIN/UIN"] == ""
    if "Voucher Narration" in df.columns:
        df.loc[missing_gstin, "GSTIN/UIN"] = (
            df.loc[missing_gstin, "Voucher Narration"]
            .apply(extract_gstin_from_narration)
            .apply(_clean_gstin_uin)
        )

    bill_col = "Bill No" if "Bill No" in df.columns else "Voucher Ref."

    def _resolve_bill_no(row):
        for col in (bill_col, "Bill No", "Voucher Ref.", "Voucher Ref"):
            if col in row.index and pd.notna(row.get(col)):
                val = normalize_invoice(row.get(col))
                if val:
                    return val
        return ""

    df["Bill No"] = df.apply(_resolve_bill_no, axis=1)
    df["con"] = df.apply(lambda r: make_con_key(r["GSTIN/UIN"], r["Bill No"]), axis=1)
    df = df[~df.apply(_skip_expense_output_row, axis=1)].copy().reset_index(drop=True)

    igst, cgst, sgst, cess = sum_expense_tax_columns(df)
    df["IGST"] = igst
    df["CGST"] = cgst
    df["SGST"] = sgst
    df["Cess"] = cess
    df["Total ITC"] = df["IGST"] + df["CGST"] + df["SGST"] + df["Cess"]
    if "Rate" in df.columns:
        df["Tax Rate"] = df["Rate"]

    if "Voucher Date" in df.columns:
        df["Voucher Date"] = df["Voucher Date"].apply(_format_date_only)
    if "Bill Date" in df.columns:
        df["Bill Date"] = df["Bill Date"].apply(_format_date_only)

    issues = []
    validation_by_row = {}
    for idx, row in df.iterrows():
        row_issues = []
        ok, msg = validate_gstin(row["GSTIN/UIN"])
        if not ok:
            row_issues.append("GSTIN is missing" if msg == "Missing GSTIN" else msg)
        ok, msg = validate_invoice(row["Bill No"])
        if not ok:
            row_issues.append(msg)
        ok, msg = validate_gst_head(
            row["GSTIN/UIN"], None,
            row["IGST"], row["CGST"], row["SGST"],
        )
        if not ok and row["Category"] not in ("Import", "ISD"):
            row_issues.append(msg)
        if row_issues:
            validation_by_row[idx] = "; ".join(row_issues)
            issues.append({
                "Row": idx + 1, "CON": row["con"],
                "Category": row["Category"], "Issues": "; ".join(row_issues),
            })

    df["Validation"] = "OK"
    for idx, msg in validation_by_row.items():
        df.at[idx, "Validation"] = msg

    ineligible_mask = df["Category"] == "Ineligible"
    ineligible_mask &= df["Rate Label"].apply(is_ineligible_debit)
    ineligible = df[ineligible_mask & (df["Validation"] == "OK")].copy()

    eligible = df[~df["Category"].isin(["Ineligible"]) & (df["Validation"] == "OK")].copy()

    category_counts = df["Category"].value_counts().to_dict()

    output_cols = [
        "Created by", "Voucher Date", "Voucher No.", "Voucher Type",
        "Voucher Narration", "Voucher Ref.", "Bill No", "Bill Date",
        "Particulars", "GSTIN/UIN", "Category", "Remark", "Asc Remarks",
        "Rate Label", "Tax Rate",
        "con", "IGST", "CGST", "SGST", "Cess", "Total ITC", "Validation",
    ]
    output_cols = [c for c in output_cols if c in df.columns]

    summary = {
        "total_rows": len(df),
        "eligible_rows": len(eligible),
        "ineligible_rows": len(ineligible),
        "ineligible_credit_ignored": int(
            (df["Rate Label"].apply(is_ineligible_credit)).sum()
        ),
        "wrong_entries": len(issues),
        "category_counts": category_counts,
        "remark_counts": df["Remark"].value_counts().to_dict(),
        "purchase_rows": category_counts.get("Purchase", 0),
        "expense_rows": category_counts.get("Expense", 0),
        "total_igst": float(df["IGST"].sum()),
        "total_cgst": float(df["CGST"].sum()),
        "total_sgst": float(df["SGST"].sum()),
        "total_itc": float(df["Total ITC"].sum()),
    }

    return {
        "all_expenses": df[output_cols],
        "eligible": eligible[output_cols],
        "ineligible": ineligible[output_cols] if len(ineligible) else pd.DataFrame(),
        "issues": pd.DataFrame(issues),
        "summary": summary,
    }
