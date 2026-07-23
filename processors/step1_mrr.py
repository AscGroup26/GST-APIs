"""Step 1: Compile MRR and MRR Return data."""
import re
import pandas as pd
from .utils import (
    normalize_gstin, normalize_invoice, make_con_key, make_con_key_match,
    validate_gstin, validate_invoice, safe_numeric,
)


def _is_blank_voucher(val):
    s = str(val or "").strip()
    return s in ("", "nan", "None")


def _invoice_tail(val):
    s = str(val or "").strip().upper()
    if "/" in s:
        s = s.rsplit("/", 1)[-1]
    return re.sub(r"[^A-Z0-9]", "", s)


def _fiscal_normalize_invoice(val):
    s = normalize_invoice(val)
    return re.sub(r"(\d{4})-(\d{4})", lambda m: f"{m.group(1)}-{m.group(2)[-2:]}", s)


def _expense_lookup_keys(gstin, invoice_no):
    g = normalize_gstin(gstin)
    inv = str(invoice_no or "")
    fiscal = _fiscal_normalize_invoice(inv)
    keys = [
        make_con_key(g, inv), make_con_key(g, fiscal),
        make_con_key_match(g, inv), make_con_key_match(g, fiscal),
    ]
    tail = _invoice_tail(inv)
    if tail:
        keys.append(g + tail)
    return keys


def _is_blank_date(val):
    if pd.isna(val):
        return True
    return str(val).strip().lower() in ("", "nan", "none")


def _format_voucher_date(val):
    if _is_blank_date(val):
        return ""
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return str(val).strip()
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return str(val).strip()


def _build_expense_field_map(expense_df, field_name, is_blank=_is_blank_voucher, formatter=None):
    field_map = {}
    if expense_df is None or expense_df.empty or field_name not in expense_df.columns:
        return field_map
    exp = expense_df.copy()
    exp["_has_value"] = ~exp[field_name].apply(is_blank)
    exp = exp.sort_values("_has_value", ascending=False)
    for _, row in exp.iterrows():
        val = row.get(field_name, "")
        if is_blank(val):
            continue
        val = formatter(val) if formatter else str(val).strip()
        if is_blank(val):
            continue
        for key in _expense_lookup_keys(row.get("GSTIN/UIN", ""), row.get("Bill No", "")):
            field_map.setdefault(key, val)
    return field_map


def _build_expense_voucher_map(expense_df):
    return _build_expense_field_map(expense_df, "Voucher No.")


def _build_expense_voucher_date_map(expense_df):
    return _build_expense_field_map(
        expense_df, "Voucher Date", is_blank=_is_blank_date, formatter=_format_voucher_date
    )


def _lookup_expense_field(field_map, gstin, invoice_no):
    for key in _expense_lookup_keys(gstin, invoice_no):
        if key in field_map:
            return field_map[key]
    return ""


def _lookup_expense_voucher(voucher_map, gstin, invoice_no):
    return _lookup_expense_field(voucher_map, gstin, invoice_no)


def _lookup_expense_voucher_date(voucher_date_map, gstin, invoice_no):
    return _lookup_expense_field(voucher_date_map, gstin, invoice_no)


def _fill_mrr_return_vouchers(combined, expense_df=None, voucher_map=None, voucher_date_map=None):
    if "voucher_no" not in combined.columns:
        combined["voucher_no"] = ""
    if "voucher_date" not in combined.columns:
        combined["voucher_date"] = ""
    if voucher_map is None and expense_df is not None and not expense_df.empty:
        voucher_map = _build_expense_voucher_map(expense_df)
    if voucher_date_map is None and expense_df is not None and not expense_df.empty:
        voucher_date_map = _build_expense_voucher_date_map(expense_df)
    return_mask = (
        combined["Remark"].astype(str).str.contains("Return", case=False, na=False)
        & combined["voucher_no"].apply(_is_blank_voucher)
    )
    if not return_mask.any() and not (
        combined["Remark"].astype(str).str.contains("Return", case=False, na=False)
        & combined["voucher_date"].apply(_is_blank_date)
    ).any():
        return combined
    mrr_mapped = combined[combined["Remark"].eq("MRR") & ~combined["voucher_no"].apply(_is_blank_voucher)]
    if not mrr_mapped.empty:
        for idx in combined.index[return_mask]:
            gstin = combined.at[idx, "supplier_gstin"]
            tail = _invoice_tail(combined.at[idx, "invoice_no"])
            for _, mrr_row in mrr_mapped.iterrows():
                if mrr_row["supplier_gstin"] == gstin and _invoice_tail(mrr_row["invoice_no"]) == tail:
                    combined.at[idx, "voucher_no"] = mrr_row["voucher_no"]
                    if _is_blank_date(combined.at[idx, "voucher_date"]) and not _is_blank_date(mrr_row.get("voucher_date", "")):
                        combined.at[idx, "voucher_date"] = mrr_row["voucher_date"]
                    break
    return_mask = (
        combined["Remark"].astype(str).str.contains("Return", case=False, na=False)
        & combined["voucher_no"].apply(_is_blank_voucher)
    )
    if voucher_map and return_mask.any():
        for idx in combined.index[return_mask]:
            voucher = _lookup_expense_voucher(voucher_map, combined.at[idx, "supplier_gstin"], combined.at[idx, "invoice_no"])
            if voucher:
                combined.at[idx, "voucher_no"] = voucher
    date_return_mask = (
        combined["Remark"].astype(str).str.contains("Return", case=False, na=False)
        & combined["voucher_date"].apply(_is_blank_date)
    )
    if voucher_date_map and date_return_mask.any():
        for idx in combined.index[date_return_mask]:
            voucher_date = _lookup_expense_voucher_date(
                voucher_date_map, combined.at[idx, "supplier_gstin"], combined.at[idx, "invoice_no"]
            )
            if voucher_date:
                combined.at[idx, "voucher_date"] = voucher_date
    return combined


def read_mrr_csv(file_obj):
    df = pd.read_csv(file_obj, encoding="latin-1")
    df.columns = [c.strip().lower() for c in df.columns]
    return df


def process_step1(mrr_file, mrr_return_file, expense_df=None):
    mrr = read_mrr_csv(mrr_file)
    mrr_return = read_mrr_csv(mrr_return_file)

    mrr["Remark"] = "MRR"
    mrr_return["Remark"] = "MRR Return"

    tax_cols = ["sgstamt", "cgstamt", "igstamt", "ugstamt", "cessamt"]
    for col in tax_cols:
        if col in mrr_return.columns:
            mrr_return[col] = -safe_numeric(mrr_return[col])

    combined = pd.concat([mrr, mrr_return], ignore_index=True)

    combined["supplier_gstin"] = combined.get("gstin2", combined.get("supplier", "")).apply(normalize_gstin)
    combined["invoice_no"] = combined.get("sup_inv", combined.get("inv_no", "")).apply(normalize_invoice)
    combined["con"] = combined.apply(lambda r: make_con_key(r["supplier_gstin"], r["invoice_no"]), axis=1)

    issues = []
    for idx, row in combined.iterrows():
        row_issues = []
        ok, msg = validate_gstin(row["supplier_gstin"])
        if not ok:
            row_issues.append(msg)
        ok, msg = validate_invoice(row["invoice_no"])
        if not ok:
            row_issues.append(msg)
        if row_issues:
            issues.append({"Row": idx + 1, "CON": row["con"], "Issues": "; ".join(row_issues)})

    if expense_df is not None and not expense_df.empty:
        voucher_map = _build_expense_voucher_map(expense_df)
        voucher_date_map = _build_expense_voucher_date_map(expense_df)
        combined["voucher_no"] = combined.apply(
            lambda r: _lookup_expense_voucher(voucher_map, r["supplier_gstin"], r["invoice_no"]), axis=1
        )
        combined["voucher_date"] = combined.apply(
            lambda r: _lookup_expense_voucher_date(voucher_date_map, r["supplier_gstin"], r["invoice_no"]), axis=1
        )
    else:
        voucher_map = {}
        voucher_date_map = {}
        combined["voucher_no"] = ""
        combined["voucher_date"] = ""

    combined = _fill_mrr_return_vouchers(combined, expense_df, voucher_map, voucher_date_map)
    combined["Voucher Date"] = combined["voucher_date"]

    for col in tax_cols:
        if col not in combined.columns:
            combined[col] = 0
        combined[col] = safe_numeric(combined[col])

    pivot = combined.groupby(
        ["supplier_gstin", "supp_name", "ship_state", "invoice_no", "voucher_no", "voucher_date", "taxslab", "Remark"],
        dropna=False,
        as_index=False,
    ).agg({
        "inv_tot": "sum",
        "sgstamt": "sum",
        "cgstamt": "sum",
        "igstamt": "sum",
        "cessamt": "sum",
        "sup_dt": "first",
    })

    pivot.rename(columns={
        "supplier_gstin": "Supplier GSTIN",
        "supp_name": "Supplier Name",
        "ship_state": "Ship State",
        "invoice_no": "Invoice No.",
        "sup_dt": "Invoice Date",
        "voucher_no": "Voucher No.",
        "voucher_date": "Voucher Date",
        "taxslab": "Tax Slab",
        "inv_tot": "Taxable Value",
        "sgstamt": "SGST",
        "cgstamt": "CGST",
        "igstamt": "IGST",
        "cessamt": "Cess",
    }, inplace=True)

    pivot["Category"] = pivot["Remark"].apply(
        lambda r: "Purchase Return" if "MRR Return" in str(r) else "Purchase"
    )
    pivot["con"] = pivot.apply(lambda r: make_con_key(r["Supplier GSTIN"], r["Invoice No."]), axis=1)

    summary = {
        "mrr_rows": len(mrr),
        "mrr_return_rows": len(mrr_return),
        "combined_rows": len(combined),
        "pivot_rows": len(pivot),
        "validation_issues": len(issues),
        "total_igst": float(pivot["IGST"].sum()),
        "total_cgst": float(pivot["CGST"].sum()),
        "total_sgst": float(pivot["SGST"].sum()),
    }

    return {
        "combined": combined,
        "pivot": pivot,
        "issues": pd.DataFrame(issues),
        "summary": summary,
    }
