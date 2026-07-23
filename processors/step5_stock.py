"""Step 5: Stock Received — GIT, pivot, GSTR-2B TAKEN IN 3B / Reclaimed."""
import pandas as pd
from .utils import normalize_invoice, safe_numeric


def _find_sheet(xl, *patterns):
    for name in xl.sheet_names:
        low = name.lower()
        if any(p.lower() in low for p in patterns):
            return name
    return None


def read_stock_sheets(file_obj):
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    xl = pd.ExcelFile(file_obj)
    stock_recd_sheet = _find_sheet(xl, "stock recd in", "stock received")
    stock_trf_sheet = _find_sheet(xl, "stock transfer gst", "stock transfer")
    git_current_sheet = _find_sheet(xl, "stocknotrecd", "git) in")
    git_prev_sheet = _find_sheet(xl, "git of")

    if not stock_recd_sheet:
        raise ValueError(f"Stock received sheet not found. Available: {xl.sheet_names}")
    if not stock_trf_sheet:
        raise ValueError(f"Stock transfer sheet not found. Available: {xl.sheet_names}")

    def _read(sheet):
        if hasattr(file_obj, "seek"):
            file_obj.seek(0)
        return pd.read_excel(file_obj, sheet_name=sheet)

    stock_recd = _read(stock_recd_sheet)
    stock_trf = _read(stock_trf_sheet)
    git_current = _read(git_current_sheet) if git_current_sheet else pd.DataFrame()
    git_prev = _read(git_prev_sheet) if git_prev_sheet else pd.DataFrame()
    return stock_recd, stock_trf, git_current, git_prev, stock_recd_sheet


def process_step5(stock_file, g2b_df=None, reco_month_label=""):
    """
    Step 5 workflow:
    - Stock received in current month (month from invoice date)
    - Stock transfer mapping + GIT current month
    - GIT previous month mapping with current stock received
    - Pivot by ship state / month / IGST
    - Update GSTR-2B TAKEN IN 3B / Reclaimed remarks
    """
    stock_recd, stock_trf, git_current, git_prev, _ = read_stock_sheets(stock_file)

    # Step 5.3: Month column from invoice date
    stock_recd = stock_recd.copy()
    if "inv_no" in stock_recd.columns:
        stock_recd["invoice_no_norm"] = stock_recd["inv_no"].apply(normalize_invoice)
    if "inv_date" in stock_recd.columns:
        stock_recd["inv_date"] = pd.to_datetime(stock_recd["inv_date"], errors="coerce")
        stock_recd["Month"] = stock_recd["inv_date"].dt.strftime("%b-%y")

    for col in ["igstamt", "cgstamt", "sgstamt", "cessamt", "inv_tot"]:
        if col in stock_recd.columns:
            stock_recd[col] = safe_numeric(stock_recd[col])

    # Step 5.4: Stock transfer mapping with stock received
    stock_trf = stock_trf.copy()
    if "inv_no" in stock_trf.columns:
        stock_trf["invoice_no_norm"] = stock_trf["inv_no"].apply(normalize_invoice)

    recd_invoices = set(stock_recd["invoice_no_norm"].dropna().unique()) if "invoice_no_norm" in stock_recd.columns else set()
    recd_totals = (
        stock_recd.groupby("invoice_no_norm")["inv_tot"].sum().to_dict()
        if "inv_tot" in stock_recd.columns else {}
    )

    stock_trf["matched_in_stock_recd"] = stock_trf["invoice_no_norm"].isin(recd_invoices) if "invoice_no_norm" in stock_trf.columns else False
    stock_trf["stock_recd_total"] = stock_trf["invoice_no_norm"].map(recd_totals) if "invoice_no_norm" in stock_trf.columns else 0
    if "inv_tot" in stock_trf.columns:
        stock_trf["value_diff"] = safe_numeric(stock_trf["inv_tot"]) - safe_numeric(stock_trf["stock_recd_total"])

    # GIT of current month = stock trf not found in stock received
    git_current_calc = stock_trf[~stock_trf["matched_in_stock_recd"]].copy() if "matched_in_stock_recd" in stock_trf.columns else pd.DataFrame()

    # Step 5.5: GIT of previous month — map with current stock received
    git_prev_mapped = pd.DataFrame()
    if not git_prev.empty and "inv_no" in git_prev.columns:
        git_prev_mapped = git_prev.copy()
        git_prev_mapped["invoice_no_norm"] = git_prev_mapped["inv_no"].apply(normalize_invoice)
        if "inv_date" in git_prev_mapped.columns:
            git_prev_mapped["inv_date"] = pd.to_datetime(git_prev_mapped["inv_date"], errors="coerce")
            git_prev_mapped["Month"] = git_prev_mapped["inv_date"].dt.strftime("%b-%y")
        git_prev_mapped["matched_in_current_recd"] = git_prev_mapped["invoice_no_norm"].isin(recd_invoices)
        git_prev_mapped["stock_recd_total"] = git_prev_mapped["invoice_no_norm"].map(recd_totals)
        if "inv_tot" in git_prev_mapped.columns:
            git_prev_mapped["value_diff"] = (
                safe_numeric(git_prev_mapped["inv_tot"]) - safe_numeric(git_prev_mapped["stock_recd_total"])
            )

    # Step 5.6: Pivot — ship state, taxable value, IGST month-wise
    pivot = pd.DataFrame()
    if "ship_state" in stock_recd.columns and "Month" in stock_recd.columns:
        pivot = stock_recd.groupby(["ship_state", "Month"], dropna=False).agg({
            "inv_tot": "sum",
            "igstamt": "sum",
            "cgstamt": "sum",
            "sgstamt": "sum",
        }).reset_index()
        pivot.rename(columns={
            "inv_tot": "Taxable Value",
            "igstamt": "IGST Amount",
            "cgstamt": "CGST Amount",
            "sgstamt": "SGST Amount",
        }, inplace=True)

    # Step 5.7: Update GSTR-2B — TAKEN IN 3B / Reclaimed by invoice mapping
    g2b_stock_updates = pd.DataFrame()
    g2b_updated = g2b_df.copy() if g2b_df is not None else pd.DataFrame()

    if g2b_df is not None and not g2b_df.empty:
        g2b = g2b_df.copy()
        inv_col = "Invoice No." if "Invoice No." in g2b.columns else None
        if inv_col:
            g2b["invoice_norm"] = g2b[inv_col].apply(normalize_invoice)
            matched_mask = g2b["invoice_norm"].isin(recd_invoices)
            g2b.loc[matched_mask, "TAKEN IN 3B"] = "Yes"
            g2b.loc[matched_mask, "Reclaimed"] = ""
            if "Add. Remarks" in g2b.columns:
                stock_tag = f"Stock Received {reco_month_label}".strip()
                g2b.loc[matched_mask, "Add. Remarks"] = g2b.loc[matched_mask, "Add. Remarks"].apply(
                    lambda x: f"{x}; {stock_tag}" if x and str(x) not in ("nan", "") else stock_tag
                )
            g2b_stock_updates = g2b[matched_mask].copy()
            g2b_updated = g2b

    summary = {
        "stock_recd_rows": len(stock_recd),
        "stock_trf_rows": len(stock_trf),
        "git_current_rows": len(git_current),
        "git_prev_rows": len(git_prev),
        "git_current_month_calc": len(git_current_calc),
        "git_prev_mapped_rows": len(git_prev_mapped),
        "total_igst_stock": float(stock_recd["igstamt"].sum()) if "igstamt" in stock_recd.columns else 0,
        "total_taxable_stock": float(stock_recd["inv_tot"].sum()) if "inv_tot" in stock_recd.columns else 0,
        "g2b_stock_matched": len(g2b_stock_updates),
    }

    return {
        "stock_recd": stock_recd,
        "stock_trf": stock_trf,
        "git_current": git_current_calc if len(git_current_calc) else git_current,
        "git_prev": git_prev_mapped if len(git_prev_mapped) else git_prev,
        "pivot": pivot,
        "g2b_stock_updates": g2b_stock_updates,
        "g2b_updated": g2b_updated,
        "summary": summary,
    }
