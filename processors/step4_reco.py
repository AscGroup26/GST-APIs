"""Step 4: Reconcile Books with GSTR-2B — eligible + ineligible ITC."""
import pandas as pd
from .utils import make_con_key, make_con_key_match, safe_numeric, tax_match


def read_gstr2b(file_obj):
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name="GSTR-2B", header=3)
    df.columns = [str(c).strip() for c in df.columns]
    if "Con" not in df.columns:
        df["Con"] = df.apply(
            lambda r: make_con_key(r.get("GSTIN", ""), r.get("Invoice No.", "")),
            axis=1,
        )
    df["Con_Match"] = df.apply(
        lambda r: make_con_key_match(r.get("GSTIN", ""), r.get("Invoice No.", "")),
        axis=1,
    )
    if "Eligibility" in df.columns:
        df["ITC Type"] = df["Eligibility"].apply(
            lambda x: "Eligible" if str(x).upper() in ("Y", "YES", "ELIGIBLE") else "Ineligible"
        )
    else:
        df["ITC Type"] = "Eligible"
    return df


def _find_g2b_row(con, con_match, g2b_indexed, g2b_match_indexed):
    if con in g2b_indexed.index:
        row = g2b_indexed.loc[con]
        return row.iloc[0] if isinstance(row, pd.DataFrame) else row
    if con_match in g2b_match_indexed.index:
        row = g2b_match_indexed.loc[con_match]
        return row.iloc[0] if isinstance(row, pd.DataFrame) else row
    return None


def _get_books_remark(con, con_match, books):
    books_by_con = books.set_index("CON")
    books_by_match = books.set_index("CON_Match") if "CON_Match" in books.columns else None
    if con in books_by_con.index:
        val = books_by_con.loc[con, "Remarks"]
        return val.iloc[0] if isinstance(val, pd.Series) else val
    if books_by_match is not None and con_match in books_by_match.index:
        val = books_by_match.loc[con_match, "Remarks"]
        return val.iloc[0] if isinstance(val, pd.Series) else val
    return None


def _reconcile_books_subset(books_df, g2b_df, g2b_indexed, g2b_match_indexed, tolerance, preserve_existing_remarks, add_remark_tag):
    books = books_df.copy()
    remarks_list = []

    for idx, row in books.iterrows():
        existing_remark = str(row.get("Remarks", "") or "").strip()
        if preserve_existing_remarks and existing_remark and existing_remark not in ("", "nan"):
            continue

        con = row["CON"]
        con_match = row.get("CON_Match", con)
        b_total = float(row.get("Total ITC As per Books", 0) or 0)
        b_igst = float(row.get("IGST BOOKS", 0) or 0)
        b_cgst = float(row.get("CGST BOOKS", 0) or 0)
        b_sgst = float(row.get("SGST BOOKS", 0) or 0)

        g_row = _find_g2b_row(con, con_match, g2b_indexed, g2b_match_indexed)
        if g_row is None:
            remarks_list.append({
                "idx": idx, "Remarks": "Not in 2B",
                "add remarks": add_remark_tag,
            })
            continue

        g_igst = float(g_row.get("Integrated Tax(₹)", 0) or 0)
        g_cgst = float(g_row.get("Central Tax(₹)", 0) or 0)
        g_sgst = float(g_row.get("State-UT Tax(₹)", 0) or 0)
        g_total = float(g_row.get("NET TAX", 0) or 0)

        igst_ok, _ = tax_match(b_igst, g_igst, tolerance)
        cgst_ok, _ = tax_match(b_cgst, g_cgst, tolerance)
        sgst_ok, _ = tax_match(b_sgst, g_sgst, tolerance)
        total_ok, diff = tax_match(b_total, g_total, tolerance)

        if igst_ok and cgst_ok and sgst_ok and total_ok:
            remark = "PERFECT MATCHED"
        elif total_ok or (igst_ok and cgst_ok and sgst_ok):
            remark = "PROBABLE MATCHED"
        else:
            remark = "Mismatched"

        remarks_list.append({
            "idx": idx, "Remarks": remark,
            "add remarks": f"{add_remark_tag} (diff: {diff:.2f})" if remark == "Mismatched" else add_remark_tag,
            "g2b_igst": g_igst, "g2b_cgst": g_cgst, "g2b_sgst": g_sgst, "g2b_total": g_total,
            "g2b_con": g_row.get("Con", ""),
        })

    for item in remarks_list:
        books.at[item["idx"], "Remarks"] = item["Remarks"]
        books.at[item["idx"], "add remarks"] = item.get("add remarks", "")
        if "g2b_igst" in item:
            books.at[item["idx"], "G2B IGST"] = item["g2b_igst"]
            books.at[item["idx"], "G2B CGST"] = item["g2b_cgst"]
            books.at[item["idx"], "G2B SGST"] = item["g2b_sgst"]
            books.at[item["idx"], "G2B Total"] = item["g2b_total"]
            books.at[item["idx"], "Diff"] = (
                float(books.at[item["idx"], "Total ITC As per Books"] or 0) - item["g2b_total"]
            )

    return books, remarks_list


def reconcile_books_g2b(
    books_df, g2b_df, tolerance=1.0,
    preserve_existing_remarks=False,
    books_ineligible_df=None,
    reco_month_label="",
):
    """
    Step 4: Reconcile eligible + ineligible ITC with GSTR-2B.
    Updates Books remarks and GSTR-2B Books / Add. Remarks columns.
    """
    g2b = g2b_df.copy()
    tag = f"Auto Reconciliation {reco_month_label}".strip()

    for col in ["Integrated Tax(₹)", "Central Tax(₹)", "State-UT Tax(₹)", "Cess tax", "NET TAX"]:
        if col in g2b.columns:
            g2b[col] = safe_numeric(g2b[col])

    g2b_indexed = g2b.set_index("Con", drop=False)
    g2b_match_indexed = g2b.set_index("Con_Match", drop=False)

    # Prepare eligible books
    books = books_df.copy()
    books["CON"] = books["CON"].astype(str).str.strip()
    books["CON_Match"] = books.apply(
        lambda r: make_con_key_match(
            r.get("GSTIN", r.get("Supplier GSTIN", "")),
            r.get("Invoice No.", ""),
        ),
        axis=1,
    )
    for col in ["IGST BOOKS", "CGST BOOKS", "SGST BOOKS", "Cess tax", "Total ITC As per Books"]:
        if col in books.columns:
            books[col] = safe_numeric(books[col])

    # Split eligible vs ineligible if combined
    if books_ineligible_df is not None and not books_ineligible_df.empty:
        ineligible_books = books_ineligible_df.copy()
        eligible_books = books[~books["CON"].isin(ineligible_books["CON"])].copy()
    elif "Category" in books.columns:
        ineligible_books = books[books["Category"] == "Ineligible"].copy()
        eligible_books = books[books["Category"] != "Ineligible"].copy()
    else:
        eligible_books = books.copy()
        ineligible_books = pd.DataFrame()

    # Step 4.2: Eligible ITC reconciliation
    eligible_reco, _ = _reconcile_books_subset(
        eligible_books, g2b, g2b_indexed, g2b_match_indexed,
        tolerance, preserve_existing_remarks, tag,
    )

    # Step 4.3: Ineligible ITC reconciliation
    ineligible_reco = pd.DataFrame()
    if not ineligible_books.empty:
        ineligible_reco, _ = _reconcile_books_subset(
            ineligible_books, g2b, g2b_indexed, g2b_match_indexed,
            tolerance, preserve_existing_remarks,
            f"Ineligible Sheet {reco_month_label}".strip(),
        )

    books_reconciled = pd.concat(
        [eligible_reco, ineligible_reco], ignore_index=True
    ) if not ineligible_reco.empty else eligible_reco

    books_cons = set(books_reconciled["CON"].dropna().unique())
    books_match_cons = set(books_reconciled["CON_Match"].dropna().unique())

    not_in_books = g2b[
        ~g2b["Con_Match"].isin(books_match_cons) & ~g2b["Con"].isin(books_cons)
    ].copy()
    not_in_books["Books_Status"] = "Not in Books"

    # Step 4.2/4.3: Update GSTR-2B Books + Add. Remarks columns
    g2b_updated = g2b.copy()
    g2b_updated["Books"] = g2b_updated.apply(
        lambda r: _get_books_remark(r["Con"], r["Con_Match"], books_reconciled) or "Not in Books",
        axis=1,
    )
    remark_to_add = books_reconciled.set_index("CON")["add remarks"].to_dict() if "add remarks" in books_reconciled.columns else {}
    g2b_updated["Add. Remarks"] = g2b_updated.apply(
        lambda r: remark_to_add.get(r["Con"], r.get("Add. Remarks", "")),
        axis=1,
    )

    remark_counts = books_reconciled["Remarks"].value_counts().to_dict()
    eligible_matched = eligible_reco[
        eligible_reco["Remarks"].isin(["PERFECT MATCHED", "PROBABLE MATCHED"])
    ] if len(eligible_reco) else pd.DataFrame()
    ineligible_unmatched = books_reconciled[
        books_reconciled["Category"].eq("Ineligible") &
        books_reconciled["Remarks"].isin(["Not in 2B", "Mismatched"])
    ] if "Category" in books_reconciled.columns else books_reconciled[
        books_reconciled["Remarks"].isin(["Not in 2B", "Mismatched"])
    ]
    eligible_unmatched = eligible_reco[
        eligible_reco["Remarks"].isin(["Not in 2B", "Mismatched"])
    ] if len(eligible_reco) else pd.DataFrame()

    summary = {
        "total_books": len(books_reconciled),
        "eligible_books": len(eligible_reco),
        "ineligible_books": len(ineligible_reco),
        "total_g2b": len(g2b),
        "perfect_matched": remark_counts.get("PERFECT MATCHED", 0),
        "probable_matched": remark_counts.get("PROBABLE MATCHED", 0),
        "not_in_2b": remark_counts.get("Not in 2B", 0),
        "mismatched": remark_counts.get("Mismatched", 0),
        "not_in_books": len(not_in_books),
        "remark_counts": remark_counts,
        "eligible_itc_books": float(eligible_matched["Total ITC As per Books"].sum()) if len(eligible_matched) else 0,
        "ineligible_itc_books": float(ineligible_unmatched["Total ITC As per Books"].sum()) if len(ineligible_unmatched) else 0,
        "eligible_not_in_2b": len(eligible_unmatched[eligible_unmatched["Remarks"] == "Not in 2B"]) if len(eligible_unmatched) else 0,
        "total_itc_books": float(books_reconciled["Total ITC As per Books"].sum()),
        "total_itc_g2b": float(g2b["NET TAX"].sum()) if "NET TAX" in g2b.columns else 0,
    }

    return {
        "books_reconciled": books_reconciled,
        "books_eligible_reco": eligible_reco,
        "books_ineligible_reco": ineligible_reco,
        "g2b_updated": g2b_updated,
        "not_in_books": not_in_books,
        "summary": summary,
    }
