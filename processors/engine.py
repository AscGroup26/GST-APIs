"""Main ITC reconciliation orchestrator — Steps 1–3 active; 4–5 commented."""
import io
import pandas as pd

from .step1_mrr import process_step1
from .step2_expenses import process_step2
from .step3_books import build_books_from_sources, load_existing_books, format_unified_books
from .utils import infer_data_received_month
# from .step4_reco import read_gstr2b, reconcile_books_g2b
# from .step5_stock import process_step5


STEPS_INFO = [
    {
        "step": 1,
        "title": "MRR & MRR Return",
        "files": ["GST MRR (CSV)", "MRRR Return (CSV)"],
        "description": "Compile MRR and MRR Return, negate return taxes, validate GSTIN/invoice, create pivot.",
    },
    {
        "step": 2,
        "title": "Master Expenses",
        "files": ["Master Expenses (Excel)"],
        "description": "Process Working (All Type) sheet — categorize Purchase/Expense/ISD/Import/RCM/Ineligible.",
    },
    {
        "step": 3,
        "title": "Books Compilation",
        "files": ["Built from Step 1 + Step 2", "Books Final (Excel) — optional"],
        "description": (
            "Post MRR pivot + eligible expenses into Eligible ITC; "
            "post ineligible expenses into Ineligible ITC — CON, Category, State, remark columns."
        ),
    },
    # {
    #     "step": 4,
    #     "title": "GSTR-2B Reconciliation",
    #     "files": ["GSTR-2B (Excel)"],
    #     "description": "Match Books vs GSTR-2B by CON key — Perfect/Probable/Not in 2B/Mismatched.",
    # },
    # {
    #     "step": 5,
    #     "title": "Stock Received",
    #     "files": ["Stock Received (Excel)"],
    #     "description": "Process stock received, GIT tracking, update GSTR-2B remarks.",
    # },
]


def run_full_reconciliation(
    files: dict,
    use_existing_books=False,
    tolerance=1.0,
    month_filter=None,
    preserve_remarks=False,
):
    results = {"steps": {}, "errors": []}
    # month_label = _month_label(month_filter)

    expense_result = None
    if files.get("expenses"):
        try:
            expense_result = process_step2(files["expenses"])
            results["steps"]["step2"] = expense_result
        except Exception as e:
            results["errors"].append(f"Step 2 failed: {e}")

    mrr_result = None
    if files.get("mrr") and files.get("mrr_return"):
        try:
            exp_df = expense_result["all_expenses"] if expense_result else None
            mrr_result = process_step1(files["mrr"], files["mrr_return"], exp_df)
            results["steps"]["step1"] = mrr_result
        except Exception as e:
            results["errors"].append(f"Step 1 failed: {e}")

    if not results["steps"] and not (use_existing_books and files.get("books")):
        results["errors"].append(
            "No steps processed. Upload MRR + MRRR Return (Step 1) and/or Master Expenses (Step 2), "
            "or upload Books Final with 'Use existing Books file' selected."
        )

    # ── Step 3: Books Compilation ──
    books_result = None
    if use_existing_books and files.get("books"):
        try:
            books_df = load_existing_books(files["books"], month_filter=month_filter)
            eligible_itc = (
                float(books_df["Total ITC As per Books"].sum())
                if len(books_df) and "Total ITC As per Books" in books_df.columns else 0
            )
            books_result = {
                "books": books_df,
                "books_eligible": books_df,
                "books_ineligible": pd.DataFrame(),
                "books_combined": format_unified_books(books_df),
                "ineligible": pd.DataFrame(),
                "summary": {
                    "total_rows": len(books_df),
                    "eligible_rows": len(books_df),
                    "ineligible_rows": 0,
                    "source": "existing",
                    "eligible_itc": eligible_itc,
                    "ineligible_itc": 0,
                    "total_itc": eligible_itc,
                },
            }
            results["steps"]["step3"] = books_result
        except Exception as e:
            results["errors"].append(f"Step 3 (existing books) failed: {e}")
    elif mrr_result or expense_result:
        try:
            pivot = mrr_result["pivot"] if mrr_result else None
            eligible = expense_result["eligible"] if expense_result else None
            ineligible = expense_result["ineligible"] if expense_result else None
            all_exp = expense_result["all_expenses"] if expense_result else None
            data_month = infer_data_received_month(
                month_filter=month_filter,
                files=files,
                mrr_result=mrr_result,
                expense_result=expense_result,
            )
            books_result = build_books_from_sources(
                pivot, eligible, ineligible,
                all_expenses=all_exp,
                data_month=data_month,
            )
            results["steps"]["step3"] = books_result
        except Exception as e:
            results["errors"].append(f"Step 3 failed: {e}")

    # ── Step 4: GSTR-2B ── COMMENTED OUT
    # reco_result = None
    # if books_result and files.get("gstr2b"):
    #     reco_result = reconcile_books_g2b(...)

    # ── Step 5: Stock ── COMMENTED OUT
    # if files.get("stock"):
    #     stock_result = process_step5(...)

    results["itc_summary"] = _build_itc_summary(results)
    return results


def _build_itc_summary(results):
    summary = {
        "mrr_rows": 0,
        "mrr_pivot_rows": 0,
        "mrr_validation_issues": 0,
        "mrr_total_igst": 0,
        "expense_rows": 0,
        "eligible_expense_rows": 0,
        "ineligible_expense_rows": 0,
        "expense_validation_issues": 0,
        "eligible_expense_itc": 0,
        "ineligible_expense_itc": 0,
        "books_eligible_rows": 0,
        "books_ineligible_rows": 0,
        "books_eligible_itc": 0,
        "books_ineligible_itc": 0,
    }

    if "step1" in results["steps"]:
        s1 = results["steps"]["step1"]["summary"]
        summary.update({
            "mrr_rows": s1.get("combined_rows", 0),
            "mrr_pivot_rows": s1.get("pivot_rows", 0),
            "mrr_validation_issues": s1.get("validation_issues", 0),
            "mrr_total_igst": s1.get("total_igst", 0),
        })

    if "step2" in results["steps"]:
        s2 = results["steps"]["step2"]["summary"]
        eligible = results["steps"]["step2"]["eligible"]
        ineligible = results["steps"]["step2"]["ineligible"]
        summary.update({
            "expense_rows": s2.get("total_rows", 0),
            "eligible_expense_rows": s2.get("eligible_rows", 0),
            "ineligible_expense_rows": s2.get("ineligible_rows", 0),
            "expense_validation_issues": s2.get("wrong_entries", 0),
            "eligible_expense_itc": float(eligible["Total ITC"].sum()) if len(eligible) else 0,
            "ineligible_expense_itc": float(ineligible["Total ITC"].sum()) if len(ineligible) else 0,
        })

    if "step3" in results["steps"]:
        s3 = results["steps"]["step3"]["summary"]
        summary.update({
            "books_eligible_rows": s3.get("eligible_rows", 0),
            "books_ineligible_rows": s3.get("ineligible_rows", 0),
            "books_eligible_itc": s3.get("eligible_itc", 0),
            "books_ineligible_itc": s3.get("ineligible_itc", 0),
        })

    # if "step4" in results["steps"]:
    #     summary.update({...})

    return summary


def _build_all_errors_df(results):
    """Collect validation, processing, and reconciliation errors into one sheet."""
    rows = []

    for msg in results.get("errors", []):
        step = msg.split(" failed")[0].replace("Step ", "").strip() if msg.startswith("Step ") else ""
        rows.append({
            "Step": step or "—", "Source": "System", "Row": "", "CON": "",
            "GSTIN": "", "Invoice No.": "", "Category / Remark": "",
            "Error Type": "Processing Error", "Error Detail": msg,
        })

    if "step1" in results["steps"]:
        issues = results["steps"]["step1"].get("issues", pd.DataFrame())
        if issues is not None and not issues.empty:
            for _, r in issues.iterrows():
                rows.append({
                    "Step": "1", "Source": "MRR / MRRR Return",
                    "Row": r.get("Row", ""), "CON": r.get("CON", ""),
                    "GSTIN": "", "Invoice No.": "", "Category / Remark": r.get("Remark", ""),
                    "Error Type": "Validation", "Error Detail": r.get("Issues", ""),
                })

    if "step2" in results["steps"]:
        issues = results["steps"]["step2"].get("issues", pd.DataFrame())
        if issues is not None and not issues.empty:
            for _, r in issues.iterrows():
                rows.append({
                    "Step": "2", "Source": "Master Expenses",
                    "Row": r.get("Row", ""), "CON": r.get("CON", ""),
                    "GSTIN": "", "Invoice No.": "", "Category / Remark": r.get("Category", ""),
                    "Error Type": "Validation", "Error Detail": r.get("Issues", ""),
                })

    # if "step4" in results["steps"]:
    #     ... reconciliation errors ...

    if not rows:
        return pd.DataFrame([{
            "Step": "—", "Source": "—", "Row": "", "CON": "",
            "GSTIN": "", "Invoice No.": "", "Category / Remark": "",
            "Error Type": "—", "Error Detail": "No errors found — all checks passed.",
        }])

    return pd.DataFrame(rows)


def _format_data_sheet(writer, sheet_name, df, min_width=12, max_width=48, padding=4):
    worksheet = writer.sheets[sheet_name]
    for col_idx, col_name in enumerate(df.columns):
        sample = df[col_name].fillna("").astype(str).head(1000)
        data_len = int(sample.map(len).max()) if len(sample) else 0
        width = min(max(len(str(col_name)), data_len) + padding, max_width)
        width = max(width, min_width)
        col_low = str(col_name).lower()
        if any(k in col_low for k in ("date", "voucher", "narration", "particulars")):
            width = max(width, 18)
        worksheet.set_column(col_idx, col_idx, width)
    worksheet.freeze_panes(1, 0)


def export_to_excel(results) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        _build_all_errors_df(results).to_excel(writer, sheet_name="All_Errors", index=False)

        if "step1" in results["steps"]:
            s1 = results["steps"]["step1"]
            s1_pivot = s1["pivot"].copy()
            if "Taxable Value" in s1_pivot.columns:
                s1_pivot = s1_pivot.drop(columns=["Taxable Value"])
            s1_pivot.to_excel(writer, sheet_name="S1_MRR_Pivot", index=False)
            _format_data_sheet(writer, "S1_MRR_Pivot", s1_pivot)
            s1["combined"].to_excel(writer, sheet_name="S1_MRR_Combined", index=False)
            _format_data_sheet(writer, "S1_MRR_Combined", s1["combined"])

        if "step2" in results["steps"]:
            s2_expenses = results["steps"]["step2"]["all_expenses"]
            s2_expenses.to_excel(writer, sheet_name="S2_All_Expenses", index=False)
            _format_data_sheet(writer, "S2_All_Expenses", s2_expenses)

        if "step3" in results["steps"]:
            s3 = results["steps"]["step3"]
            s3_combined = s3.get("books_combined")
            if s3_combined is None:
                s3_combined = pd.DataFrame()
            s3_combined.to_excel(writer, sheet_name="S3_Books_Combined", index=False)
            if not s3_combined.empty:
                _format_data_sheet(writer, "S3_Books_Combined", s3_combined)
            s3_ineligible = s3.get("ineligible")
            if s3_ineligible is None:
                s3_ineligible = pd.DataFrame()
            if not s3_ineligible.empty:
                s3_ineligible.to_excel(writer, sheet_name="ineligible", index=False)
                _format_data_sheet(writer, "ineligible", s3_ineligible)
            elif s3_ineligible is not None:
                s3_ineligible.to_excel(writer, sheet_name="ineligible", index=False)

        # if "step4" in results["steps"]:
        #     ...
        # if "step5" in results["steps"]:
        #     ...

        itc = results.get("itc_summary", {})
        pd.DataFrame([{"Metric": k, "Value": v} for k, v in itc.items()]).to_excel(
            writer, sheet_name="ITC_Summary", index=False
        )

    output.seek(0)
    return output.getvalue()
