"""ITC Reconciliation dashboard — Steps 1–3 active; Steps 4–5 commented out."""
import streamlit as st
import pandas as pd
from processors.engine import run_full_reconciliation, export_to_excel, STEPS_INFO
from saas_auth import log_download, show_announcements_banner


def _format_inr(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "Rs.0"
    return f"Rs.{val:,.2f}"


def show_itc_dashboard(user: dict):
    """Render the ITC reconciliation module."""
    show_announcements_banner()

    st.markdown("""
<div style="background:linear-gradient(135deg,#1f2f60 0%,#2d4a8a 100%);
            padding:1.2rem 1.5rem;border-radius:10px;color:#fff;margin-bottom:1.2rem;">
    <h2 style="color:#fff;margin:0;font-size:1.5rem;">ITC Reconciliation</h2>
    <p style="color:#c8dff5;margin:0.3rem 0 0 0;font-size:0.95rem;">
        Modicare GST Input Tax Credit — Step 1 (MRR), Step 2 (Expenses) &amp; Step 3 (Books)
    </p>
</div>
""", unsafe_allow_html=True)

    with st.sidebar:
        st.header("ITC Steps")
        # reco_mode = st.radio(
        #     "Books source (Step 3)",
        #     ["Build from Step 1 + Step 2", "Use existing Books file"],
        #     key="itc_reco_mode",
        # )
        # use_existing_books = reco_mode == "Use existing Books file"
        # month_filter = st.text_input(
        #     "Filter month (YYYY-MM, optional)", value="", key="itc_month_filter"
        # )
        with st.expander("Steps Guide", expanded=False):
            for step in STEPS_INFO:
                st.markdown(f"**Step {step['step']}: {step['title']}**")
                st.caption(step["description"])

    tab_upload, tab_results, tab_steps = st.tabs(["Upload Files", "ITC Results", "Step Details"])

    with tab_upload:
        st.subheader("Upload Source Files")
        st.info(
            "Currently active: **Step 1** (MRR + MRRR Return) and **Step 2** (Master Expenses). "
            "Step 3 (Books) is built automatically from Step 1 + 2. Steps 4–5 are disabled."
        )

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Step 1 — MRR & MRR Return**")
            mrr_file = st.file_uploader("1. GST MRR (CSV)", type=["csv"], key="itc_mrr")
            mrr_return_file = st.file_uploader("2. MRRR Return (CSV)", type=["csv"], key="itc_mrr_return")
            st.caption("3. Console MRR — reference only (processing uses CSV files)")

            st.markdown("**Step 2 — Master Expenses**")
            expense_file = st.file_uploader("4. Master Expenses (Excel)", type=["xlsx"], key="itc_expenses")

        with col2:
            # st.markdown("**Step 3 — Books (optional)**")
            # books_file = st.file_uploader(
            #     "5. Books Final (Excel)", type=["xlsx"], key="itc_books"
            # )
            pass

        has_step1 = mrr_file is not None and mrr_return_file is not None
        has_step2 = expense_file is not None
        ready_count = sum([has_step1, has_step2])
        st.progress(ready_count / 2, text=f"{ready_count} of 2 steps ready")

        can_run = has_step1 or has_step2

        if st.button("Run ITC Processing", type="primary", disabled=not can_run, use_container_width=True):
            with st.spinner("Processing Steps 1–3..."):
                files = {}
                if has_step1:
                    files["mrr"] = mrr_file
                    files["mrr_return"] = mrr_return_file
                if has_step2:
                    files["expenses"] = expense_file
                results = run_full_reconciliation(files)
                st.session_state["itc_results"] = results
                st.session_state["itc_processed"] = True

            if results.get("errors"):
                for err in results["errors"]:
                    st.error(err)
            elif results.get("steps"):
                st.success("Steps 1–3 processing completed!")
            else:
                st.warning("No steps were processed.")

            for key in ["itc_mrr", "itc_mrr_return", "itc_expenses"]:
                st.session_state.pop(key, None)

        if not can_run:
            st.warning("Upload **MRR + MRRR Return** (Step 1) and/or **Master Expenses** (Step 2) to run.")

    with tab_results:
        if not st.session_state.get("itc_processed"):
            st.info("Upload files and click **Run ITC Processing** to see results here.")
        else:
            results = st.session_state["itc_results"]
            itc = results.get("itc_summary", {})

            st.subheader("Processing Summary")

            if "step1" in results["steps"]:
                st.markdown("**Step 1 — MRR**")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Combined Rows", f"{itc.get('mrr_rows', 0):,}")
                c2.metric("Pivot Rows", f"{itc.get('mrr_pivot_rows', 0):,}")
                c3.metric("Validation Issues", f"{itc.get('mrr_validation_issues', 0):,}")
                c4.metric("Total IGST", _format_inr(itc.get("mrr_total_igst", 0)))

            if "step2" in results["steps"]:
                st.markdown("**Step 2 — Expenses**")
                s2 = results["steps"]["step2"]["summary"]
                cats = s2.get("category_counts", {})
                remarks = s2.get("remark_counts", {})
                st.write(f"**Category breakdown:** {cats}")
                st.write(f"**Remark breakdown:** {remarks}")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Purchase Rows", f"{s2.get('purchase_rows', cats.get('Purchase', 0)):,}")
                c2.metric("Expense Rows", f"{s2.get('expense_rows', cats.get('Expense', 0)):,}")
                c3.metric("Validation Issues", f"{s2.get('wrong_entries', 0):,}")
                c4.metric("Eligible ITC", _format_inr(itc.get("eligible_expense_itc", 0)))

            # ── Step 4 reconciliation results COMMENTED OUT ──
            # if "step4" in results["steps"]:
            #     s4 = results["steps"]["step4"]["summary"]
            #     st.subheader("Reconciliation Breakdown")
            #     ...

            if "step3" in results["steps"]:
                st.markdown("**Step 3 — Books**")
                s3 = results["steps"]["step3"]["summary"]
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Eligible Books Rows", f"{s3.get('eligible_rows', 0):,}")
                c2.metric("Ineligible Books Rows", f"{s3.get('ineligible_rows', 0):,}")
                c3.metric("MRR in Books", f"{s3.get('mrr_rows', 0):,}")
                c4.metric("Eligible Books ITC", _format_inr(s3.get("eligible_itc", 0)))

            st.divider()
            st.caption(
                "Excel sheets: **All_Errors**, **S1_MRR_Pivot**, **S1_MRR_Combined**, "
                "**S2_All_Expenses**, **S3_Books_Combined** (Expense + Purchase + MRR Return), "
                "**ineligible** (Expense-Ineligible debit), **ITC_Summary**."
            )
            excel_bytes = export_to_excel(results)
            fname = "ITC_Step1_Step3_Report.xlsx"
            if st.download_button(
                label="Download Step 1–3 Report (Excel)",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="itc_download",
            ):
                log_download(user.get("username", ""), fname)

    with tab_steps:
        if not st.session_state.get("itc_processed"):
            st.info("Run processing first to see step-by-step details.")
        else:
            results = st.session_state["itc_results"]

            if results.get("errors"):
                st.error("Errors encountered:")
                for err in results["errors"]:
                    st.write(f"- {err}")

            step_labels = {
                "step1": ("Step 1: MRR & MRR Return (compile + pivot)", "pivot"),
                "step2": ("Step 2: Master Expenses (categorize + validate)", "all_expenses"),
                "step3": ("Step 3: Books — Combined (Expense + MRR Return)", "books_combined"),
                # ── Steps 4–5 COMMENTED OUT ──
                # "step4": ("Step 4: GSTR-2B Reconciliation", "books_reconciled"),
                # "step5": ("Step 5: Stock Received + GIT + GSTR-2B updates", "pivot"),
            }

            for key, (label, data_key) in step_labels.items():
                if key in results["steps"]:
                    step_data = results["steps"][key]
                    summary = step_data.get("summary", {})

                    with st.expander(label, expanded=True):
                        for k, v in summary.items():
                            if isinstance(v, dict):
                                st.write(f"{k}: {v}")
                            elif isinstance(v, float):
                                st.write(f"{k}: {v:,.2f}")
                            elif isinstance(v, int):
                                st.write(f"{k}: {v:,}")
                            else:
                                st.write(f"{k}: {v}")

                    df = step_data.get(data_key)
                    if df is not None and not df.empty:
                        st.dataframe(df.head(50), use_container_width=True)

                    if key == "step3":
                        inel_df = step_data.get("ineligible")
                        if inel_df is not None and not inel_df.empty:
                            st.subheader("Step 3 — ineligible (Expense-Ineligible)")
                            st.dataframe(inel_df.head(50), use_container_width=True)

                    if key == "step2" and "ineligible" in step_data:
                        inel = step_data["ineligible"]
                        if inel is not None and not inel.empty:
                            st.markdown(
                                f"**Ineligible ITC (Debit only):** {len(inel)} rows — "
                                "included in **S2_All_Expenses** with Remark = Ineligible."
                            )

                    if "issues" in step_data and not step_data["issues"].empty:
                        st.warning(f"Validation issues: {len(step_data['issues'])}")
                        st.dataframe(step_data["issues"], use_container_width=True)
                else:
                    st.write(f"{label} — Not processed (files not uploaded)")
