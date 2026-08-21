"""GSTR-3B dashboard — Steps 1–4 (output liability, ITC, set-off, table mapping)."""
import datetime as _dt

import pandas as pd
import streamlit as st

from processors import tmpdir_setup
from processors.gstr3b_config import DEFAULT_IGST_SPLIT
from processors.gstr3b_engine import (GSTR3B_STEPS_INFO, exceptions_frame,
                                      export_gstr3b_excel, gstr3b_frame,
                                      portal_view, run_gstr3b, step1_frame,
                                      step2_frame, step3_frame)
from saas_auth import log_download, show_announcements_banner

# Keep Excel export off the system drive — see processors/tmpdir_setup.py.
tmpdir_setup.use_project_tmpdir()
tmpdir_setup.purge(older_than_seconds=3600)

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _format_inr(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "Rs.0"
    return f"Rs.{val:,.2f}"


def show_gstr3b_dashboard(user: dict):
    """Render the GSTR-3B module."""
    show_announcements_banner()

    st.markdown("""
<div style="background:linear-gradient(135deg,#1f2f60 0%,#2d4a8a 100%);
            padding:1.2rem 1.5rem;border-radius:10px;color:#fff;margin-bottom:1.2rem;">
    <h2 style="color:#fff;margin:0;font-size:1.5rem;">GSTR-3B Preparation</h2>
    <p style="color:#c8dff5;margin:0.3rem 0 0 0;font-size:0.95rem;">
        Modicare GST &mdash; Step 1 (Output Liability), Step 2 (ITC),
        Step 3 (Set-off &amp; Cash) &amp; Step 4 (Table Mapping)
    </p>
</div>
""", unsafe_allow_html=True)

    with st.sidebar:
        st.header("GSTR-3B Settings")
        _now = _dt.datetime.now()
        col_m, col_y = st.columns(2)
        with col_m:
            sel_month = st.selectbox("Month", _MONTHS, index=_now.month - 1,
                                     key="g3b_month")
        with col_y:
            years = list(range(2020, 2031))
            sel_year = st.selectbox("Year", years,
                                    index=years.index(_now.year) if _now.year in years else 0,
                                    key="g3b_year")
        period = f"{_MONTHS.index(sel_month) + 1:02d}{sel_year}"
        st.caption(f"Tax period: **{period}** — drives the 'MONTH OF ITC TAKEN' filter")

        igst_split = st.slider(
            "Residual IGST credit to CGST", 0.0, 1.0, DEFAULT_IGST_SPLIT, 0.05,
            key="g3b_split",
            help="Rule 88A allows residual IGST credit to go to CGST and SGST in any "
                 "order and proportion. 0.50 mirrors the existing working.",
        )
        st.caption(f"CGST {igst_split:.0%} / SGST {1 - igst_split:.0%}")

        with st.expander("Steps Guide", expanded=False):
            for step in GSTR3B_STEPS_INFO:
                st.markdown(f"**Step {step['step']}: {step['title']}**")
                st.caption(step["description"])

    tab_upload, tab_results, tab_gstin, tab_steps = st.tabs(
        ["Upload Files", "GSTR-3B Results", "Per GSTIN", "Step Details"])

    # ── UPLOAD ──
    with tab_upload:
        st.subheader("Upload Source Files")
        st.caption("Upload the monthly working files. All files are required — "
                   "the run button stays disabled until every one is in.")

        # Order, labels and required markers mirror the standalone GSTR-3B app
        # exactly, so the same file goes in the same place in either one.
        _U = lambda label, key, help=None: st.file_uploader(
            label, type=["xlsx", "xlsm", "xls"], key=key, help=help)

        col1, col2 = st.columns(2)
        with col1:
            sale_file = _U("Sale Summary (combined GSTR-1) *", "g3b_sale")
            books25_file = _U("Books FY 2025-26 (ITC as per books) *", "g3b_books25")
            books26_file = _U("Books FY 2026-27 (ITC as per books) *", "g3b_books26")
            stock_file = _U("Stock Received summary *", "g3b_stock")
            isd_file = _U("ISD distribution *", "g3b_isd")
            expense_file = _U("Final Expenses (RCM) *", "g3b_expense")
            rev42a_file = _U(
                "Rule 42 expenses FY 25-26 (Reversal on Exempted Sales) *",
                "g3b_rev42a",
                "'Expenses 25-26.xlsx'. Without it the Rule 42 reversal falls "
                "back to the books subtotal, which understates CGST/SGST.")
        with col2:
            rev42b_file = _U(
                "Rule 42 expenses FY 26-27 (Reversal on Exempted Sales) *",
                "g3b_rev42b", "'Expenses 26-27.xlsx'.")
            g2b_file = _U(
                "GSTR-2B — current FY (holds the filing month) *", "g3b_2b",
                "The FY carrying the filing month itself, e.g. "
                "'GSTR-2B FY 26-27_May.xlsx' for 052026. Supplies the "
                "current-month half of table 4A(5).")
            g2b_prev_file = _U(
                "GSTR-2B — earlier FY (reclaim rows) *", "g3b_2b_prev",
                "The previous financial year's 2B, e.g. "
                "'GSTR-2B FY 25-26_May.xlsx'. Reclaim rows and the opening "
                "ledger come from here.")
            liab_file = _U("Liability Summary (cross charge + PMT balance) *",
                           "g3b_liab")
            pmt_file = _U(
                "PMT Balance (opening credit ledger) *", "g3b_pmt",
                "Standalone PMT workbook. Left empty, the 'PMT Balance' sheet "
                "inside the Liability Summary is used instead.")
            itc_summary_file = _U(
                "ITC as per books — summary (cross-check) *", "g3b_itc_summary",
                "State / IGST / CGST / SGST / Cess / Remarks. Overrides the "
                "'a. ITC FY 25-26' and 'b. ITC FY 26-27' figures computed from "
                "the raw Books files.")

        uploaded = [f for f in (sale_file, books25_file, books26_file, itc_summary_file,
                                stock_file, isd_file, expense_file, liab_file,
                                g2b_file, g2b_prev_file, rev42a_file, rev42b_file,
                                pmt_file)
                    if f is not None]
        st.progress(len(uploaded) / 13, text=f"{len(uploaded)} of 13 files uploaded")

        # Every file is required - the run button stays disabled until all are
        # in. Same rule as the standalone app.
        _missing = [lbl for lbl, f in (
            ("Sale Summary (combined GSTR-1)", sale_file),
            ("Books FY 2025-26", books25_file),
            ("Books FY 2026-27", books26_file),
            ("Stock Received summary", stock_file),
            ("ISD distribution", isd_file),
            ("Final Expenses (RCM)", expense_file),
            ("Rule 42 expenses FY 25-26", rev42a_file),
            ("Rule 42 expenses FY 26-27", rev42b_file),
            ("GSTR-2B — current FY", g2b_file),
            ("GSTR-2B — earlier FY", g2b_prev_file),
            ("Liability Summary", liab_file),
            ("PMT Balance", pmt_file),
            ("ITC as per books — summary", itc_summary_file),
        ) if f is None]
        if _missing:
            st.warning("**Still needed:** " + ", ".join(_missing))
        can_run = not _missing
        if st.button("Run GSTR-3B Processing", type="primary",
                     disabled=not can_run, use_container_width=True):
            with st.spinner("Processing Steps 1–4..."):
                files = {
                    "sale": sale_file, "books25": books25_file, "books26": books26_file,
                    "itc_summary": itc_summary_file,
                    "stock": stock_file, "isd": isd_file, "expense": expense_file,
                    "liab": liab_file, "gstr2b": g2b_file,
                    "gstr2b_prev": g2b_prev_file,
                    "rev42a": rev42a_file, "rev42b": rev42b_file,
                    "pmt": pmt_file,
                }
                results = run_gstr3b({k: v for k, v in files.items() if v is not None},
                                     period=period, igst_split=igst_split)
                st.session_state["g3b_results"] = results
                st.session_state["g3b_processed"] = True

            if results.get("errors"):
                for err in results["errors"]:
                    st.error(err)
            if results.get("steps"):
                st.success(f"Processed {len(results['steps'])} step(s) for period {period}.")
            else:
                st.warning("No steps were processed.")

        if not can_run:
            st.warning("Upload **Sale Summary** to run — Step 1 cannot proceed without it.")

    # ── RESULTS ──
    with tab_results:
        if not st.session_state.get("g3b_processed"):
            st.info("Upload files and click **Run GSTR-3B Processing** to see results here.")
        else:
            results = st.session_state["g3b_results"]
            summ = results.get("gstr3b_summary", {})
            steps = results.get("steps", {})

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Net Taxable Sales", _format_inr(summ.get("net_taxable_sales", 0)))
            c2.metric("Net Exempt Sales", _format_inr(summ.get("net_exempt_sales", 0)))
            c3.metric("Output Liability", _format_inr(
                summ.get("output_igst", 0) + summ.get("output_cgst", 0)
                + summ.get("output_sgst", 0)))
            c4.metric("ITC Available", _format_inr(summ.get("itc_available_total", 0)))
            c5.metric("Cash Payable", _format_inr(summ.get("cash_payable_total", 0)))

            n_err, n_warn = summ.get("errors", 0), summ.get("warnings", 0)
            if n_err:
                st.error(f"{n_err} error(s) and {n_warn} warning(s) — review the "
                         "**Step Details** tab before filing.")
            elif n_warn:
                st.warning(f"{n_warn} warning(s) — see the **Step Details** tab.")
            else:
                st.success("No errors or warnings raised.")

            if "step4" in steps:
                st.subheader("GSTR-3B Figures, State-wise")
                st.caption(
                    "Verify tables 3.1/3.2 against the portal's auto-populated figures. "
                    "Where auto-population is hard-locked, corrections must go through "
                    "GSTR-1A before filing 3B."
                )
                st.dataframe(gstr3b_frame(steps["step4"]["tables"]),
                             use_container_width=True, height=430)

            st.divider()
            st.caption(
                "Excel sheets: **Run_Log**, **Step1_Liability**, **Step2_ITC**, "
                "**Step3_Cash**, **Step4_GSTR-3B**, **Exceptions**."
            )
            excel_bytes = export_gstr3b_excel(results)
            fname = f"GSTR3B_Computation_{results.get('period', '')}.xlsx"
            if st.download_button(
                label="Download GSTR-3B Report (Excel)",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="g3b_download",
            ):
                log_download(user.get("username", ""), fname)

    # ── PER GSTIN ──
    with tab_gstin:
        if not st.session_state.get("g3b_processed"):
            st.info("Run processing first to see the per-GSTIN return view.")
        else:
            results = st.session_state["g3b_results"]
            tables = results.get("steps", {}).get("step4", {}).get("tables", {})
            if not tables:
                st.warning("Step 4 did not produce any tables.")
            else:
                state = st.selectbox("State", sorted(tables), key="g3b_state")
                st.caption(f"Return period {results.get('period', '')}")
                st.dataframe(portal_view(state, tables),
                             use_container_width=True, hide_index=True)

                cash = results.get("steps", {}).get("step3", {}).get("cash", {})
                if state in cash:
                    st.markdown("**Credit utilisation**")
                    used = cash[state]["utilised"]
                    st.dataframe(
                        pd.DataFrame([{"Movement": k.replace("_", " ").upper(),
                                       "Amount": round(v, 2)}
                                      for k, v in used.items() if abs(v) > 0.005]),
                        use_container_width=True, hide_index=True)

    # ── STEP DETAILS ──
    with tab_steps:
        if not st.session_state.get("g3b_processed"):
            st.info("Run processing first to see step-by-step details.")
        else:
            results = st.session_state["g3b_results"]
            steps = results.get("steps", {})

            if results.get("errors"):
                st.error("Errors encountered:")
                for err in results["errors"]:
                    st.write(f"- {err}")

            with st.expander("Source files loaded", expanded=False):
                for line in results.get("load_log", []):
                    st.write(f"- {line}")

            step_labels = {
                "step1": "Step 1: Output Tax Liability",
                "step2": "Step 2: ITC Availability",
                "step3": "Step 3: Set-off & Cash Liability",
                "step4": "Step 4: GSTR-3B Table Mapping",
            }
            frame_for = {
                "step1": lambda d: step1_frame(d["grid"]),
                "step2": lambda d: step2_frame(d["itc"]),
                "step3": lambda d: step3_frame(d["cash"]),
                "step4": lambda d: gstr3b_frame(d["tables"]),
            }

            for key, label in step_labels.items():
                if key not in steps:
                    st.write(f"{label} — Not processed (files not uploaded)")
                    continue
                data = steps[key]
                with st.expander(label, expanded=(key == "step1")):
                    for k, v in data.get("summary", {}).items():
                        if isinstance(v, bool):
                            st.write(f"{k}: {v}")
                        elif isinstance(v, float):
                            st.write(f"{k}: {v:,.2f}")
                        elif isinstance(v, int):
                            st.write(f"{k}: {v:,}")
                        else:
                            st.write(f"{k}: {v}")
                    df = frame_for[key](data)
                    if df is not None and not df.empty:
                        st.dataframe(df.head(50), use_container_width=True)

            exc = results.get("exceptions", [])
            if exc:
                st.subheader("Validation Exceptions")
                exc_df = exceptions_frame(exc)
                sev = st.multiselect("Severity", ["ERROR", "WARN", "INFO"],
                                     default=["ERROR", "WARN"], key="g3b_sev")
                st.dataframe(exc_df[exc_df["Severity"].isin(sev)],
                             use_container_width=True, height=380)
