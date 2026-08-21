"""GSTR-3B orchestrator — Steps 1–4, Excel export, step metadata."""
import io

import pandas as pd

from . import gstr3b_config as G
from .gstr3b_steps import (compliance_checks, parse_period, process_itc_summary,
                           process_step1, process_step2, process_step3,
                           process_step4, read_table, reclaim_opening,
                           rule37_movements)

GSTR3B_STEPS_INFO = [
    {
        "step": 1,
        "title": "Output Tax Liability",
        "files": ["Sale Summary (Excel)"],
        "description": (
            "Tag every line Exempt (slab 0) or Taxable, pivot state-wise across "
            "Sale / Stock Transfer / Sale Return / Asset Sale / Cross Charge, "
            "derive net taxable, net exempt and the IGST/CGST/SGST liability."
        ),
    },
    {
        "step": 2,
        "title": "ITC Availability",
        "files": ["Books FY 25-26 & 26-27 (or ITC-as-per-books summary)", "Stock Received",
                  "ISD", "Final Expenses", "Liability Summary (PMT balance)"],
        "description": (
            "Eight buckets — FY 25-26, FY 26-27, Import, Stock received, Cross charge, "
            "ISD, RCM, less Rule 42 reversal — filtered on 'MONTH OF ITC TAKEN', "
            "then added to the opening electronic credit ledger. The FY 25-26/26-27 "
            "buckets can be sourced either from the raw Books workbooks or from a "
            "single pre-aggregated 'ITC as per books' state-wise summary, which "
            "overrides the raw-file figures when uploaded."
        ),
    },
    {
        "step": 3,
        "title": "Set-off & Cash Liability",
        "files": ["Derived from Steps 1 + 2"],
        "description": (
            "Utilise credit per s.49(5), 49A, 49B and Rule 88A — IGST first, residual "
            "IGST to CGST/SGST, no CGST/SGST cross-utilisation. RCM is added as cash."
        ),
    },
    {
        "step": 4,
        "title": "GSTR-3B Table Mapping",
        "files": ["GSTR-2B (Excel) — optional"],
        "description": (
            "Map onto portal tables 3.1(a), 3.1(c), 3.1(d), 4A(1), 4A(3), 4A(4), "
            "4A(5), 4B(1), 4B(2), 4C and 6.1."
        ),
    },
]

# (upload key, sheet-name hints, header key tokens, minimum column count)
#
# min_cols guards against picking a same-named but wrong sheet. The ISD
# workbook, for example, holds both a 'Turnover' summary and a month sheet;
# only the month sheet carries the distributed columns O/P/Q that Step 2 needs.
GSTR3B_SOURCES = {
    "sale":        (["Sale Summary"], G.SALE_KEYS, 0),
    "books25":     (["25-26", "Sheet1"], G.BOOKS_KEYS, 0),
    "books26":     (["26-27", "25-26", "Sheet1"], G.BOOKS_KEYS, 0),
    "stock":       (["Stock Recd_Summary"], G.STOCK_KEYS, 0),
    "isd":         ([], G.ISD_KEYS, G.ISD_POS["sgst"] + 1),
    "expense":     (["Final", "RCM ONLY"], G.EXPENSE_KEYS, G.EXPENSE_POS["pay_igst"] + 1),
    "gstr2b":      (["GSTR-2B"], G.GSTR2B_KEYS, 0),
    # The earlier FY's 2B. Reclaim rows for 4A(5) sit in it - the current-FY
    # workbook only carries the filing month and the one or two before it.
    "gstr2b_prev": (["GSTR-2B"], G.GSTR2B_KEYS, 0),
    "liab":        (["Cross charge"], G.CROSS_KEYS, 0),
    "itc_summary": (["Sheet1"], G.ITC_SUMMARY_KEYS, 0),
    # Rule 42 expense pivots, one per FY, from 'Reversal on Exempted Sales'.
    "rev42a":      (["Pivot"], G.REV42_EXPENSE_KEYS, 0),
    "rev42b":      (["Pivot"], G.REV42_EXPENSE_KEYS, 0),
    # Standalone PMT balance; the same table also rides in the liability book.
    "pmt":         (["Sheet1", "PMT Balance"], G.PMT_KEYS, 0),
}

_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _candidate_sheets(src, hints):
    """Sheet names ordered by preference: exact hint, partial hint, then the rest."""
    try:
        names = list(pd.ExcelFile(src).sheet_names)
    except Exception:
        return []
    ordered, seen = [], set()

    def add(n):
        if n not in seen:
            seen.add(n)
            ordered.append(n)

    for h in hints:
        for n in names:
            if str(n).strip().lower() == str(h).strip().lower():
                add(n)
    for h in hints:
        for n in names:
            if str(h).strip().lower() in str(n).strip().lower():
                add(n)
    for n in names:
        add(n)
    return ordered


def _load(src, hints, tokens, min_cols=0):
    """
    Try each candidate sheet until one parses AND satisfies min_cols.

    Returns (dataframe, sheet_name, header_row_index). Raises when no sheet
    in the workbook yields a usable table — the error names every sheet that
    was tried, so a wrong-file-in-wrong-slot upload is obvious immediately
    instead of surfacing whichever sheet happened to be tried last.
    """
    candidates = _candidate_sheets(src, hints)
    if not candidates:
        raise ValueError("workbook has no readable sheets")
    last_err = None
    for sheet in candidates:
        try:
            df, hdr = read_table(src, sheet, tokens)
        except Exception as e:
            last_err = e
            continue
        if len(df.columns) < min_cols:
            last_err = ValueError(
                "sheet '%s' has %d columns, needs at least %d"
                % (sheet, len(df.columns), min_cols))
            continue
        return df, sheet, hdr
    raise ValueError(
        "none of the %d sheet(s) in this workbook — %s — has a header row "
        "with any of %s. %s Check that the right file was uploaded to this slot."
        % (len(candidates), candidates, sorted(set(tokens)),
           last_err if last_err else ""))


# ─────────────────────────────────────────────────────────────────
# MAIN ENTRY
# ─────────────────────────────────────────────────────────────────
def run_gstr3b(files: dict, period="052026", igst_split=G.DEFAULT_IGST_SPLIT):
    """
    Run Steps 1–4. Only `sale` is mandatory; every other source is optional
    and the affected step degrades rather than failing.
    """
    results = {"steps": {}, "errors": [], "load_log": [], "exceptions": []}

    try:
        month, year = parse_period(period)
    except ValueError as e:
        results["errors"].append(str(e))
        return results

    # ── load ──
    # The ISD and stock workbooks name their working sheet after the month
    # ("May'26"), so the period drives the sheet hint before the static ones.
    mon = _MONTH_ABBR[month - 1]
    month_hints = ["%s'%s" % (mon, str(year)[-2:]), "%s-%s" % (mon, str(year)[-2:]), mon]

    frames = {}
    for key, (hints, tokens, min_cols) in GSTR3B_SOURCES.items():
        upload = files.get(key)
        if upload is None:
            frames[key] = None
            results["load_log"].append("%s: not uploaded" % key)
            continue
        try:
            use_hints = (month_hints + list(hints)) if key in ("isd", "stock") else hints
            df, sheet, hdr = _load(upload, use_hints, tokens, min_cols)
            frames[key] = df
            results["load_log"].append(
                "%s: '%s', header row %d, %d rows" % (key, sheet, hdr + 1, len(df)))
        except Exception as e:
            frames[key] = None
            results["load_log"].append("%s: FAILED — %s" % (key, e))
            results["errors"].append("Could not read %s: %s" % (key, e))

    if frames.get("sale") is None:
        results["errors"].append(
            "Sale Summary is required — Step 1 cannot run without it.")
        return results

    # ── Step 1 ──
    try:
        s1 = process_step1(frames["sale"])
        results["steps"]["step1"] = s1
        results["exceptions"] += s1["exceptions"]
    except Exception as e:
        results["errors"].append("Step 1 failed: %s" % e)
        return results
    grid = s1["grid"]

    # ── Step 2 ──
    books = {}
    if frames.get("books25") is not None:
        books["FY 25-26"] = frames["books25"]
    if frames.get("books26") is not None:
        books["FY 26-27"] = frames["books26"]

    # A standalone PMT workbook wins when given - it is the more specific
    # source. Otherwise the same table rides along in the liability summary.
    opening = frames.get("pmt")
    if opening is None and files.get("liab") is not None:
        try:
            opening, _ = read_table(files["liab"], "PMT Balance", G.PMT_KEYS)
        except Exception:
            results["load_log"].append("liab: 'PMT Balance' sheet not found — "
                                       "opening credit ledger treated as nil")

    itc_summary = None
    if frames.get("itc_summary") is not None:
        a_sum, b_sum, exc_sum = process_itc_summary(frames["itc_summary"])
        results["exceptions"] += exc_sum
        itc_summary = (a_sum, b_sum)

    itc = {}
    try:
        _rev42 = [f for f in (frames.get("rev42a"), frames.get("rev42b"))
                  if f is not None and len(f)]
        s2 = process_step2(grid, books, frames.get("stock"), frames.get("liab"),
                           frames.get("isd"), frames.get("expense"), opening,
                           month, year, itc_summary=itc_summary,
                           rev42_expense=_rev42 or None)
        results["steps"]["step2"] = s2
        results["exceptions"] += s2["exceptions"]
        itc = s2["itc"]
    except Exception as e:
        results["errors"].append("Step 2 failed: %s" % e)

    # ── Step 3 ──
    cash = {}
    try:
        s3 = process_step3(grid, itc, igst_split=igst_split)
        results["steps"]["step3"] = s3
        results["exceptions"] += s3["exceptions"]
        cash = s3["cash"]
    except Exception as e:
        results["errors"].append("Step 3 failed: %s" % e)

    # ── Step 4 ──
    try:
        # Both financial years feed table 4A(5): the current FY carries the
        # filing month, the earlier one the reclaim rows. Concatenating keeps
        # step 4 unchanged - it filters on MONTH either way.
        _g2b = [f for f in (frames.get("gstr2b"), frames.get("gstr2b_prev"))
                if f is not None and len(f)]
        _g2b_all = (pd.concat(_g2b, ignore_index=True) if len(_g2b) > 1
                    else (_g2b[0] if _g2b else None))
        # The opening reclaim ledger is derived from the GSTR-2B sheets - earlier
        # months, 'Considered', with 'TAKEN IN 3B' blank, naming this month, or
        # carrying a Split remark. No separate opening workbook is needed.
        _r37 = rule37_movements(files.get("liab"))
        s4 = process_step4(grid, itc, cash, _g2b_all, period=period,
                           rule37=_r37)
        if _g2b_all is None:
            # Keys are Title-cased here - the exception frame is built from
            # 'Severity'/'Check'/'State'/'Detail', and lowercase ones crash it.
            s4["exceptions"].append({
                "Severity": "WARN", "Check": "Reclaim opening balance",
                "State": "—",
                "Detail": "No GSTR-2B loaded, so the reclaim ledger opens at nil."})
        results["steps"]["step4"] = s4
        results["exceptions"] += s4["exceptions"]
    except Exception as e:
        results["errors"].append("Step 4 failed: %s" % e)

    try:
        results["exceptions"] += compliance_checks(grid, itc, cash)
    except Exception as e:
        results["errors"].append("Compliance checks failed: %s" % e)

    results["period"] = period
    results["gstr3b_summary"] = _build_gstr3b_summary(results)
    return results


def _build_gstr3b_summary(results):
    summary = {
        "states": 0,
        "net_taxable_sales": 0, "net_exempt_sales": 0,
        "output_igst": 0, "output_cgst": 0, "output_sgst": 0,
        "itc_available_total": 0, "rule42_reversal_total": 0,
        "cash_on_supply_total": 0, "rcm_cash_total": 0, "cash_payable_total": 0,
        "errors": 0, "warnings": 0,
    }
    steps = results.get("steps", {})
    if "step1" in steps:
        summary.update({k: v for k, v in steps["step1"]["summary"].items()
                        if k in summary or k == "states"})
    if "step2" in steps:
        s2 = steps["step2"]["summary"]
        summary["itc_available_total"] = s2.get("itc_available_total", 0)
        summary["rule42_reversal_total"] = s2.get("rule42_reversal_total", 0)
    if "step3" in steps:
        s3 = steps["step3"]["summary"]
        summary["cash_on_supply_total"] = s3.get("cash_on_supply_total", 0)
        summary["rcm_cash_total"] = s3.get("rcm_cash_total", 0)
        summary["cash_payable_total"] = s3.get("cash_payable_total", 0)
    exc = results.get("exceptions", [])
    summary["errors"] = sum(1 for e in exc if e["Severity"] == "ERROR")
    summary["warnings"] = sum(1 for e in exc if e["Severity"] == "WARN")
    return summary


# ─────────────────────────────────────────────────────────────────
# DATAFRAME BUILDERS
# ─────────────────────────────────────────────────────────────────
def _with_total(df):
    if not len(df):
        return df
    tot = {}
    for c in df.columns:
        tot[c] = round(float(df[c].sum()), 2) if df[c].dtype.kind in "fi" else ""
    tot[df.columns[0]] = "TOTAL"
    return pd.concat([df, pd.DataFrame([tot])], ignore_index=True)


def step1_frame(grid):
    rows = []
    for s in sorted(grid):
        r = {"State": s, "Code": G.CODE_BY_STATE.get(s, "")}
        for c in G.STEP1_COLS:
            r["%s - %s" % (c, G.STEP1_LABELS[c])] = round(grid[s][c], 2)
        rows.append(r)
    return _with_total(pd.DataFrame(rows))


def step2_frame(itc):
    buckets = [("a_2526", "a. ITC FY 25-26"), ("b_2627", "b. ITC FY 26-27"),
               ("c_import", "c. Import"), ("d_stock", "d. Stock received"),
               ("e_cross", "e. Cross charge"), ("f_isd", "f. ISD"),
               ("g_rcm", "g. RCM"), ("h_rev42", "h. Rule 42 reversal"),
               ("current", "Current month net"), ("opening", "Opening ECL"),
               ("available", "Total available")]
    rows = []
    for s in sorted(itc):
        r = {"State": s, "Exempt ratio %": round(itc[s].get("exempt_ratio", 0) * 100, 4)}
        for key, label in buckets:
            for h in ("igst", "cgst", "sgst"):
                r["%s - %s" % (label, G.HEAD_LABEL[h])] = round(itc[s][key][h], 2)
        rows.append(r)
    return pd.DataFrame(rows)


def step3_frame(cash):
    rows = []
    for s in sorted(cash):
        c = cash[s]
        r = {"State": s}
        for label, key in (("Liability", "liability"), ("ITC available", "available"),
                           ("Cash on supply", "cash_on_supply"),
                           ("RCM cash", "rcm_liability"), ("TOTAL CASH", "cash_total"),
                           ("Closing credit", "closing_credit")):
            for h in ("igst", "cgst", "sgst"):
                r["%s - %s" % (label, G.HEAD_LABEL[h])] = round(c[key].get(h, 0.0), 2)
        for k, v in c["utilised"].items():
            r["Util - %s" % k] = round(v, 2)
        rows.append(r)
    return _with_total(pd.DataFrame(rows))


def gstr3b_frame(tables):
    """One row per state, columns laid out as the portal tables."""
    rows = []
    for s in sorted(tables):
        t = tables[s]
        r = {"State": s, "Code": G.CODE_BY_STATE.get(s, "")}
        r["3.1(a) Taxable value"] = round(t["t3_1a"]["taxable"], 2)
        for h in G.HEADS:
            r["3.1(a) %s" % G.HEAD_LABEL[h]] = round(t["t3_1a"][h], 2)
        r["3.1(c) Nil/Exempt value"] = round(t["t3_1c"]["taxable"], 2)
        r["3.1(d) RCM taxable value"] = round(t["t3_1d"]["taxable"], 2)
        for h in ("igst", "cgst", "sgst"):
            r["3.1(d) %s" % G.HEAD_LABEL[h]] = round(t["t3_1d"][h], 2)
        # Imports are IGST-only - customs levies IGST, never CGST/SGST - so the
        # two always-zero columns under 4A(1) become the 4A(2) services line.
        r["4A(1) Import of goods IGST"] = round(t["t4a1"]["igst"], 2)
        r["4A(2) Import of services IGST"] = round(
            t.get("t4a2", {}).get("igst", 0.0), 2)
        for key, label in (("t4a3", "4A(3) Inward RCM"),
                           ("t4a4", "4A(4) ISD"), ("t4a5", "4A(5) All other ITC"),
                           ("t4b1", "4B(1) Rules 38/42/43 & 17(5)"),
                           ("t4b2", "4B(2) Others"), ("t4c", "4C Net ITC available"),
                           ("t4d1", "4D(1) ITC reclaimed (reversed earlier under 4B(2))"),
                           ("reclaim_open", "Reclaim ledger opening"),
                           ("reclaim_close", "Reclaim ledger closing"),
                           ("current_2b", "Current month 2B (Considered)")):
            for h in ("igst", "cgst", "sgst"):
                r["%s %s" % (label, G.HEAD_LABEL[h])] = round(t[key][h], 2)
        for h in ("igst", "cgst", "sgst"):
            r["6.1 Cash payable %s" % G.HEAD_LABEL[h]] = round(
                t["t6_1"]["cash_total"].get(h, 0.0), 2)
        rows.append(r)
    return _with_total(pd.DataFrame(rows))


def portal_view(state, tables):
    """The GSTR-3B form for one state, as it appears on the portal."""
    t = tables[state]
    rows = [
        ("3.1(a)", "Outward taxable supplies (other than zero rated, nil, exempt)",
         t["t3_1a"]["taxable"], t["t3_1a"]["igst"], t["t3_1a"]["cgst"],
         t["t3_1a"]["sgst"], t["t3_1a"]["cess"]),
        ("3.1(c)", "Other outward supplies (nil rated, exempted)",
         t["t3_1c"]["taxable"], 0, 0, 0, 0),
        ("3.1(d)", "Inward supplies liable to reverse charge",
         t["t3_1d"]["taxable"], t["t3_1d"]["igst"], t["t3_1d"]["cgst"],
         t["t3_1d"]["sgst"], 0),
        ("4A(1)", "ITC — Import of goods", None,
         t["t4a1"]["igst"], t["t4a1"]["cgst"], t["t4a1"]["sgst"], t["t4a1"]["cess"]),
        ("4A(3)", "ITC — Inward supplies liable to reverse charge", None,
         t["t4a3"]["igst"], t["t4a3"]["cgst"], t["t4a3"]["sgst"], 0),
        ("4A(4)", "ITC — Inward supplies from ISD", None,
         t["t4a4"]["igst"], t["t4a4"]["cgst"], t["t4a4"]["sgst"], 0),
        ("4A(5)", "ITC — All other ITC", None,
         t["t4a5"]["igst"], t["t4a5"]["cgst"], t["t4a5"]["sgst"], t["t4a5"]["cess"]),
        ("4B(1)", "ITC reversed — rules 38, 42, 43 and section 17(5)", None,
         t["t4b1"]["igst"], t["t4b1"]["cgst"], t["t4b1"]["sgst"], 0),
        ("4B(2)", "ITC reversed — others", None,
         t["t4b2"]["igst"], t["t4b2"]["cgst"], t["t4b2"]["sgst"], 0),
        ("4C", "Net ITC available (4A − 4B)", None,
         t["t4c"]["igst"], t["t4c"]["cgst"], t["t4c"]["sgst"], t["t4c"]["cess"]),
        ("6.1", "Tax payable in CASH (incl. RCM)", None,
         t["t6_1"]["cash_total"].get("igst", 0), t["t6_1"]["cash_total"].get("cgst", 0),
         t["t6_1"]["cash_total"].get("sgst", 0), t["t6_1"]["cash_total"].get("cess", 0)),
    ]
    return pd.DataFrame(rows, columns=["Table", "Particulars", "Taxable value",
                                       "IGST", "CGST", "SGST", "Cess"])


def exceptions_frame(exceptions):
    if not exceptions:
        return pd.DataFrame([{
            "Severity": "—", "Check": "—", "State": "—",
            "Detail": "No exceptions raised — all checks passed.",
        }])
    order = {"ERROR": 0, "WARN": 1, "INFO": 2}
    df = pd.DataFrame(exceptions)
    df["_o"] = df["Severity"].map(lambda v: order.get(v, 9))
    return df.sort_values(["_o", "Check", "State"]).drop(columns="_o").reset_index(drop=True)


def _build_all_errors_df(results):
    """Processing errors + validation exceptions in one sheet (mirrors ITC module)."""
    rows = []
    for msg in results.get("errors", []):
        step = msg.split(" failed")[0].replace("Step ", "").strip() \
            if msg.startswith("Step ") else ""
        rows.append({"Step": step or "—", "Severity": "ERROR", "State": "",
                     "Check": "Processing Error", "Detail": msg})
    for e in results.get("exceptions", []):
        rows.append({"Step": "", "Severity": e["Severity"], "State": e["State"],
                     "Check": e["Check"], "Detail": e["Detail"]})
    if not rows:
        return pd.DataFrame([{"Step": "—", "Severity": "—", "State": "—",
                              "Check": "—", "Detail": "No errors found."}])
    order = {"ERROR": 0, "WARN": 1, "INFO": 2}
    df = pd.DataFrame(rows)
    df["_o"] = df["Severity"].map(lambda v: order.get(v, 9))
    return df.sort_values(["_o", "Check"]).drop(columns="_o").reset_index(drop=True)


def _format_data_sheet(writer, sheet_name, df, min_width=12, max_width=48, padding=4):
    worksheet = writer.sheets[sheet_name]
    for col_idx, col_name in enumerate(df.columns):
        sample = df[col_name].fillna("").astype(str).head(1000)
        data_len = int(sample.map(len).max()) if len(sample) else 0
        width = min(max(len(str(col_name)), data_len) + padding, max_width)
        worksheet.set_column(col_idx, col_idx, max(width, min_width))
    worksheet.freeze_panes(1, 1)


def formula_map_frame():
    """
    Where every reported figure comes from: its formula, source workbook,
    sheet, and the columns and filters applied. Mirrors the standalone app.
    """
    rows = [
        ("Step1_Liability", "AA Taxable value", "J+N+S+V+CO",
         "Sale Summary", "Sale Summary", "positional columns"),
        ("Step1_Liability", "AB Exempt value", "R+U+Z",
         "Sale Summary", "Sale Summary", "positional columns"),
        ("Step1_Liability", "AC Total value", "AA + AB", "derived", "-", "-"),
        ("Step1_Liability", "AD IGST", "K+O+T+W+CP",
         "Sale Summary", "Sale Summary", "positional columns"),
        ("Step1_Liability", "AE CGST", "L+P+X+CQ",
         "Sale Summary", "Sale Summary", "positional columns"),
        ("Step1_Liability", "AF SGST", "M+Q+Y+CR",
         "Sale Summary", "Sale Summary", "positional columns"),
        ("Step1_Liability", "Exempt ratio", "AB / AC", "derived", "-",
         "drives the Rule 42 reversal"),
        ("Step2_ITC", "a. ITC FY 25-26", "sum of tax by state",
         "Books FY 2025-26", "25-26",
         "MONTH OF ITC TAKEN = period; remark accepted; non-import"),
        ("Step2_ITC", "b. ITC FY 26-27", "sum of tax by state",
         "Books FY 2026-27", "26-27", "same filters as (a)"),
        ("Step2_ITC", "a and b (override)", "replaces (a) and (b) where present",
         "ITC as per books - summary", "Sheet1",
         "State / IGST / CGST / SGST / Cess / Remarks; Remarks names the FY"),
        ("Step2_ITC", "c. Import", "sum of tax by state",
         "Books FY 25-26 / 26-27", "25-26 / 26-27", "Category = Import"),
        ("Step2_ITC", "d. Stock received", "sum of tax by state",
         "Stock Received summary", "Stock Recd_Summary", "-"),
        ("Step2_ITC", "e. Cross charge", "sum of tax by state",
         "Liability Summary", "Cross charge", "-"),
        ("Step2_ITC", "f. ISD", "sum of tax by state",
         "ISD distribution", "<month> / Turnover", "positional columns"),
        ("Step2_ITC", "g. RCM", "sum of tax by state",
         "Final Expenses (RCM)", "Final ... / RCM ONLY", "positional columns"),
        ("Step2_ITC", "h. Rule 42 reversal",
         "(expense + g.RCM + f.ISD + e.Cross) x exempt ratio",
         "Rule 42 expenses FY 25-26 + FY 26-27", "Pivot",
         "Row Labels = state; Sum of IGST/CGST/SGST BOOKS"),
        ("Step2_ITC", "Opening credit ledger", "as reported",
         "PMT Balance (or Liability Summary)", "Sheet1 / PMT Balance",
         "Credit block: IGST / CGST / SGST / cess"),
        ("Step2_ITC", "Available ITC", "a+b+c+d+e+f+g - h + opening",
         "derived", "-", "-"),
        ("Step3_Cash", "Credit utilised", "s.49(5), 49A, 49B and Rule 88A order",
         "derived", "-",
         "IGST first; residual IGST split CGST/SGST by the sidebar slider"),
        ("Step3_Cash", "Cash on supply", "liability - credit utilised",
         "derived", "-", "floored at zero per head"),
        ("Step3_Cash", "TOTAL CASH", "cash on supply + RCM liability",
         "derived", "-", "RCM is always cash - proviso to s.49(4)"),
        ("Step4_GSTR-3B", "3.1(a) Outward taxable", "Step 1 AA / AD / AE / AF / AG",
         "derived", "Step1_Liability", "-"),
        ("Step4_GSTR-3B", "3.1(c) Nil / exempt", "Step 1 AB",
         "derived", "Step1_Liability", "-"),
        ("Step4_GSTR-3B", "3.1(d) Inward RCM", "RCM liability",
         "Final Expenses (RCM)", "Final ... / RCM ONLY", "-"),
        ("Step4_GSTR-3B", "4A(1) Import of goods", "Step 2 (c) Import",
         "Books FY 25-26 / 26-27", "25-26 / 26-27",
         "Category = Import; IGST only"),
        ("Step4_GSTR-3B", "4A(2) Import of services", "no separate source yet",
         "-", "-", "reads zero until the import bucket is split"),
        ("Step4_GSTR-3B", "4A(3) Inward RCM", "Step 2 (g) RCM",
         "Final Expenses (RCM)", "Final ... / RCM ONLY", "-"),
        ("Step4_GSTR-3B", "4A(4) ISD", "Step 2 (f) ISD",
         "ISD distribution", "<month> / Turnover", "-"),
        ("Step4_GSTR-3B", "4A(5) All other ITC",
         "current month 2B + 4D(1) reclaim (2B part)",
         "GSTR-2B current FY + earlier FY", "GSTR-2B",
         "C MONTH; D STATE; AE = Considered; O:R tax"),
        ("Step4_GSTR-3B", "4B(1) Rules 38/42/43 & 17(5)",
         "2B ineligible (this month) + Step 2 (h) Rule 42",
         "GSTR-2B + Rule 42 expenses", "GSTR-2B / Pivot",
         "AG = 'Ineligible <month>'; AE = Considered"),
        ("Step4_GSTR-3B", "4B(2) Others", "sum of tax by state",
         "GSTR-2B (both FYs)", "GSTR-2B",
         "AI Other Reversal = filing month; AE = Considered"),
        ("Step4_GSTR-3B", "4C Net ITC available", "4A - 4B", "derived", "-",
         "= 4A(1)+4A(3)+4A(4)+4A(5) - 4B(1) - 4B(2)"),
        ("Step4_GSTR-3B", "4D(1) ITC reclaimed",
         "2B reclaim (this month) + Rule 37 reclaim",
         "GSTR-2B + Liability Summary", "GSTR-2B / ITC as per new format",
         "AH = 'Reclaim <month>' or 'Reclaim Ineligible <month>'"),
        ("Step4_GSTR-3B", "Reclaim ledger opening", "sum of tax by state",
         "GSTR-2B (both FYs)", "GSTR-2B",
         "C MONTH before the period; AE = Considered; "
         "AG blank / names the month / starts 'Split'"),
        ("Step4_GSTR-3B", "Reclaim ledger closing", "opening + 4B(2) - 4D(1)",
         "derived", "-",
         "reconcile against the Electronic Reclaim Ledger after filing"),
        ("Step4_GSTR-3B", "Current month 2B (Considered)", "sum of tax by state",
         "GSTR-2B current FY", "GSTR-2B",
         "C MONTH = filing period; AE = Considered"),
        ("Step4_GSTR-3B", "6.1 Cash payable", "Step 3 TOTAL CASH",
         "derived", "Step3_Cash", "-"),
        ("Exceptions", "Severity / Check / State / Detail",
         "findings raised while computing", "-", "-",
         "ERROR blocks filing; WARN needs a human call; INFO is audit"),
        ("Run_Log", "Load record",
         "file, sheet, header row and row count per source", "-", "-",
         "check this first if a figure looks wrong"),
    ]
    return pd.DataFrame(rows, columns=[
        "Output sheet", "Field", "Formula", "Source file", "Source sheet",
        "Columns / filter"])


def export_gstr3b_excel(results) -> bytes:
    output = io.BytesIO()
    steps = results.get("steps", {})
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        if "step1" in steps:
            df = step1_frame(steps["step1"]["grid"])
            df.to_excel(writer, sheet_name="Step1_Liability", index=False)
            _format_data_sheet(writer, "Step1_Liability", df)

        if "step2" in steps:
            df = step2_frame(steps["step2"]["itc"])
            df.to_excel(writer, sheet_name="Step2_ITC", index=False)
            _format_data_sheet(writer, "Step2_ITC", df)

        if "step3" in steps:
            df = step3_frame(steps["step3"]["cash"])
            df.to_excel(writer, sheet_name="Step3_Cash", index=False)
            _format_data_sheet(writer, "Step3_Cash", df)

        if "step4" in steps:
            df = gstr3b_frame(steps["step4"]["tables"])
            df.to_excel(writer, sheet_name="Step4_GSTR-3B", index=False)
            _format_data_sheet(writer, "Step4_GSTR-3B", df)

        exc = exceptions_frame(results.get("exceptions", []))
        exc.to_excel(writer, sheet_name="Exceptions", index=False)
        _format_data_sheet(writer, "Exceptions", exc)

        _fm = formula_map_frame()
        _fm.to_excel(writer, sheet_name="Formula_Map", index=False)
        _format_data_sheet(writer, "Formula_Map", _fm)

        # Run details last - provenance, not a working sheet, so it should not
        # sit in front of Step 1.
        _run = _build_all_errors_df(results)
        _run.to_excel(writer, sheet_name="Run_Log", index=False)
        _format_data_sheet(writer, "Run_Log", _run)

    output.seek(0)
    return output.getvalue()
