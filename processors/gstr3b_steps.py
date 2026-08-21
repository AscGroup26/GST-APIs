"""
GSTR-3B Steps 1–4.

Step 1  Output tax liability      — taxable/exempt tagging, state-wise pivot
Step 2  ITC availability          — 8 buckets + Rule 42 reversal + opening ECL
Step 3  Set-off and cash          — s.49(5), 49A, 49B and Rule 88A
Step 4  Portal table mapping      — 3.1, 4A, 4B, 4C, 6.1

Each function takes loaded dataframes and returns (result_dict, exceptions).
Nothing here reads files or contacts the GST portal.
"""
import re
from collections import defaultdict

import pandas as pd

from . import gstr3b_config as G
from .utils import VOUCHER_STATE_ALIASES, safe_numeric

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}


# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def _ex(sev, check, state, detail):
    return {"Severity": sev, "Check": check, "State": state, "Detail": detail}


def _zero():
    return {h: 0.0 for h in G.HEADS}


def norm_state(val):
    """Normalise a state name to the project's Title-Case convention."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = " ".join(str(val).split())
    if not s:
        return ""
    up = s.upper()
    if up in G.EXTRA_STATE_ALIASES:
        return G.EXTRA_STATE_ALIASES[up]
    if up in VOUCHER_STATE_ALIASES:
        return VOUCHER_STATE_ALIASES[up]
    up2 = up.replace(" AND ", " & ")
    if up2 in G.EXTRA_STATE_ALIASES:
        return G.EXTRA_STATE_ALIASES[up2]
    return s.title()


def norm_gst_state(val):
    """
    Like `norm_state`, but blanks anything that is not a real GST state.

    The source pivots carry 'Grand Total' / 'Total' rows. Without this guard
    they survive as a pseudo-state and get counted alongside the real ones,
    silently doubling the figures they summarise.
    """
    s = norm_state(val)
    return s if s in G.CODE_BY_STATE else ""


def _norm_header(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(val).replace("₹", "").strip().lower())


def _read_excel(src, sheet, header, nrows=None):
    """
    Read a sheet, falling back to calamine when openpyxl refuses the file.

    The GSTR-2B workbooks carry an autofilter openpyxl cannot parse - it raises
    "Value must be either numerical or a string containing a wildcard" before a
    single row is read. calamine ignores autofilters entirely and reads them
    fine, so it is the fallback rather than a hard dependency.
    """
    try:
        return pd.read_excel(src, sheet_name=sheet, header=header,
                             nrows=nrows, dtype=object)
    except Exception:
        try:
            src.seek(0)          # an uploaded file object is already consumed
        except (AttributeError, OSError):
            pass
        return pd.read_excel(src, sheet_name=sheet, header=header,
                             nrows=nrows, engine="calamine")


def read_table(src, sheet, key_tokens, max_scan=15):
    """
    Read `sheet` locating the header row by looking for `key_tokens`.

    The working files put headers on row 2, 3, 4 or 7 depending on the file,
    so the row is detected rather than hard-coded.
    Returns (dataframe, zero-based header row index).
    """
    raw = _read_excel(src, sheet, header=None, nrows=max_scan)
    want = {_norm_header(t) for t in key_tokens if t}
    best, best_hits = None, 0
    for i in range(len(raw)):
        hits = len(want & {_norm_header(v) for v in raw.iloc[i].tolist()})
        if hits > best_hits:
            best, best_hits = i, hits
    if best is None or best_hits == 0:
        raise ValueError(
            f"No header row found in sheet '{sheet}' — looked for {sorted(want)} "
            f"in the first {max_scan} rows."
        )
    df = _read_excel(src, sheet, header=best)
    df.columns = [str(c).strip() for c in df.columns]
    return df, best


def resolve(df, spec):
    """Map logical field names onto real dataframe columns (fuzzy)."""
    norm = {}
    for c in df.columns:
        norm.setdefault(_norm_header(c), c)
    out = {}
    for field, candidates in spec.items():
        hit = None
        for cand in candidates:
            nc = _norm_header(cand)
            if nc in norm:
                hit = norm[nc]
                break
        if hit is None:
            for cand in candidates:
                nc = _norm_header(cand)
                for k, real in norm.items():
                    if nc and (k.startswith(nc) or nc in k):
                        hit = real
                        break
                if hit:
                    break
        out[field] = hit
    return out


def _col(df, name):
    if name is None or name not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index)
    return safe_numeric(df[name])


def _pos(df, idx):
    if idx >= len(df.columns):
        return pd.Series([0.0] * len(df), index=df.index)
    return safe_numeric(df.iloc[:, idx])


def _text(df, name):
    if name is None or name not in df.columns:
        return pd.Series([""] * len(df), index=df.index)
    return df[name].map(
        lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    )


def parse_period(period):
    """'052026' -> (5, 2026)."""
    p = str(period or "").strip()
    if len(p) == 6 and p.isdigit():
        return int(p[:2]), int(p[2:])
    raise ValueError(f"Tax period must be MMYYYY, got {period!r}")


def month_matches(value, month, year):
    """True when value denotes the given month/year (datetime, 'May-26', ISO text)."""
    v_month, v_year = _extract_month_year(value)
    return v_month == month and v_year == year


def _extract_month_year(value):
    """Best-effort (month, year) from a datetime, 'May-26' text, or ISO text."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, None
    if isinstance(value, pd.Timestamp):
        return value.month, value.year
    try:
        import datetime as _dt
        if isinstance(value, (_dt.datetime, _dt.date)):
            return value.month, value.year
    except Exception:
        pass
    s = str(value).strip().lower()
    if not s:
        return None, None
    m = re.search(r"(\d{4})-(\d{2})-\d{2}", s)
    if m:
        return int(m.group(2)), int(m.group(1))
    m = re.search(r"([a-z]{3})[a-z]*[\s\-/]*(\d{2,4})", s)
    if m:
        mon = _MONTHS.get(m.group(1))
        if mon is None:
            return None, None
        yr = int(m.group(2))
        yr += 2000 if yr < 100 else 0
        return mon, yr
    return None, None


_FY_TAG_RE = re.compile(r"(\d{2,4})\s*[-/]\s*(\d{2,4})")


def _fy_start_from_tag(text):
    """'F.Y. 2025-26' / '25-26' -> 2025 (the FY's starting calendar year)."""
    m = _FY_TAG_RE.search(text)
    if not m:
        return None
    y = int(m.group(1))
    return y if y > 100 else 2000 + y


def classify_itc_summary_bucket(remark):
    """
    Which Step 2 bucket ('a_2526' or 'b_2627') an 'ITC as per books' summary
    row belongs to. The Remarks column carries either a plain date/period
    (an invoice-taken date, e.g. '2026-05-01') or an explicit FY tag
    ('Related to F.Y. 2025-26'). Date parsing is tried FIRST — a raw date
    like '2026-05-01' would otherwise be misread by the FY-tag regex as if
    '2026' were an FY start year, which is wrong for Jan/Feb/Mar dates
    (e.g. Feb-2026 belongs to FY 2025-26, not FY 2026-27, under the
    April-March Indian fiscal year). Returns None when the remark names
    neither of the two years this app tracks.
    """
    r = " ".join(str(remark or "").split())
    if not r:
        return None
    mon, yr = _extract_month_year(r)
    if mon is not None:
        fy_start = yr if mon >= 4 else yr - 1
    else:
        fy_start = _fy_start_from_tag(r)
    if fy_start == G.FY_START_A:
        return "a_2526"
    if fy_start == G.FY_START_B:
        return "b_2627"
    return None


def classify_remark(remark):
    """
    Decide whether a books line counts towards availed ITC.

    The Remarks column is hand-typed and carries misspellings
    ('PROBABLE MATHCED', 'TAKRN AS PER BOOKS'). Matching on the leading word
    keeps those lines in; the second value tags anything non-canonical so it
    is reported rather than silently absorbed.
    """
    r = " ".join(str(remark or "").split()).upper()
    if not r:
        return False, None
    for canon in ("PERFECT", "PROBABLE"):
        if r.startswith(canon):
            if r[len(canon):].strip().startswith("MATCHED"):
                return True, None
            return True, "typo:" + r
    if "AS PER BOOKS" in r:
        return True, "typo:" + r
    if r.startswith("REVERSE IN 3B"):
        return True, "reversal:" + r
    return False, None


# ─────────────────────────────────────────────────────────────────
# STEP 1 — OUTPUT TAX LIABILITY
# ─────────────────────────────────────────────────────────────────
def process_step1(sale_df):
    """Tag exempt/taxable, pivot state-wise, derive the net figures."""
    cols = resolve(sale_df, G.SALE_SPEC)
    missing = [k for k in ("state", "slab", "taxable", "type") if not cols[k]]
    if missing:
        raise ValueError(f"Sale sheet missing required columns: {missing}")

    exceptions = []
    st = sale_df[cols["state"]].map(norm_gst_state)
    typ = _text(sale_df, cols["type"])
    slab = _col(sale_df, cols["slab"])
    taxable = _col(sale_df, cols["taxable"])
    igst, cgst = _col(sale_df, cols["igst"]), _col(sale_df, cols["cgst"])
    sgst, cess = _col(sale_df, cols["sgst"]), _col(sale_df, cols["cess"])

    grid = defaultdict(lambda: {c: 0.0 for c in G.STEP1_COLS})
    skipped_blank = 0

    for i in range(len(sale_df)):
        s = st.iat[i]
        if not s:
            skipped_blank += 1          # embedded total rows carry a blank state
            continue
        t = typ.iat[i]
        if t not in G.KNOWN_TYPES:
            exceptions.append(_ex("ERROR", "Unknown TYPE", s,
                                  "row %d: TYPE=%r not in %s — line EXCLUDED"
                                  % (i + 2, t, sorted(G.KNOWN_TYPES))))
            continue

        exempt = abs(slab.iat[i]) < 1e-9
        b = G.STEP1_BUCKETS[(t, exempt)]
        g = grid[s]
        g[b["value"]] += taxable.iat[i]

        if b["tax"]:
            ig, cg, sg = b["tax"]
            if ig:
                g[ig] += igst.iat[i]
            if cg:
                g[cg] += cgst.iat[i]
            if sg:
                g[sg] += sgst.iat[i]
        elif igst.iat[i] or cgst.iat[i] or sgst.iat[i]:
            exceptions.append(_ex("ERROR", "Tax on exempt line", s,
                                  "row %d: slab 0 (%s) but carries tax IGST=%.2f "
                                  "CGST=%.2f SGST=%.2f"
                                  % (i + 2, t, igst.iat[i], cgst.iat[i], sgst.iat[i])))
        g["AG"] += cess.iat[i]

    # derived columns — verified against the live '3B Liability' sheet
    for s, g in grid.items():
        g["AA"] = g["J"] + g["N"] + g["S"] + g["V"] + g["CO"]
        g["AB"] = g["R"] + g["U"] + g["Z"]
        g["AC"] = g["AA"] + g["AB"]
        g["AD"] = g["K"] + g["O"] + g["T"] + g["W"] + g["CP"]
        g["AE"] = g["L"] + g["P"] + g["X"] + g["CQ"]
        g["AF"] = g["M"] + g["Q"] + g["Y"] + g["CR"]

    for s, g in grid.items():
        if g["V"] > G.TOLERANCE:
            exceptions.append(_ex("ERROR", "Sale return positive", s,
                                  "Sale return %.2f is positive; returns must be "
                                  "carried negative" % g["V"]))
        if abs(g["AE"] - g["AF"]) > G.TOLERANCE:
            exceptions.append(_ex("ERROR", "CGST != SGST", s,
                                  "CGST %.2f vs SGST %.2f — intra-state tax must "
                                  "split equally" % (g["AE"], g["AF"])))

    if cols["g1remarks"]:
        flags = sale_df[sale_df[cols["g1remarks"]].notna()]
        if len(flags):
            tot = _col(flags, cols["taxable"]).sum()
            exceptions.append(_ex(
                "INFO", "GSTR-1 reclassification flags",
                ", ".join(sorted({norm_state(v) for v in flags[cols["state"]]})),
                "%d line(s), taxable %.2f, flagged for B2C reclassification. "
                "No effect on 3.1(a) totals; GSTR-1 must agree." % (len(flags), tot)))

    grid = dict(grid)
    summary = {
        "source_rows": len(sale_df),
        "total_rows_skipped": skipped_blank,
        "states": len(grid),
        "net_taxable_sales": sum(g["AA"] for g in grid.values()),
        "net_exempt_sales": sum(g["AB"] for g in grid.values()),
        "output_igst": sum(g["AD"] for g in grid.values()),
        "output_cgst": sum(g["AE"] for g in grid.values()),
        "output_sgst": sum(g["AF"] for g in grid.values()),
    }
    return {"grid": grid, "summary": summary, "exceptions": exceptions}


# ─────────────────────────────────────────────────────────────────
# STEP 2 — ITC AVAILABILITY
# ─────────────────────────────────────────────────────────────────
def _books_itc(df, label, month, year):
    """Split a books sheet by state into (non-import, import, expense-only)."""
    cols = resolve(df, G.BOOKS_SPEC)
    exc = []
    if not cols["state"]:
        return {}, {}, {}, [_ex("ERROR", "Books columns", "—",
                                "%s: no State column found" % label)]

    st = df[cols["state"]].map(norm_gst_state)
    cat = _text(df, cols["category"]).str.upper()
    rem = _text(df, cols["remarks"])
    igst, cgst = _col(df, cols["igst"]), _col(df, cols["cgst"])
    sgst, cess = _col(df, cols["sgst"]), _col(df, cols["cess"])

    mcol = cols.get("month_taken")
    if not mcol:
        exc.append(_ex("WARN", "Month of ITC taken", "—",
                       "%s: no 'MONTH OF ITC TAKEN' column — every matched line is "
                       "counted, which overstates ITC." % label))
        in_month = pd.Series([True] * len(df), index=df.index)
    else:
        in_month = df[mcol].map(lambda v: month_matches(v, month, year))

    other, imports, expense = defaultdict(_zero), defaultdict(_zero), defaultdict(_zero)
    odd, excluded = defaultdict(float), defaultdict(float)
    kept = dropped = 0
    month_mismatch = 0          # valid state, but MONTH OF ITC TAKEN != selected period

    for i in range(len(df)):
        s = st.iat[i]
        if not s:
            continue
        if not in_month.iat[i]:
            month_mismatch += 1
            continue
        ok, tag = classify_remark(rem.iat[i])
        if not ok:
            dropped += 1
            excluded[rem.iat[i] or "(blank)"] += igst.iat[i]
            continue
        if tag:
            odd[tag] += igst.iat[i]
        kept += 1
        vals = {"igst": igst.iat[i], "cgst": cgst.iat[i],
                "sgst": sgst.iat[i], "cess": cess.iat[i]}
        tgt = imports if G.CAT_IMPORT in cat.iat[i] else other
        for h in G.HEADS:
            tgt[s][h] += vals[h]
        if G.CAT_EXPENSE in cat.iat[i]:
            for h in G.HEADS:
                expense[s][h] += vals[h]

    exc.append(_ex("INFO", "Books filter", "—",
                   "%s: %d line(s) taken for the selected period; %d excluded for an "
                   "unrecognised remark; %d excluded because their 'MONTH OF ITC "
                   "TAKEN' is a different period." % (label, kept, dropped, month_mismatch)))

    # If every valid-state row was thrown out on the month filter, the selected
    # period almost certainly does not match this file — surface the period
    # the data actually carries so the mistake is a one-look fix, not a hunt.
    if kept == 0 and month_mismatch > 0 and mcol:
        suggestion = ""
        try:
            sample = df.loc[st != "", mcol].dropna()
            found = sorted({str(v) for v in sample.head(500)})[:5]
            if found:
                suggestion = " Periods seen in this file include: %s." % ", ".join(found)
        except Exception:
            pass
        exc.append(_ex(
            "ERROR", "Books period mismatch", "—",
            "%s: 0 line(s) matched the selected period (month=%d, year=%d) out of "
            "%d line(s) with a valid state — every one was excluded by the 'MONTH "
            "OF ITC TAKEN' filter. This bucket's ITC is showing as ZERO because of "
            "this, not because the file is empty.%s Check the Month/Year selector "
            "in the sidebar matches the return period you are filing."
            % (label, month, year, month_mismatch, suggestion)))
    for tagged, amt in sorted(odd.items(), key=lambda kv: -abs(kv[1])):
        kind, _, spelling = tagged.partition(":")
        if kind == "reversal":
            exc.append(_ex("INFO", "Books reversal adjustment", "—",
                           "%s: remark %r carries IGST %.2f and is INCLUDED as a "
                           "negative adjustment." % (label, spelling, amt)))
        else:
            exc.append(_ex("WARN", "Misspelled books remark", "—",
                           "%s: remark %r carries IGST %.2f. Read as a match remark "
                           "and INCLUDED. Correct the spelling at source."
                           % (label, spelling, amt)))
    for spelling, amt in sorted(excluded.items(), key=lambda kv: -abs(kv[1]))[:8]:
        if abs(amt) > 1:
            exc.append(_ex("WARN", "Books line excluded", "—",
                           "%s: remark %r carries IGST %.2f and was EXCLUDED."
                           % (label, spelling, amt)))
    return dict(other), dict(imports), dict(expense), exc


def process_itc_summary(df):
    """
    Parse the 'ITC as per books' state-wise summary — a pre-aggregated
    alternative to the raw Books FY 25-26 / FY 26-27 workbooks. One row per
    state per FY (Remarks names the FY or carries a taken-date), covering
    Expense + Purchase ITC only — Import ITC is tracked separately (bucket c)
    and is NOT in this file.

    Returns (a_2526_by_state, b_2627_by_state, exceptions). These override
    the corresponding buckets computed from the raw Books files in
    process_step2 when this summary is supplied.
    """
    cols = resolve(df, G.ITC_SUMMARY_SPEC)
    exceptions = []
    if not cols["state"]:
        return {}, {}, [_ex("ERROR", "ITC summary columns", "—",
                            "No State column found in the ITC-as-per-books summary.")]

    igst, cgst = _col(df, cols["igst"]), _col(df, cols["cgst"])
    sgst, cess = _col(df, cols["sgst"]), _col(df, cols["cess"])
    rem = _text(df, cols["remarks"])

    a_summary, b_summary = defaultdict(_zero), defaultdict(_zero)
    seen_a, seen_b = set(), set()
    declared_total = None

    for i in range(len(df)):
        raw_state = df[cols["state"]].iat[i]
        s = norm_state(raw_state)
        if not s:
            continue
        if s.upper() in ("TOTAL", "GRAND TOTAL"):
            declared_total = {"igst": igst.iat[i], "cgst": cgst.iat[i],
                              "sgst": sgst.iat[i], "cess": cess.iat[i]}
            continue

        bucket = classify_itc_summary_bucket(rem.iat[i])
        vals = {"igst": igst.iat[i], "cgst": cgst.iat[i],
                "sgst": sgst.iat[i], "cess": cess.iat[i]}
        if bucket == "a_2526":
            for h in G.HEADS:
                a_summary[s][h] += vals[h]
            seen_a.add(s)
        elif bucket == "b_2627":
            for h in G.HEADS:
                b_summary[s][h] += vals[h]
            seen_b.add(s)
        elif any(vals.values()):
            exceptions.append(_ex(
                "WARN", "Unrecognised ITC-summary remark", s,
                "row %d: remark %r does not name FY %d-%s or FY %d-%s, and does "
                "not parse as a date in either — IGST %.2f / CGST %.2f / SGST "
                "%.2f EXCLUDED from Step 2."
                % (i + 4, rem.iat[i], G.FY_START_A, str(G.FY_START_A + 1)[-2:],
                   G.FY_START_B, str(G.FY_START_B + 1)[-2:],
                   vals["igst"], vals["cgst"], vals["sgst"])))

    a_summary, b_summary = dict(a_summary), dict(b_summary)

    if declared_total is not None:
        computed = {h: sum(v[h] for v in a_summary.values())
                    + sum(v[h] for v in b_summary.values()) for h in G.HEADS}
        for h in ("igst", "cgst", "sgst"):
            if abs(computed[h] - declared_total[h]) > G.TOLERANCE:
                exceptions.append(_ex(
                    "WARN", "ITC-summary total mismatch", "—",
                    "%s total in the file is %.2f but the state-wise rows sum to "
                    "%.2f (diff %.2f) — some row may carry an unrecognised remark."
                    % (G.HEAD_LABEL[h], declared_total[h], computed[h],
                       computed[h] - declared_total[h])))

    exceptions.append(_ex(
        "INFO", "ITC-as-per-books summary used", "—",
        "%d state(s) read for FY %d-%s, %d state(s) for FY %d-%s from the "
        "uploaded summary — these OVERRIDE any 'a. ITC FY' / 'b. ITC FY' figures "
        "computed from the raw Books workbooks. Import ITC (bucket c) and the "
        "Rule 42 common-credit base still need the raw Books files if you want "
        "those populated."
        % (len(seen_a), G.FY_START_A, str(G.FY_START_A + 1)[-2:],
           len(seen_b), G.FY_START_B, str(G.FY_START_B + 1)[-2:])))

    return a_summary, b_summary, exceptions


def _rev42_expense(frames, exceptions):
    """
    Per-state expense ITC for the Rule 42 common-credit pool.

    Returns None when no usable pivot was supplied, so the caller can fall back
    to the books-derived subtotal.
    """
    if not frames:
        return None
    out = defaultdict(_zero)
    seen = False
    for df in frames:
        if df is None or not len(df):
            continue
        cols = resolve(df, G.REV42_EXPENSE_SPEC)
        if not cols["state"]:
            exceptions.append(_ex("WARN", "Rule 42 expense pivot", "—",
                                  "No state column found; this sheet was ignored."))
            continue
        st = df[cols["state"]].map(norm_gst_state)
        vals = {h: _col(df, cols[h]) for h in G.HEADS}
        for i in range(len(df)):
            if st.iat[i]:
                seen = True
                for h in G.HEADS:
                    out[st.iat[i]][h] += vals[h].iat[i]
    return dict(out) if seen else None


def process_step2(step1_grid, books, stock_df, cross_df, isd_df,
                  expense_df, opening_df, month, year, itc_summary=None,
                  rev42_expense=None):
    """
    Build the eight ITC buckets, the Rule 42 reversal and the opening ledger.

      a FY 25-26   b FY 26-27   c Import   d Stock received
      e Cross charge   f ISD   g RCM   h Rule 42/43 reversal

      current   = a+b+c+d+e+f+g − h
      available = current + opening electronic credit ledger
    """
    exceptions = []
    keys = ("a_2526", "b_2627", "c_import", "d_stock", "e_cross", "f_isd",
            "g_rcm", "h_rev42", "expense_only", "rule42_base",
            "current", "opening", "available")
    out = defaultdict(lambda: {k: _zero() for k in keys})

    # a, b, c — books
    for fy, df in books.items():
        if df is None or not len(df):
            continue
        other, imp, expense, exc = _books_itc(df, fy, month, year)
        exceptions += exc
        bucket = "a_2526" if "25" in fy else "b_2627"
        for s, v in other.items():
            for h in G.HEADS:
                out[s][bucket][h] += v[h]
        for s, v in imp.items():
            for h in G.HEADS:
                out[s]["c_import"][h] += v[h]
        for s, v in expense.items():
            for h in G.HEADS:
                out[s]["expense_only"][h] += v[h]

    # d — stock received (IGST only)
    if stock_df is not None and len(stock_df):
        c = resolve(stock_df, G.STOCK_SPEC)
        st = stock_df[c["state"] or stock_df.columns[0]].map(norm_gst_state)
        vals = _col(stock_df, c.get("total"))
        for i in range(len(stock_df)):
            if st.iat[i]:
                out[st.iat[i]]["d_stock"]["igst"] += vals.iat[i]

    # e — cross charge received (second 'Row Labels' block of the pivot)
    if cross_df is not None and len(cross_df):
        labels = [c for c in cross_df.columns if str(c).lower().startswith("row label")]
        if len(labels) >= 2:
            gcol = labels[1]
            at = list(cross_df.columns).index(gcol)
            icol = cross_df.columns[at + 2] if at + 2 < len(cross_df.columns) else None
            st = cross_df[gcol].map(norm_gst_state)
            vals = _col(cross_df, icol)
            for i in range(len(cross_df)):
                if st.iat[i]:
                    out[st.iat[i]]["e_cross"]["igst"] += vals.iat[i]
        else:
            exceptions.append(_ex("WARN", "Cross charge layout", "—",
                                  "Could not locate the 'received' pivot block; "
                                  "cross-charge ITC treated as nil."))

    # f — ISD (distributed columns O/P/Q, addressed by position)
    if isd_df is not None and len(isd_df):
        st = isd_df.iloc[:, G.ISD_POS["state"]].map(norm_gst_state)
        ig = _pos(isd_df, G.ISD_POS["igst"])
        cg = _pos(isd_df, G.ISD_POS["cgst"])
        sg = _pos(isd_df, G.ISD_POS["sgst"])
        for i in range(len(isd_df)):
            if st.iat[i]:
                out[st.iat[i]]["f_isd"]["igst"] += ig.iat[i]
                out[st.iat[i]]["f_isd"]["cgst"] += cg.iat[i]
                out[st.iat[i]]["f_isd"]["sgst"] += sg.iat[i]

    # g — RCM credit, plus the full RCM liability (which includes payment-only)
    rcm_liab = defaultdict(lambda: dict(_zero(), taxable=0.0))
    if expense_df is not None and len(expense_df):
        P = G.EXPENSE_POS
        st = expense_df.iloc[:, P["state"]].map(norm_gst_state)
        v = {k: _pos(expense_df, P[k]) for k in P if k != "state"}
        for i in range(len(expense_df)):
            s = st.iat[i]
            if not s:
                continue
            out[s]["g_rcm"]["igst"] += v["gta_igst"].iat[i] + v["oth_igst"].iat[i]
            out[s]["g_rcm"]["cgst"] += v["gta_cgst"].iat[i] + v["oth_cgst"].iat[i]
            out[s]["g_rcm"]["sgst"] += v["gta_sgst"].iat[i] + v["oth_sgst"].iat[i]
            rcm_liab[s]["igst"] += (v["gta_igst"].iat[i] + v["oth_igst"].iat[i]
                                    + v["pay_igst"].iat[i])
            rcm_liab[s]["cgst"] += (v["gta_cgst"].iat[i] + v["oth_cgst"].iat[i]
                                    + v["pay_cgst"].iat[i])
            rcm_liab[s]["sgst"] += (v["gta_sgst"].iat[i] + v["oth_sgst"].iat[i]
                                    + v["pay_sgst"].iat[i])
            rcm_liab[s]["taxable"] += (v["gta_net"].iat[i] + v["oth_net"].iat[i]
                                       + v["pay_net"].iat[i])

    # h — Rule 42 reversal. Common credit = expenses + RCM + ISD + cross charge.
    # Trading purchases are excluded: attributable to taxable outward supply.
    # The expense leg comes from the Credit Reversal working's own pivot when
    # supplied - see the note on REV42_EXPENSE_SPEC for why the books subtotal
    # is not equivalent.
    rev_exp = _rev42_expense(rev42_expense, exceptions)
    if rev_exp is None:
        exceptions.append(_ex(
            "WARN", "Rule 42 expense base", "—",
            "No Rule 42 expense pivot supplied; falling back to the books "
            "Category='Expense' subtotal. That subtotal excludes Rule 37 "
            "reclaim lines, so the CGST/SGST reversal may be understated. "
            "Upload 'Expenses <FY>.xlsx' from 'Reversal on Exempted Sales'."))

    for s, rec in out.items():
        g1 = step1_grid.get(s)
        total = g1["AC"] if g1 else 0.0
        exempt = g1["AB"] if g1 else 0.0
        ratio = (exempt / total) if abs(total) > 1e-9 else 0.0
        rec["exempt_ratio"] = ratio
        base_expense = rec["expense_only"] if rev_exp is None else \
            rev_exp.get(s, _zero())
        rec["rule42_expense"] = dict(base_expense)
        for h in G.HEADS:
            common = (base_expense[h] + rec["g_rcm"][h]
                      + rec["f_isd"][h] + rec["e_cross"][h])
            rec["rule42_base"][h] = common
            rec["h_rev42"][h] = common * ratio

    # optional 'ITC as per books' summary — overrides a_2526/b_2627 for any
    # state it names, in place of the raw Books-file computation above.
    # Import (c_import) and expense_only (the Rule 42 base) are left as
    # computed from the raw Books files, since this summary does not carry
    # that split.
    if itc_summary is not None:
        a_summary, b_summary = itc_summary
        for s, v in a_summary.items():
            for h in G.HEADS:
                out[s]["a_2526"][h] = v[h]
        for s, v in b_summary.items():
            for h in G.HEADS:
                out[s]["b_2627"][h] = v[h]

    # opening electronic credit ledger
    if opening_df is not None and len(opening_df):
        c = resolve(opening_df, G.PMT_SPEC)
        st = opening_df[c["state"] or opening_df.columns[0]].map(norm_gst_state)
        vals = {h: _col(opening_df, c[h]) for h in G.HEADS}
        for i in range(len(opening_df)):
            if st.iat[i]:
                for h in G.HEADS:
                    out[st.iat[i]]["opening"][h] += vals[h].iat[i]

    for s, rec in out.items():
        for h in G.HEADS:
            rec["current"][h] = (rec["a_2526"][h] + rec["b_2627"][h]
                                 + rec["c_import"][h] + rec["d_stock"][h]
                                 + rec["e_cross"][h] + rec["f_isd"][h]
                                 + rec["g_rcm"][h] - rec["h_rev42"][h])
            rec["available"][h] = rec["current"][h] + rec["opening"][h]
        rec["rcm_liability"] = dict(rcm_liab.get(s, dict(_zero(), taxable=0.0)))
        if rec["available"]["igst"] < -G.TOLERANCE:
            exceptions.append(_ex("ERROR", "Negative ITC available", s,
                                  "IGST credit available is %.2f"
                                  % rec["available"]["igst"]))

    out = dict(out)
    summary = {
        "states": len(out),
        "itc_current_igst": sum(v["current"]["igst"] for v in out.values()),
        "itc_current_cgst": sum(v["current"]["cgst"] for v in out.values()),
        "itc_current_sgst": sum(v["current"]["sgst"] for v in out.values()),
        "itc_available_total": sum(
            sum(v["available"][h] for h in ("igst", "cgst", "sgst")) for v in out.values()),
        "rule42_reversal_total": sum(
            sum(v["h_rev42"][h] for h in ("igst", "cgst", "sgst")) for v in out.values()),
    }
    return {"itc": out, "summary": summary, "exceptions": exceptions}


# ─────────────────────────────────────────────────────────────────
# STEP 3 — SET-OFF AND CASH LIABILITY
# ─────────────────────────────────────────────────────────────────
def process_step3(step1_grid, itc, igst_split=G.DEFAULT_IGST_SPLIT):
    """
    Utilise credit against liability per s.49(5), 49A, 49B and Rule 88A.

      1 IGST credit against IGST liability first
      2 residual IGST credit against CGST and SGST, split by `igst_split`
      3 CGST credit against CGST, then residual against IGST
      4 SGST credit against SGST, then residual against IGST
      5 CGST credit never touches SGST liability, or vice versa (s.49(5))

    RCM is added afterwards and is always cash — credit cannot discharge it
    (proviso to s.49(4)).
    """
    exceptions = []
    out = {}

    for s in sorted(set(step1_grid) | set(itc)):
        g1 = step1_grid.get(s, {c: 0.0 for c in G.STEP1_COLS})
        rec = itc.get(s)
        avail = dict(rec["available"]) if rec else _zero()
        rcm = dict(rec["rcm_liability"]) if rec else dict(_zero(), taxable=0.0)

        liab = {"igst": g1["AD"], "cgst": g1["AE"], "sgst": g1["AF"], "cess": g1["AG"]}
        rem_l, rem_c = dict(liab), dict(avail)
        used = defaultdict(float)

        u = max(min(rem_c["igst"], rem_l["igst"]), 0.0)
        used["igst_in_igst"] = u
        rem_c["igst"] -= u
        rem_l["igst"] -= u

        if rem_c["igst"] > 0:
            share = rem_c["igst"]
            to_c = min(share * igst_split, rem_l["cgst"])
            to_s = min(share * (1 - igst_split), rem_l["sgst"])
            left = share - to_c - to_s
            if left > 0:
                extra = max(min(left, rem_l["cgst"] - to_c), 0.0)
                to_c += extra
                left -= extra
            if left > 0:
                to_s += max(min(left, rem_l["sgst"] - to_s), 0.0)
            used["igst_in_cgst"], used["igst_in_sgst"] = to_c, to_s
            rem_c["igst"] -= to_c + to_s
            rem_l["cgst"] -= to_c
            rem_l["sgst"] -= to_s

        for head, other in (("cgst", "igst"), ("sgst", "igst")):
            u = max(min(rem_c[head], rem_l[head]), 0.0)
            used["%s_in_%s" % (head, head)] = u
            rem_c[head] -= u
            rem_l[head] -= u
            u = max(min(rem_c[head], rem_l[other]), 0.0)
            used["%s_in_%s" % (head, other)] = u
            rem_c[head] -= u
            rem_l[other] -= u

        u = max(min(rem_c["cess"], rem_l["cess"]), 0.0)
        used["cess_in_cess"] = u
        rem_c["cess"] -= u
        rem_l["cess"] -= u

        cash_supply = {h: max(rem_l[h], 0.0) for h in G.HEADS}
        out[s] = {
            "liability": liab,
            "available": avail,
            "utilised": dict(used),
            "closing_credit": {h: max(rem_c[h], 0.0) for h in G.HEADS},
            "cash_on_supply": cash_supply,
            "rcm_liability": rcm,
            "cash_total": {h: cash_supply[h] + rcm.get(h, 0.0) for h in G.HEADS},
        }

    summary = {
        "states": len(out),
        "cash_on_supply_total": sum(
            sum(v["cash_on_supply"][h] for h in ("igst", "cgst", "sgst")) for v in out.values()),
        "rcm_cash_total": sum(
            sum(v["rcm_liability"].get(h, 0.0) for h in ("igst", "cgst", "sgst"))
            for v in out.values()),
        "cash_payable_total": sum(
            sum(v["cash_total"][h] for h in ("igst", "cgst", "sgst")) for v in out.values()),
        "igst_split_to_cgst": igst_split,
    }
    return {"cash": out, "summary": summary, "exceptions": exceptions}


# ─────────────────────────────────────────────────────────────────
# STEP 4 — PORTAL TABLE MAPPING
# ─────────────────────────────────────────────────────────────────
_MON_ABBR = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
             "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
_MON_FULL = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
             "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]


def _reclaim_tokens(period):
    """
    The 'Reclaimed' labels that belong to this return.

    The suffix names the month the credit is being taken back in, not the 2B
    month it came from - so filing May 2026 counts 'Reclaim May-26' and
    'Reclaim Ineligible May-26' and nothing else. The sheet spells some months
    out ('June-25', 'April-26') and abbreviates others, so both are accepted.
    """
    s = "".join(ch for ch in str(period or "") if ch.isdigit())
    if len(s) != 6:
        return set()
    m, y = int(s[:2]), int(s[2:])
    if not 1 <= m <= 12:
        return set()
    yy = str(y)[-2:]
    return {f"{_MON_ABBR[m - 1]}-{yy}", f"{_MON_FULL[m - 1]}-{yy}"}


def _g2b_month_key(period):
    """
    'MMYYYY' -> the MONTH value the GSTR-2B sheet uses.

    That column stores the period as month*10000 + year, so May 2026 is 52026
    and December 2025 is 122025 - written without a leading zero, which is why
    a plain string compare against '052026' never matches.
    """
    s = "".join(ch for ch in str(period or "") if ch.isdigit())
    if len(s) != 6:
        return None
    try:
        return int(s[:2]) * 10000 + int(s[2:])
    except ValueError:
        return None


def _opening_header_row(raw):
    """Row index whose next four cells are IGST / CGST / SGST / Cess."""
    for i in range(min(15, len(raw))):
        heads = [str(raw.iat[i, j]).strip().upper()
                 for j in range(1, min(5, raw.shape[1]))]
        if heads[:3] == ["IGST", "CGST", "SGST"]:
            return i
    return None


def reclaim_opening(src, sheet="ITC as per new format"):
    """
    Opening reclaim-ledger balance per state - state in column A, then
    IGST / CGST / SGST / Cess.

    Reads either a standalone opening-balance workbook or the liability
    workbook's 'ITC as per new format' sheet; the header sits on a different
    row in each, so it is located rather than assumed.
    """
    out = {}
    if src is None:
        return out
    raw = None
    for candidate in (sheet, "Sheet1"):
        try:
            raw = _read_excel(src, candidate, header=None)
            break
        except Exception:
            continue
    if raw is None or len(raw) < 2:
        return out
    hdr = _opening_header_row(raw)
    if hdr is None:
        return out
    for i in range(hdr + 1, len(raw)):
        s = norm_state(raw.iat[i, 0])
        if not s:
            continue
        # A state can appear twice - 'BIHAR' and 'Bihar'. Accumulate, because
        # assigning lets a later zero row wipe out an earlier balance.
        rec = out.setdefault(s, _zero())
        for j, h in enumerate(("igst", "cgst", "sgst", "cess"), start=1):
            v = pd.to_numeric(pd.Series([raw.iat[i, j]]), errors="coerce").fillna(0)
            rec[h] += float(v.iat[0])
    return out


def rule37_movements(liab_src):
    """
    Rule 37 reversal and reclaim per state, from the liability workbook.

    Credit reversed for non-payment within 180 days is a 4B(2) reversal, and
    reclaiming it once the supplier is paid is a 4D(1) reclaim - the same two
    buckets the GSTR-2B rows feed. Located by column name, not letter.
    """
    out = {}
    if liab_src is None:
        return out
    try:
        raw = _read_excel(liab_src, "ITC as per new format", header=None)
    except Exception:
        return out
    if len(raw) < 5:
        return out

    starts = {}
    for i in range(1, raw.shape[1] - 3):
        g = raw.iat[2, i]
        if not isinstance(g, str):
            continue
        label = g.strip().lower().replace("ule", "ul")
        if "ul 37" not in label:
            continue
        starts["reclaim" if label.startswith("reclaim") else "otherrev"] = i
    if not starts:
        return out

    for i in range(4, len(raw)):
        s = norm_state(raw.iat[i, 0])
        if not s:
            continue
        rec = out.setdefault(s, {"otherrev": _zero(), "reclaim": _zero()})
        for kind, start in starts.items():
            for j, h in enumerate(("igst", "cgst", "sgst", "cess")):
                v = pd.to_numeric(pd.Series([raw.iat[i, start + j]]),
                                  errors="coerce").fillna(0)
                rec[kind][h] += float(v.iat[0])
    return out


def process_step4(step1_grid, itc, cash, g2b_df, period=None,
                  reclaim_open=None, rule37=None):
    """
    Map the computed figures onto the GSTR-3B tables.

    `period` ('MMYYYY') decides which GSTR-2B rows count as the current month.
    Without it the current month is guessed from the most common MONTH value,
    which is wrong whenever the workbook holds a full year - the mode is then
    just the busiest month, not the one being filed.
    """
    exceptions = []
    cur, availed, inelig, reclaim = (defaultdict(_zero), defaultdict(_zero),
                                     defaultdict(_zero), defaultdict(_zero))
    # 4A(5) counts only the rows the books marked 'Considered'; 4B(2) still
    # needs the unfiltered current-month total, so the two are kept apart.
    cur_cons = defaultdict(_zero)
    otherrev_acc = defaultdict(_zero)
    # Opening reclaim ledger derived from 2B itself: earlier-month credit that
    # was still pending when this period began.
    open_acc = defaultdict(_zero)
    have_2b = False

    if g2b_df is not None and len(g2b_df):
        c = resolve(g2b_df, G.GSTR2B_SPEC)
        if c["state"] and c["month"]:
            have_2b = True
            st = g2b_df[c["state"]].map(norm_gst_state)
            month = _text(g2b_df, c["month"])
            addrem = _text(g2b_df, c["addremark"]).str.upper()
            taken = _text(g2b_df, c["taken3b"]).str.upper()
            recl = _text(g2b_df, c["reclaimed"]).str.upper()
            elig = _text(g2b_df, c["eligible"]).str.upper()
            vals = {h: _col(g2b_df, c[h]) for h in G.HEADS}
            # Compare numerically - the sheet writes 52026, the period '052026'.
            month_num = pd.to_numeric(month, errors="coerce")
            want = _g2b_month_key(period)
            # Reclaim rows are those whose label names THIS month; without a
            # period every 'Reclaim ...' row counts, which overstates 4A(5).
            _rtok = _reclaim_tokens(period)
            is_reclaim_now = recl.str.startswith("RECLAIM")
            if _rtok:
                is_reclaim_now &= recl.apply(
                    lambda t: any(tok in t for tok in _rtok))

            # 'TAKEN IN 3B' tags ineligible credit by the month it was taken -
            # 'Ineligible May-26', 'Ineligible Oct-25'. Only the tag naming
            # this return belongs in 4B(1); the rest were reversed already.
            is_inelig_now = taken.str.contains(G.INELIGIBLE, na=False)
            if _rtok:
                is_inelig_now &= taken.apply(
                    lambda t: any(tok in t for tok in _rtok))

            # 4B(2) needs the credit actually taken this month: 'TAKEN IN 3B'
            # holding either the plain month tag - stored as a date, so
            # '2026-05-01' rather than 'May-26' - or 'Ineligible May-26'.
            _pd = "".join(ch for ch in str(period or "") if ch.isdigit())
            if len(_pd) == 6:
                _dtok = f"{_pd[2:]}-{_pd[:2]}-"
                is_taken_now = taken.str.contains(_dtok, na=False, regex=False)
                if _rtok:
                    is_taken_now |= taken.apply(
                        lambda t: any(tok in t for tok in _rtok))
            else:
                is_taken_now = taken.astype(bool)

            # 4B(2) is flagged per row in 'Other Reversal' (col AI), which holds
            # a date - '2026-05-01' - so match the period prefix. This is read
            # straight off the sheet, not derived from current less availed.
            otherrev = _text(g2b_df, c["otherrev"]).str.upper()
            if len(_pd) == 6:
                is_otherrev_now = otherrev.str.contains(
                    f"{_pd[2:]}-{_pd[:2]}-", na=False, regex=False)
                if _rtok:
                    is_otherrev_now |= otherrev.apply(
                        lambda t: any(tok in t for tok in _rtok))
            else:
                is_otherrev_now = pd.Series(False, index=g2b_df.index)

            # Opening reclaim ledger: earlier-month credit still pending when
            # this period began. 'TAKEN IN 3B' is blank (never taken), names
            # this month (taken now, so it was pending at the open), or carries
            # a split remark (part-matched, remainder still pending).
            is_open_row = (taken.str.strip() == "") | is_taken_now \
                | taken.str.startswith("SPLIT")

            if want is None:
                mode = month[month != ""].mode()
                cur_month = mode.iat[0] if len(mode) else ""
                month_num = None
                exceptions.append(_ex(
                    "WARN", "GSTR-2B current month guessed", "-",
                    f"No tax period supplied, so the current month for table 4A(5) "
                    f"was taken as the most common MONTH value ({cur_month!r}). "
                    "Set the tax period to make this exact."))
            elif not (month_num == want).any():
                exceptions.append(_ex(
                    "ERROR", "GSTR-2B has no current-month rows", "-",
                    f"No GSTR-2B row carries MONTH {want} (period {period}). The "
                    "current-month half of table 4A(5) will be zero - check that "
                    "the 2B workbook for the filing month's financial year was "
                    "uploaded."))

            for i in range(len(g2b_df)):
                s = st.iat[i]
                if not s:
                    continue
                is_cur = (month_num.iat[i] == want if month_num is not None
                          else month.iat[i] == cur_month)
                considered = addrem.iat[i].startswith(G.CONSIDERED)
                for h in G.HEADS:
                    v = vals[h].iat[i]
                    if not v:
                        continue
                    # Flagged for other reversal regardless of which 2B month
                    # the invoice itself came from - but only where the books
                    # marked it 'Considered', same as every other 4A/4B figure.
                    if considered and is_otherrev_now.iat[i]:
                        otherrev_acc[s][h] += v
                    if is_cur:
                        cur[s][h] += v
                        if considered:
                            cur_cons[s][h] += v
                            if is_taken_now.iat[i]:
                                availed[s][h] += v
                            # 4B(1) keys on 'TAKEN IN 3B' alone. The separate
                            # Eligibility column is not part of the rule - it
                            # flags what 2B thinks, not what was reversed here.
                            if is_inelig_now.iat[i]:
                                inelig[s][h] += v
                    else:
                        if considered and is_open_row.iat[i]:
                            open_acc[s][h] += v
                        if is_reclaim_now.iat[i]:
                            reclaim[s][h] += v
                        if considered and is_inelig_now.iat[i]:
                            inelig[s][h] += v
        else:
            exceptions.append(_ex("WARN", "GSTR-2B columns", "—",
                                  "STATE / MONTH not found in the GSTR-2B sheet; "
                                  "tables 4A(5), 4B(1) and 4B(2) fall back to the "
                                  "books-derived figures."))

    out = {}
    for s in sorted(set(step1_grid) | set(itc) | set(cash)):
        g1 = step1_grid.get(s, {c: 0.0 for c in G.STEP1_COLS})
        rec = itc.get(s) or {}
        cs = cash.get(s, {})
        rcm = rec.get("rcm_liability", dict(_zero(), taxable=0.0))
        imp = rec.get("c_import", _zero())
        rcm_itc = rec.get("g_rcm", _zero())
        isd = rec.get("f_isd", _zero())
        rev42 = rec.get("h_rev42", _zero())

        # A state can have reclaim rows without any current-month invoice, so
        # keying only on `cur` would drop it back to the books formula and lose
        # its reclaim entirely.
        if have_2b and (s in cur or s in reclaim or s in otherrev_acc):
            a5 = {h: cur_cons[s][h] + reclaim[s][h] for h in G.HEADS}
            b1 = {h: inelig[s][h] + rev42[h] for h in G.HEADS}
            # 4B(2) = current-month ITC less what was actually availed, i.e. the
            # unmatched credit carried into next month's reconciliation. Both
            # sides are the 'Considered' figure, so they subtract like for like.
            # 4B(2) comes off the 'Other Reversal' flag on each row.
            b2 = dict(otherrev_acc[s])
        else:
            a5 = {h: (rec.get("a_2526", _zero())[h] + rec.get("b_2627", _zero())[h]
                      + rec.get("d_stock", _zero())[h]
                      + rec.get("e_cross", _zero())[h]) for h in G.HEADS}
            b1, b2 = dict(rev42), _zero()

        # Rule 37 movements sit alongside the GSTR-2B ones: its reversal is a
        # 4B(2) entry, its reclaim a 4D(1) entry.
        _r37 = (rule37 or {}).get(s) or {"otherrev": _zero(), "reclaim": _zero()}
        # Rule 37's reversal belongs to the reclaim ledger only - not to the
        # reported 4B(2), and not to 4C. Those two report the GSTR-2B columns;
        # Rule 37 is tracked separately on the working sheet.
        _r37_rev = _r37["otherrev"]
        _d1 = {h: reclaim[s][h] + _r37["reclaim"][h] for h in G.HEADS}
        # The opening reclaim ledger is derived from the 2B sheets: earlier
        # months, 'Considered', with 'TAKEN IN 3B' blank, naming this month, or
        # carrying a Split remark. A supplied opening-balance workbook is only
        # the fallback, for when no GSTR-2B is loaded.
        _supplied = (reclaim_open or {}).get(s)
        _ro = (dict(open_acc[s]) if have_2b and s in open_acc
               else (dict(_supplied) if _supplied is not None else _zero()))

        out[s] = {
            "t3_1a": {"taxable": g1["AA"], "igst": g1["AD"], "cgst": g1["AE"],
                      "sgst": g1["AF"], "cess": g1["AG"]},
            "t3_1c": {"taxable": g1["AB"]},
            "t3_1d": {"taxable": rcm.get("taxable", 0.0), "igst": rcm.get("igst", 0.0),
                      "cgst": rcm.get("cgst", 0.0), "sgst": rcm.get("sgst", 0.0), "cess": 0.0},
            "t4a1": imp, "t4a3": rcm_itc, "t4a4": isd, "t4a5": a5,
            "t4b1": b1, "t4b2": b2,
            "t4c": {h: a5[h] + imp[h] + rcm_itc[h] + isd[h] - b1[h] - b2[h]
                    for h in G.HEADS},
            # 4D(1) - credit reversed under 4B(2) in an earlier period and
            # reclaimed now. Same rows that feed the reclaim half of 4A(5):
            # matched this month, but the 2B they came from is an earlier one.
            "t4d1": _d1,
            # Reclaim ledger, for the record rather than the return itself.
            # Reconcile the closing figure with the Electronic Reclaim ledger
            # on the portal after filing.
            "reclaim_open": _ro,
            # Closing = opening + 4B(2) - 4D(1), taken from the columns on the
            # Step 4 sheet itself so the ledger can be checked across the row.
            # 4B(2) is the reclaimable reversal; 4D(1) is what was reclaimed
            # back out this period.
            "reclaim_close": {h: _ro[h] + b2[h] - _d1[h] for h in G.HEADS},
            # Current-month credit from the current FY's 2B - MONTH is the
            # filing period and the books marked it 'Considered'. Shown for
            # reference beside the ledger; it is the current-month half of
            # 4A(5), not a ledger movement.
            "current_2b": dict(cur_cons[s]),
            "t6_1": {"cash_total": cs.get("cash_total", _zero()),
                     "cash_on_supply": cs.get("cash_on_supply", _zero()),
                     "rcm_cash": cs.get("rcm_liability", _zero())},
        }

    summary = {
        "states": len(out),
        "gstr2b_used": have_2b,
        "t4a5_igst": sum(v["t4a5"]["igst"] for v in out.values()),
        "t4b1_igst": sum(v["t4b1"]["igst"] for v in out.values()),
        "t4b2_igst": sum(v["t4b2"]["igst"] for v in out.values()),
        "t4c_igst": sum(v["t4c"]["igst"] for v in out.values()),
    }
    return {"tables": out, "summary": summary, "exceptions": exceptions}


# ─────────────────────────────────────────────────────────────────
# CROSS-CUTTING COMPLIANCE CHECKS
# ─────────────────────────────────────────────────────────────────
def compliance_checks(step1_grid, itc, cash):
    """Statutory checks that span the steps."""
    out = []
    for s in sorted(step1_grid):
        g1 = step1_grid[s]
        cs = cash.get(s, {})

        if g1["AA"] > G.RULE_86B_THRESHOLD:
            liab = sum(cs.get("liability", {}).get(h, 0.0) for h in ("igst", "cgst", "sgst"))
            paid = sum(cs.get("cash_on_supply", {}).get(h, 0.0) for h in ("igst", "cgst", "sgst"))
            if liab > 0 and paid < 0.01 * liab:
                out.append(_ex("WARN", "Rule 86B", s,
                               "Taxable turnover %.0f exceeds Rs 50 lakh. Rule 86B needs at "
                               "least 1%% of output liability (%.2f) in cash; the set-off pays "
                               "%.2f. Check whether a proviso exemption applies."
                               % (g1["AA"], 0.01 * liab, paid)))

        rcm = cs.get("rcm_liability", {})
        tot = sum(rcm.get(h, 0.0) for h in ("igst", "cgst", "sgst"))
        if tot > G.TOLERANCE:
            out.append(_ex("INFO", "RCM payable in cash", s,
                           "RCM liability %.2f is included in the cash total — credit "
                           "cannot discharge it (proviso to s.49(4))." % tot))

        rec = itc.get(s)
        if rec and rec.get("exempt_ratio", 0) > 0:
            out.append(_ex("INFO", "Rule 42 reversal", s,
                           "Exempt ratio %.4f%%; reversal IGST %.2f / CGST %.2f / SGST %.2f"
                           % (rec["exempt_ratio"] * 100, rec["h_rev42"]["igst"],
                              rec["h_rev42"]["cgst"], rec["h_rev42"]["sgst"])))
    return out
