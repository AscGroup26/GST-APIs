"""GSTR-3B column specifications and constants — Steps 1–4."""
from .utils import GST_STATE_CODE_TO_NAME, VOUCHER_STATE_ALIASES

# ─────────────────────────────────────────────────────────────────
# State names — reuse the project's Title-Case convention
# ─────────────────────────────────────────────────────────────────
CODE_BY_STATE = {v: k for k, v in GST_STATE_CODE_TO_NAME.items()}

EXTRA_STATE_ALIASES = {
    "JAMMU & KASHMIR": "Jammu and Kashmir",
    "ORISSA": "Odisha",
    "PONDICHERRY": "Puducherry",
    "UTTARANCHAL": "Uttarakhand",
    "NEW DELHI": "Delhi",
    "ANDAMAN & NICOBAR": "Andaman and Nicobar Islands",
    "DADRA & NAGAR HAVELI": "Dadra and Nagar Haveli",
}

# ─────────────────────────────────────────────────────────────────
# STEP 1 — Sale Summary (combined GSTR-1)
# ─────────────────────────────────────────────────────────────────
TYPE_SALE = "Sale Other"
TYPE_STOCK = "Stock Transfer"
TYPE_RETURN = "Sale Return"
TYPE_ASSET = "Assets Sale"
TYPE_CROSS = "Cross Charge"
KNOWN_TYPES = {TYPE_SALE, TYPE_STOCK, TYPE_RETURN, TYPE_ASSET, TYPE_CROSS}

SALE_SPEC = {
    "state":     ["fstate"],
    "slab":      ["taxslab"],
    "taxable":   ["inv_tot"],        # despite the name this is the taxable value
    "igst":      ["igstamt"],
    "cgst":      ["cgstamt"],
    "sgst":      ["sgstamt"],
    "cess":      ["Cess", "cessamt"],
    "type":      ["TYPE"],
    "category":  ["Category"],
    "g1remarks": ["GSTR-1 Remarks"],
}
SALE_KEYS = ["fstate", "taxslab", "TYPE"]

# Target columns, named after the '3B Liability' layout they came from
STEP1_COLS = ["J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
              "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF",
              "AG", "CO", "CP", "CQ", "CR"]

STEP1_LABELS = {
    "J": "Sales taxable value", "K": "Sales IGST", "L": "Sales CGST", "M": "Sales SGST",
    "N": "Asset sale value", "O": "Asset IGST", "P": "Asset CGST", "Q": "Asset SGST",
    "R": "Sales+Asset exempt", "S": "Stock transfer value", "T": "Stock transfer IGST",
    "U": "Stock transfer exempt", "V": "Sale return value", "W": "Sale return IGST",
    "X": "Sale return CGST", "Y": "Sale return SGST", "Z": "Sale return exempt",
    "AA": "Net taxable sales", "AB": "Net exempt sales", "AC": "Total sales",
    "AD": "Liability IGST", "AE": "Liability CGST", "AF": "Liability SGST",
    "AG": "Liability Cess", "CO": "Cross charge value", "CP": "Cross charge IGST",
    "CQ": "Cross charge CGST", "CR": "Cross charge SGST",
}

# (TYPE, is_exempt) -> value column + (igst, cgst, sgst) columns.
# Stock transfer is IGST-only (inter-state branch transfer between distinct
# persons). Asset sales share the exempt column R with regular sales.
STEP1_BUCKETS = {
    (TYPE_SALE,   False): {"value": "J",  "tax": ("K", "L", "M")},
    (TYPE_SALE,   True):  {"value": "R",  "tax": None},
    (TYPE_ASSET,  False): {"value": "N",  "tax": ("O", "P", "Q")},
    (TYPE_ASSET,  True):  {"value": "R",  "tax": None},
    (TYPE_STOCK,  False): {"value": "S",  "tax": ("T", None, None)},
    (TYPE_STOCK,  True):  {"value": "U",  "tax": None},
    (TYPE_RETURN, False): {"value": "V",  "tax": ("W", "X", "Y")},
    (TYPE_RETURN, True):  {"value": "Z",  "tax": None},
    (TYPE_CROSS,  False): {"value": "CO", "tax": ("CP", "CQ", "CR")},
    (TYPE_CROSS,  True):  {"value": "CO", "tax": None},
}

# ─────────────────────────────────────────────────────────────────
# STEP 2 — Books / stock / ISD / expenses
# ─────────────────────────────────────────────────────────────────
BOOKS_SPEC = {
    "category":    ["Category"],
    "state":       ["State", "StateBOOKS"],
    "sgst":        ["SGST BOOKS"],
    "cgst":        ["CGST BOOKS"],
    "igst":        ["IGST BOOKS"],
    "cess":        ["Cess tax"],
    "remarks":     ["Remarks"],
    "month_taken": ["MONTH OF ITC TAKEN", "Month of ITC taken"],
}
BOOKS_KEYS = ["SGST BOOKS", "CGST BOOKS", "IGST BOOKS"]

# The two adjacent financial years this app tracks — matches the file-naming
# convention of the raw Books uploads ("Books FY 2025-26" / "Books FY 2026-27").
# Update these when rolling the whole app forward to a new pair of years.
FY_START_A = 2025
FY_START_B = 2026

# 'ITC as per books' — optional pre-aggregated state-wise summary, an
# alternative to uploading the raw Books FY 25-26 / FY 26-27 workbooks.
# One row per state per FY; Remarks names the FY or carries a taken-date.
# Covers Expense + Purchase ITC only — Import is tracked separately.
ITC_SUMMARY_SPEC = {
    "state":   ["State"],
    "igst":    ["IGST"],
    "cgst":    ["CGST"],
    "sgst":    ["SGST"],
    "cess":    ["Cess"],
    "remarks": ["Remarks"],
}
ITC_SUMMARY_KEYS = ["State", "IGST", "CGST", "SGST", "Remarks"]
CAT_IMPORT = "IMPORT"
CAT_EXPENSE = "EXPENSE"

STOCK_SPEC = {"state": ["Ship_State"], "total": ["Total Input in March"]}
STOCK_KEYS = ["Ship_State"]

# The ISD sheet repeats the header 'IGST' for both the invoice split (D/E/F)
# and the DISTRIBUTED amount (O/P/Q). Step 2 needs the distributed columns,
# so they are addressed by position (0-based).
ISD_KEYS = ["State", "Avg Turnover"]
ISD_POS = {"state": 0, "igst": 14, "cgst": 15, "sgst": 16}

# Final Expenses layout (header row 2), 0-based positions:
#   A State | B RCM-GTA net, C/D/E SGST/CGST/IGST 5%
#           | F RCM-other net, G/H/I SGST/CGST/IGST 18%
#           | J RCM-payment-only net, K/L/M SGST/CGST/IGST 5%
# RCM liability (cash) = B+F+J blocks.  RCM credit = B+F blocks only —
# the 'payment only' block carries no ITC (blocked under s.17(5)).
EXPENSE_KEYS = ["State Name"]
EXPENSE_POS = {
    "state": 0,
    "gta_net": 1, "gta_sgst": 2, "gta_cgst": 3, "gta_igst": 4,
    "oth_net": 5, "oth_sgst": 6, "oth_cgst": 7, "oth_igst": 8,
    "pay_net": 9, "pay_sgst": 10, "pay_cgst": 11, "pay_igst": 12,
}

CROSS_KEYS = ["Row Labels", "inv_tot"]
PMT_SPEC = {"state": ["State"], "igst": ["IGST"], "cgst": ["CGST"],
            "sgst": ["SGST"], "cess": ["cess", "Cess", "CESS"]}
PMT_KEYS = ["State"]

# ─────────────────────────────────────────────────────────────────
# STEP 4 — GSTR-2B
# ─────────────────────────────────────────────────────────────────
GSTR2B_SPEC = {
    "month":     ["MONTH"],
    "state":     ["STATE"],
    "taxable":   ["Taxable Value"],
    "igst":      ["Integrated Tax"],
    "cgst":      ["Central Tax"],
    "sgst":      ["State-UT Tax"],
    "cess":      ["Cess tax"],
    "eligible":  ["Eligible", "Eligibility"],
    "addremark": ["Additional remark_Books"],
    "taken3b":   ["Taken in 3B"],
    "reclaimed": ["Reclaimed"],
    # Column AI - the sheet flags each row's other-reversal month directly, so
    # 4B(2) is read off it rather than derived.
    "otherrev":  ["Other Reversal"],
}
GSTR2B_KEYS = ["MONTH", "STATE", "GSTIN NO"]
CONSIDERED = "CONSIDERED"
INELIGIBLE = "INELIGIBLE"

# ─────────────────────────────────────────────────────────────────
# Tolerances
# ─────────────────────────────────────────────────────────────────
TOLERANCE = 1.00
DEFAULT_IGST_SPLIT = 0.50      # residual IGST credit applied to CGST vs SGST
HEADS = ["igst", "cgst", "sgst", "cess"]
HEAD_LABEL = {"igst": "IGST", "cgst": "CGST", "sgst": "SGST", "cess": "Cess"}
RULE_86B_THRESHOLD = 5_000_000  # monthly taxable turnover above which 1% cash applies

# Rule 42 expense pivot from the 'Reversal on Exempted Sales' workbooks. Not
# the same as the books Category='Expense' subtotal - the books remark filter
# drops Rule 37 reclaim lines the reversal working carries, which understates
# the CGST/SGST reversal.
REV42_EXPENSE_SPEC = {
    "state": ["Row Labels", "State"],
    "igst":  ["Sum of IGST BOOKS", "IGST BOOKS", "IGST"],
    "cgst":  ["Sum of CGST BOOKS", "CGST BOOKS", "CGST"],
    "sgst":  ["Sum of SGST BOOKS", "SGST BOOKS", "SGST"],
    "cess":  ["Sum of Cess tax", "Cess tax", "Cess"],
}
REV42_EXPENSE_KEYS = ["Row Labels", "Sum of IGST BOOKS"]
