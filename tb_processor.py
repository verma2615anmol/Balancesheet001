"""
Trial Balance → Balance Sheet Processor
Reads a trial balance, classifies accounts under BS/P&L heads,
and injects aggregated values into a BS template.
Zero formatting change in the output BS file.
Memory-efficient: single workbook open, read_only where possible.
"""

import re
import os
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from copy import copy


# ═══════════════════════════════════════════════════════════════════════
# ACCOUNT CLASSIFICATION RULES
# ═══════════════════════════════════════════════════════════════════════

BS_HEADS = {
    # --- LIABILITIES ---
    "capital": {
        "label": "Owner's Capital / Partners Capital",
        "side": "liability",
        "keywords": [
            "capital", "partner", "proprietor", "owner", "equity",
            "share capital", "reserves", "surplus", "retained earning",
            "general reserve", "capital reserve", "securities premium",
            "profit & loss", "profit and loss", "p&l appropriation",
            "current account", "drawing", "partner current",
            "capital account", "share application",
        ],
        "negative_keywords": ["capital gain", "capital goods", "working capital loan"],
    },
    "lt_borrowings": {
        "label": "Long Term Borrowings",
        "side": "liability",
        "keywords": [
            "long term loan", "term loan", "secured loan", "unsecured loan",
            "mortgage", "debenture", "long term borrowing",
            "loan from bank", "loan from director", "loan from partner",
            "vehicle loan", "car loan", "home loan", "housing loan",
            "hypothecation", "loan payable",
        ],
        "negative_keywords": ["short term", "od", "overdraft", "cc limit"],
    },
    "st_borrowings": {
        "label": "Short Term Borrowings",
        "side": "liability",
        "keywords": [
            "short term loan", "overdraft", "od account", "cash credit",
            "cc limit", "cc account", "working capital", "packing credit",
            "bank od", "bank overdraft", "short term borrowing",
        ],
        "negative_keywords": [],
    },
    "trade_payables": {
        "label": "Trade Payables",
        "side": "liability",
        "keywords": [
            "trade payable", "creditor", "sundry creditor",
            "accounts payable", "supplier", "purchase payable",
            "bills payable", "trade creditor",
        ],
        "negative_keywords": [],
    },
    "other_cl": {
        "label": "Other Current Liabilities",
        "side": "liability",
        "keywords": [
            "other current liabilit", "statutory", "tds payable",
            "gst payable", "gst output", "output cgst", "output sgst",
            "output igst", "tax payable", "duty payable", "cess payable",
            "salary payable", "wages payable", "rent payable",
            "interest payable", "expense payable", "outstanding",
            "advance from customer", "customer advance", "security deposit received",
            "audit fee payable", "professional fee payable",
            "electricity payable", "telephone payable",
            "provision for expense", "payable",
            "income received in advance", "unearned",
        ],
        "negative_keywords": ["provision for tax", "provision for depreciation", "provision for bad debt"],
    },
    "st_provisions": {
        "label": "Short Term Provisions",
        "side": "liability",
        "keywords": [
            "provision for tax", "provision for income tax",
            "provision for depreciation", "provision for bad debt",
            "provision for doubtful", "short term provision",
            "provision for gratuity", "provision for bonus",
            "provision for leave", "provision for warranty",
        ],
        "negative_keywords": [],
    },

    # --- ASSETS ---
    "fixed_assets": {
        "label": "Fixed Assets / PPE",
        "side": "asset",
        "keywords": [
            "fixed asset", "property", "plant", "equipment", "ppe",
            "land", "building", "furniture", "fixture", "vehicle",
            "computer", "machinery", "office equipment", "electrical",
            "air condition", "ac ", "motor car", "scooter", "bike",
            "mobile", "telephone instrument", "printer", "laptop",
            "intangible", "goodwill", "patent", "trademark", "copyright",
            "software", "leasehold", "capital wip", "cwip",
        ],
        "negative_keywords": [
            "depreciation", "accumulated", "provision for",
            "repair", "maintenance", "rent",
        ],
    },
    "non_current_investments": {
        "label": "Non-Current Investments",
        "side": "asset",
        "keywords": [
            "investment", "shares", "debenture held", "mutual fund",
            "fixed deposit", "fdr", "fd ", "nsc ", "kvp",
            "government securities", "bonds",
            "investment in subsidiary", "investment in associate",
        ],
        "negative_keywords": ["provision for investment"],
    },
    "inventories": {
        "label": "Inventories / Stock",
        "side": "asset",
        "keywords": [
            "inventor", "stock", "closing stock", "opening stock",
            "raw material", "finished good", "work in progress",
            "wip", "stores", "spare", "packing material",
            "stock in trade", "goods in transit",
        ],
        "negative_keywords": ["stock broker"],
    },
    "trade_rec": {
        "label": "Trade Receivables",
        "side": "asset",
        "keywords": [
            "trade receivable", "debtor", "sundry debtor",
            "accounts receivable", "bills receivable",
            "trade debtor", "receivable from customer",
        ],
        "negative_keywords": [],
    },
    "cash_bank": {
        "label": "Cash and Bank Balances",
        "side": "asset",
        "keywords": [
            "cash", "bank", "cash in hand", "cash at bank",
            "petty cash", "savings account", "current account bank",
            "bank account", "bank balance", "cheque in hand",
            "imprest",
        ],
        "negative_keywords": ["cash credit", "cc account", "overdraft", "od account", "bank od"],
    },
    "stla": {
        "label": "Short Term Loans & Advances",
        "side": "asset",
        "keywords": [
            "loan given", "advance to", "loan to",
            "advance to supplier", "advance to staff", "staff advance",
            "prepaid", "deposit", "security deposit paid",
            "tds receivable", "tcs receivable", "input tax",
            "input cgst", "input sgst", "input igst", "gst input",
            "advance tax", "self assessment tax", "mat credit",
            "cenvat", "vat input", "excise input",
            "income tax refund", "refund receivable",
            "advance recoverable",
        ],
        "negative_keywords": ["advance from customer", "customer advance"],
    },
    "other_current_assets": {
        "label": "Other Current Assets",
        "side": "asset",
        "keywords": [
            "other current asset", "accrued income", "interest accrued",
            "accrued interest", "interest receivable",
            "other receivable", "other asset",
        ],
        "negative_keywords": [],
    },

    # --- P&L HEADS ---
    "revenue": {
        "label": "Revenue from Operations / Sales",
        "side": "pl",
        "keywords": [
            "sale", "revenue", "income from operation",
            "turnover", "gross receipt", "service income",
            "service revenue", "consulting income", "fee received",
            "commission received", "commission income",
            "export sale", "domestic sale", "local sale",
        ],
        "negative_keywords": ["sale return", "sales return", "sale of asset", "sale of investment"],
    },
    "purchases": {
        "label": "Purchases / Cost of Material",
        "side": "pl",
        "keywords": [
            "purchase", "cost of material", "cost of goods",
            "import purchase", "local purchase", "domestic purchase",
            "raw material consumed", "material consumed",
            "sub contract", "job work", "labour charge",
            "freight inward", "carriage inward", "octroi",
            "custom duty", "clearing charge",
        ],
        "negative_keywords": ["purchase return"],
    },
    "employee_expenses": {
        "label": "Employee / Salary Expenses",
        "side": "pl",
        "keywords": [
            "salary", "wage", "bonus", "gratuity", "leave",
            "staff welfare", "epf", "esi", "pf contribution",
            "employee benefit", "director remuneration",
            "partner salary", "partner remuneration",
            "stipend", "incentive", "overtime",
        ],
        "negative_keywords": ["salary payable", "wages payable"],
    },
    "other_expenses": {
        "label": "Other Expenses / Indirect Expenses",
        "side": "pl",
        "keywords": [
            "expense", "rent", "electricity", "telephone",
            "internet", "travelling", "conveyance", "vehicle running",
            "petrol", "diesel", "fuel", "repair", "maintenance",
            "insurance", "audit fee", "professional fee", "legal fee",
            "printing", "stationery", "postage", "courier",
            "advertisement", "marketing", "donation",
            "interest paid", "interest on loan", "bank charge",
            "bank interest", "discount allowed", "bad debt",
            "miscellaneous", "office expense", "general expense",
            "entertainment", "subscription", "membership",
            "rate", "tax", "municipal", "water charge",
            "loading", "unloading", "packing",
            "commission paid", "brokerage", "agency",
            "foreign exchange loss", "exchange diff",
            "penalty", "fine", "late fee", "round off",
            "festival", "gift", "welfare",
        ],
        "negative_keywords": [
            "salary", "wage", "depreciation", "purchase",
            "payable", "outstanding", "provision",
        ],
    },
    "depreciation": {
        "label": "Depreciation",
        "side": "pl",
        "keywords": [
            "depreciation", "amortisation", "amortization",
            "dep on", "accumulated depreciation",
        ],
        "negative_keywords": ["provision for depreciation"],
    },
}

# Priority order for classification (most specific first)
CLASSIFICATION_PRIORITY = [
    "depreciation",
    "trade_payables", "trade_rec",
    "cash_bank", "inventories",
    "st_provisions",
    "st_borrowings", "lt_borrowings",
    "employee_expenses", "purchases", "revenue",
    "fixed_assets", "non_current_investments",
    "stla",
    "capital",
    "other_cl", "other_current_assets",
    "other_expenses",
]


# ═══════════════════════════════════════════════════════════════════════
# TB AUTO-DETECTION ENGINE
# ═══════════════════════════════════════════════════════════════════════

def detect_tb_structure(file_path):
    """
    Auto-detect the structure of a trial balance file.
    Returns dict with: format_type, header_row, data_start_row,
    account_col, debit_col, credit_col, net_col, sheet_name, accounts[]
    """
    wb = load_workbook(file_path, read_only=True, data_only=True)
    results = []

    for sname in wb.sheetnames:
        ws = wb[sname]
        rows_data = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            rows_data.append(list(row))
            if i >= 2000:   # handle large TBs up to 2000 rows
                break

        if len(rows_data) < 2:
            continue

        det = _detect_columns(rows_data, sname)
        if det:
            results.append(det)

    wb.close()

    if not results:
        return {"error": "Could not detect trial balance structure. Please check the file format."}

    # Pick the best detection (most accounts found)
    best = max(results, key=lambda d: len(d.get("accounts", [])))
    return best


def _detect_columns(rows, sheet_name):
    """Detect which column is what in the TB."""
    header_row = None
    acct_col = None
    dr_col = None
    cr_col = None
    net_col = None
    format_type = None

    # Scan first 15 rows for headers
    for ri, row in enumerate(rows[:15]):
        if not row:
            continue
        row_lower = [str(c).strip().lower() if c else "" for c in row]

        # Look for account name column
        for ci, val in enumerate(row_lower):
            if not val:
                continue

            # Account name indicators
            if any(k in val for k in [
                "particular", "account", "ledger", "name", "head",
                "description", "party name",
            ]):
                if acct_col is None:
                    acct_col = ci
                    header_row = ri

            # Debit column — handle "Debit (₹)", "Dr.", "Debit Amount" etc
            if (val in ("dr", "debit", "dr balance", "debit balance",
                       "dr amount", "debit amount", "dr bal", "debit bal") or
                    val.startswith("debit") or val.startswith("dr ") or
                    val == "dr." or "debit" in val.split("(")[0].strip()):
                dr_col = ci

            # Credit column — handle "Credit (₹)", "Cr.", "Credit Amount" etc
            if (val in ("cr", "credit", "cr balance", "credit balance",
                       "cr amount", "credit amount", "cr bal", "credit bal") or
                    val.startswith("credit") or val.startswith("cr ") or
                    val == "cr." or "credit" in val.split("(")[0].strip()):
                cr_col = ci

            # Net/Amount column
            if val in ("amount", "net balance", "net amount", "balance",
                       "closing balance", "closing", "net"):
                net_col = ci

            # Opening column (Type 3)
            if val in ("opening", "opening balance"):
                pass  # We mostly care about closing

            # Closing column (Type 3)
            if val in ("closing", "closing balance"):
                if net_col is None:
                    net_col = ci

    # If no header found, try heuristic: first column with text, next columns with numbers
    if acct_col is None:
        for ri, row in enumerate(rows[:15]):
            if not row:
                continue
            text_cols = []
            num_cols = []
            for ci, val in enumerate(row):
                if val is None:
                    continue
                if isinstance(val, str) and len(val.strip()) > 2 and not _is_number_str(val):
                    text_cols.append(ci)
                elif isinstance(val, (int, float)) or _is_number_str(str(val)):
                    num_cols.append(ci)
            if len(text_cols) >= 1 and len(num_cols) >= 1:
                acct_col = text_cols[0]
                header_row = ri
                break

    if acct_col is None:
        return None

    # Determine format type
    # Re-check: if both dr and cr columns found but assigned same index, fix
    if dr_col is not None and cr_col is not None and dr_col != cr_col:
        format_type = 1  # Dr/Cr separate columns
    elif dr_col is not None and cr_col is not None and dr_col == cr_col:
        # Conflict - reset and try harder
        dr_col = None; cr_col = None
        format_type = None
    elif net_col is not None:
        format_type = 4  # Single amount (negative = credit)
    else:
        # Try to auto-detect number columns after the account column
        for ri, row in enumerate(rows[header_row + 1: header_row + 10], header_row + 1):
            if not row:
                continue
            num_cols_found = []
            for ci, val in enumerate(row):
                if ci == acct_col:
                    continue
                if isinstance(val, (int, float)) and val != 0:
                    num_cols_found.append(ci)
                elif isinstance(val, str) and _is_number_str(val):
                    num_cols_found.append(ci)
            if len(num_cols_found) >= 2:
                dr_col = num_cols_found[0]
                cr_col = num_cols_found[1]
                format_type = 1
                break
            elif len(num_cols_found) == 1:
                net_col = num_cols_found[0]
                format_type = 4
                break

    if format_type is None:
        return None

    # Fix: if format detected as 4 (single net column) but there are actually
    # two number columns (like Debit col A and Credit col B in this TB),
    # detect that and switch to format 1
    if format_type == 4 and net_col is not None and dr_col is None:
        # Check if there are two numeric columns around the net_col
        for ri2 in range(header_row + 1, min(header_row + 10, len(rows))):
            row2 = rows[ri2]
            if not row2: continue
            num_cols_found = []
            for ci2, v2 in enumerate(row2):
                if ci2 == acct_col: continue
                if isinstance(v2, (int, float)) and v2 != 0:
                    num_cols_found.append(ci2)
            if len(num_cols_found) >= 2:
                # Two numeric columns found — treat as Dr/Cr
                format_type = 1
                dr_col = num_cols_found[0]
                cr_col = num_cols_found[1]
                net_col = None
                break

    # Data starts after header
    data_start = header_row + 1

    # Extract accounts — handles both flat and hierarchical TB formats
    accounts = []
    total_keywords = {"total", "grand total", "difference", "net total",
                      "closing balance", "opening balance total",
                      "balance c/d", "balance b/d"}

    current_group = None  # Track current group header for hierarchical TBs

    for ri in range(data_start, len(rows)):
        row = rows[ri]
        if not row or ri >= len(rows):
            continue
        acct_name = row[acct_col] if acct_col < len(row) else None
        if not acct_name or not isinstance(acct_name, str):
            continue
        acct_name = acct_name.strip()
        if not acct_name or len(acct_name) < 2:
            continue

        # Skip totals/subtotals
        if acct_name.lower().strip() in total_keywords:
            continue
        if re.match(r'^(total|grand total|sub total|net total)\b', acct_name, re.I):
            continue

        # Get amounts
        dr_amt = 0
        cr_amt = 0
        net_amt = 0

        if format_type == 1 and dr_col is not None and cr_col is not None:
            dr_val = row[dr_col] if dr_col < len(row) else None
            cr_val = row[cr_col] if cr_col < len(row) else None
            dr_amt = _to_float(dr_val)
            cr_amt = _to_float(cr_val)
            net_amt = dr_amt - cr_amt
        elif format_type == 4 and net_col is not None:
            # For hierarchical format: check BOTH debit and credit columns
            # even if format_type was detected as 4
            if dr_col is None and cr_col is None:
                # Try columns 1 and 2 as dr/cr if they have data
                dr_try = row[1] if len(row) > 1 else None
                cr_try = row[2] if len(row) > 2 else None
                dr_f = _to_float(dr_try)
                cr_f = _to_float(cr_try)
                if dr_f != 0 or cr_f != 0:
                    dr_amt = dr_f
                    cr_amt = cr_f
                    net_amt = dr_amt - cr_amt
                else:
                    net_val = row[net_col] if net_col < len(row) else None
                    net_amt = _to_float(net_val)
                    if net_amt > 0:
                        dr_amt = net_amt
                    else:
                        cr_amt = abs(net_amt)
            else:
                net_val = row[net_col] if net_col < len(row) else None
                net_amt = _to_float(net_val)
                if net_amt > 0:
                    dr_amt = net_amt
                else:
                    cr_amt = abs(net_amt)

        # Detect group headers: rows with name but zero amounts
        # These help classify sub-items in hierarchical TBs
        if dr_amt == 0 and cr_amt == 0 and net_amt == 0:
            # This is likely a group/category header
            current_group = acct_name
            continue

        accounts.append({
            "row": ri,
            "name": acct_name,
            "group": current_group,  # parent group for classification hint
            "debit": dr_amt,
            "credit": cr_amt,
            "net": net_amt,
        })

    return {
        "format_type": format_type,
        "sheet_name": sheet_name,
        "header_row": header_row,
        "data_start_row": data_start,
        "account_col": acct_col,
        "debit_col": dr_col,
        "credit_col": cr_col,
        "net_col": net_col,
        "accounts": accounts,
    }


def _is_number_str(s):
    """Check if a string looks like a number."""
    if not s:
        return False
    s = s.strip().replace(",", "").replace("(", "-").replace(")", "")
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False


def _to_float(val):
    """Convert a value to float, handling strings and None."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        s = val.strip().replace(",", "").replace("₹", "").replace("Rs.", "").replace("Rs", "")
        # Handle parentheses = negative
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        s = s.strip()
        try:
            return float(s)
        except (ValueError, TypeError):
            return 0.0
    return 0.0


# ═══════════════════════════════════════════════════════════════════════
# ACCOUNT CLASSIFIER
# ═══════════════════════════════════════════════════════════════════════

# Group header → BS head mapping for hierarchical TBs
GROUP_HEAD_MAP = {
    "capital account": "capital",
    "bank accounts": "cash_bank",
    "bank account": "cash_bank",
    "cash-in-hand": "cash_bank",
    "cash in hand": "cash_bank",
    "fixed assets": "fixed_assets",
    "sundry creditors": "trade_payables",
    "sundry debtors": "trade_rec",
    "sundry debtor": "trade_rec",
    "purchase account": "purchases",
    "purchases": "purchases",
    "sales account": "revenue",
    "stock-in-hand": "inventories",
    "stock in hand": "inventories",
    "indirect expenses": "other_expenses",
    "direct expenses": "purchases",
    "sundry payables": "other_cl",
    "provisions": "st_provisions",
    "unsecure loans": "lt_borrowings",
    "unsecured loans": "lt_borrowings",
    "deposits (asset)": "stla",
    "duties & taxes": "stla",
    "duties and taxes": "stla",
}


def classify_accounts(accounts):
    """
    Classify each TB account under a BS/P&L head.
    Uses group headers as classification hints for hierarchical TBs.
    Returns list of dicts with added 'bs_head' and 'confidence' keys.
    """
    classified = []
    for acct in accounts:
        name = acct["name"]
        group = acct.get("group", "")
        bs_head, confidence = _classify_single(name, acct["net"], group)
        acct_copy = dict(acct)
        acct_copy["bs_head"] = bs_head
        acct_copy["confidence"] = confidence
        classified.append(acct_copy)
    return classified


def _classify_single(name, net_amount, group=None):
    """Classify a single account name. Returns (head_key, confidence)."""
    name_lower = name.lower().strip()
    group_lower = (group or "").lower().strip()

    # Step 1: Check if group header directly maps to a head
    if group_lower and group_lower in GROUP_HEAD_MAP:
        group_head = GROUP_HEAD_MAP[group_lower]
        # For items under a known group, trust the group mapping
        # but still verify it makes sense with the amount sign
        return group_head, "high"

    # Step 2: Try each head in priority order by name
    for head_key in CLASSIFICATION_PRIORITY:
        head = BS_HEADS[head_key]
        # Check negative keywords first (exclusions)
        excluded = False
        for nk in head.get("negative_keywords", []):
            if nk in name_lower:
                excluded = True
                break
        if excluded:
            continue

        # Check positive keywords
        for kw in head["keywords"]:
            if kw in name_lower:
                return head_key, "high"

    # Step 3: Partial group match
    if group_lower:
        for grp_key, head_key in GROUP_HEAD_MAP.items():
            if grp_key in group_lower or group_lower in grp_key:
                return head_key, "low"

    # Step 4: Use net amount sign as hint
    # Debit balance (positive net) → likely asset or expense
    # Credit balance (negative net) → likely liability or income
    if net_amount > 0:
        return "other_current_assets", "low"
    elif net_amount < 0:
        return "other_cl", "low"
    else:
        return "unclassified", "none"


def get_aggregated_values(classified_accounts):
    """
    Aggregate classified accounts into BS head totals.
    Returns dict: {head_key: total_amount}
    """
    totals = {}
    for acct in classified_accounts:
        head = acct["bs_head"]
        if head == "unclassified":
            continue
        amt = abs(acct["net"])
        if head not in totals:
            totals[head] = 0
        totals[head] += amt
    return totals


# ═══════════════════════════════════════════════════════════════════════
# BS TEMPLATE INJECTION — Notes-Aware Engine
# ═══════════════════════════════════════════════════════════════════════
# This BS template is 100% formula-driven on the bs sheet.
# Every cell like E7, E11, E17 etc. on the bs sheet pulls from:
#   capital!G10       → owners capital closing balance
#   'notes to bs'!D20 → long-term borrowings total
#   'notes to bs'!D48 → trade payables (creditors total)
#   'notes to bs'!D69 → other current liabilities total
#   'Fixed Assets C. Yr.'!I31 → PPE net block
#   'notes to p&l'!D24 → inventories (closing stock)
#   'notes to bs'!D104 → trade receivables total
#   'notes to bs'!D116 → cash & bank total
#   'notes to bs'!D133 → short-term loans & advances
#
# Strategy: inject into the SOURCE sheets (the plain-value cells),
# never into the formula cells on the bs sheet.
# The formulas then auto-propagate values to bs and other sheets.

def _is_formula(val):
    return isinstance(val, str) and val.strip().startswith("=")

def _safe_set(ws, row, col, value):
    """Write a plain numeric value only if the target cell is not a formula."""
    cell = ws.cell(row=row, column=col)
    if not _is_formula(cell.value):
        cell.value = round(float(value), 2)
        return True
    return False


def _detect_notes_structure(wb):
    """
    Scan the Notes sheets and build an injection map:
    { head_key: [(sheet_name, row, col, label), ...] }
    This maps each BS head to the plain-value target cell(s) where
    we should write the aggregated amount.

    Returns (injection_map, details) where details contains per-head
    info about what was found.
    """
    # ── Capital sheet ────────────────────────────────────────────────
    # The capital closing balance is computed by the formula:
    #   G10 = C10+D10-E10+F10   (opening + intro - drawings + profit)
    # We can only inject into C8 (opening), D8 (intro), E8 (drawings).
    # Profit (F8) comes from p&l sheet — we cannot inject that.
    # For a fresh TB injection the cleanest approach is:
    #   Write the closing balance from TB directly into a single
    #   "lump sum" row that doesn't break the formula structure.
    # We inject into capital!D8 (Capital Introduced) as a net plug
    # if the opening balance is already in C8 via =G11 formula.
    # In practice: set capital!D8 = (TB_closing - capital!C8_resolved)
    # But since C8=G11=prev year closing which is read-only, safest is:
    # Write total capital to capital!G10 if it's not a formula,
    # else find a writable row.

    capital_map = []
    if "capital" in wb.sheetnames:
        ws_cap = wb["capital"]
        # Scan rows 8-15 for plain-value cells in column G (closing bal)
        for r in range(8, 16):
            v = ws_cap.cell(r, 7).value
            if v is not None and not _is_formula(v):
                capital_map.append(("capital", r, 7, "Closing Balance"))
                break
        # If G10 is a formula (=C10+D10-E10+F10), inject into D8 (Capital Introduced)
        # after zeroing withdrawals so closing = opening + introduced
        if not capital_map:
            # We'll use a special "capital_plug" strategy
            capital_map.append(("capital", "plug", None, "Capital plug via D8"))

    # ── notes to bs ─────────────────────────────────────────────────
    notes_map = {}
    if "notes to bs" in wb.sheetnames:
        ws_n = wb["notes to bs"]
        for row in ws_n.iter_rows(min_row=1, max_row=200, min_col=2, max_col=4, values_only=True):
            b_val, _, d_val = row
            if b_val is not None:
                r = ws_n.cell(1,1).row  # placeholder
                break
        # Fast scan using iter_rows
        for r_idx, row in enumerate(ws_n.iter_rows(min_row=1, max_row=200,
                                                     min_col=2, max_col=4), 1):
            b_val = row[0].value; d_val = row[2].value
            label = str(b_val).strip().lower() if b_val else ""
            if d_val is not None and not _is_formula(d_val):
                notes_map[r_idx] = (label, d_val)

    # ── Details sheet ────────────────────────────────────────────────
    details_map = {}
    if "Details" in wb.sheetnames:
        ws_det = wb["Details"]
        for r_idx, row in enumerate(ws_det.iter_rows(min_row=1, max_row=200,
                                                       min_col=2, max_col=4), 1):
            b_val = row[0].value; d_val = row[2].value
            if b_val is not None and not _is_formula(d_val if d_val is not None else "x"):
                details_map[r_idx] = (str(b_val).strip().lower(), d_val)

    # ── GROSS PROFIT / notes to p&l for inventory ────────────────────
    gp_map = {}
    if "GROSS PROFIT" in wb.sheetnames:
        ws_gp = wb["GROSS PROFIT"]
        for r in range(1, 30):
            b_val = ws_gp.cell(r, 4).value  # E-side labels
            e_val = ws_gp.cell(r, 5).value
            if b_val and "closing stock" in str(b_val).lower() and e_val is not None and not _is_formula(e_val):
                gp_map["closing_stock"] = ("GROSS PROFIT", r, 5)
                break

    return {
        "capital_map": capital_map,
        "notes_map": notes_map,
        "details_map": details_map,
        "gp_map": gp_map,
    }


def _fuzzy_match_name(tb_name, template_name):
    """Check if two party names roughly match (for creditor/debtor matching)."""
    import re
    def normalize(s):
        s = s.lower()
        s = re.sub(r'\bm/s\.?\s*', '', s)
        s = re.sub(r'[^a-z0-9\s]', ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
        for city in ['ludhiana', 'delhi', 'jalandhar', 'surat', 'ahmedabad',
                     'ahemadabad', 'varanasi', 'mumbai', 'ambala', 'citi']:
            s = re.sub(r'\b' + city + r'\b', '', s).strip()
        s = re.sub(r'\s+', ' ', s).strip()
        return s

    a = normalize(tb_name)
    b = normalize(template_name)
    if not a or not b:
        return False
    # Exact after normalization
    if a == b:
        return True
    # Only allow substring if the shorter is at least 6 chars
    # AND the match is not just a common word like 'textiles'
    COMMON_WORDS = {'textiles', 'trading', 'enterprises', 'creation', 'fashion',
                    'fabrics', 'industries', 'pvt', 'ltd', 'co', 'and', 'sons'}
    if len(a) >= 6 and len(b) >= 6:
        shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
        if shorter in longer:
            # Verify the match isn't just on common words
            unique_words = [w for w in shorter.split() if w not in COMMON_WORDS and len(w) > 3]
            if unique_words:
                return True
    # Check first 2+ significant words match exactly
    words_a = [w for w in a.split() if len(w) > 3 and w not in COMMON_WORDS]
    words_b = [w for w in b.split() if len(w) > 3 and w not in COMMON_WORDS]
    if words_a and words_b:
        common = sum(1 for w in words_a if w in words_b)
        if common >= min(2, len(words_a), len(words_b)):
            return True
    # Single distinctive word match (>=7 chars, not common)
    if words_a and words_b:
        long_a = [w for w in words_a if len(w) >= 7]
        long_b = [w for w in words_b if len(w) >= 7]
        if any(w in long_b for w in long_a):
            return True
    return False


def _col_letter(col_idx):
    """Convert 0-indexed column to letter."""
    from openpyxl.utils import get_column_letter
    return get_column_letter(col_idx + 1)



def _load_sheet_cache(wb, sheet_name, max_row=200, max_col=10):
    """Pre-load all cell values into dict for fast O(1) lookup."""
    cache = {}
    if sheet_name not in wb.sheetnames:
        return cache
    ws = wb[sheet_name]
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            if cell.value is not None:
                cache[(cell.row, cell.column)] = cell.value
    return cache


def inject_into_bs(bs_template_path, output_path, aggregated_values,
                   mapping_overrides=None, individual_accounts=None):
    """
    Inject aggregated TB values into BS template by writing to SOURCE sheets.

    The BS sheet itself is all formulas — writing there does nothing.
    Instead we write to:
      capital      → owner's capital closing balance
      notes to bs  → long-term loans, trade payables, OCL, cash, receivables,
                     short-term loans
      Details      → individual creditor/debtor amounts (matched by name)
      GROSS PROFIT → closing stock (inventories)
      Fixed Assets C. Yr. → asset additions (col C = >180 days)

    individual_accounts: list of {name, bs_head, net, debit, credit}
    from the classified TB accounts — used for name-matching creditors/debtors.
    """
    import shutil
    log = []

    # Copy template to output first
    shutil.copy2(bs_template_path, output_path)

    wb = load_workbook(output_path)
    structure = _detect_notes_structure(wb)
    injected = []
    skipped = []

    # Pre-load sheet caches for fast lookup (avoids slow ws.cell(r,c) in loops)
    _cache = {}
    for _sn in ["notes to bs", "notes to p&l", "Details", "GROSS PROFIT",
                "Fixed Assets C. Yr.", "capital"]:
        _cache[_sn] = _load_sheet_cache(wb, _sn, max_row=250, max_col=12)

    # ────────────────────────────────────────────────────────────────
    # 1. CAPITAL — inject closing balance
    # ────────────────────────────────────────────────────────────────
    capital_amt = aggregated_values.get("capital", 0)
    if capital_amt and "capital" in wb.sheetnames:
        ws_cap = wb["capital"]
        # Strategy: The closing balance formula is =C10+D10-E10+F10
        # C10=SUM(C8:C9)=opening (from PY closing via =G11)
        # D10=SUM(D8:D9)=introduced, E10=SUM(E8:E9)=withdrawals, F10=profit from p&l
        # We cannot break the formula chain. Best approach:
        # Set D8 (Capital Introduced during year) = closing_from_TB - C8 - F8 + E8
        # But F8 = profit which is a formula from p&l.
        # Cleanest: just write to D8 as a "net plug" = capital_amt - (C8_value or 0)
        # After user processes p&l separately, D8 will reconcile.
        # For now, write closing directly to G8 if not formula, else use D8 plug.
        g10 = ws_cap.cell(10, 7).value
        c8 = ws_cap.cell(8, 3).value
        opening = 0
        if c8 is not None and not _is_formula(c8):
            opening = float(c8) if c8 else 0
        elif _is_formula(str(c8 or "")):
            # Try to read G11 (PY closing) as opening
            g11 = ws_cap.cell(11, 7).value
            if g11 is not None and not _is_formula(g11):
                opening = float(g11) if g11 else 0

        if _is_formula(str(g10 or "")):
            # G10 is formula — inject via D8: Introduced = closing - opening
            # (sets withdrawals and profit contribution aside for now)
            e8 = ws_cap.cell(8, 5).value
            f8 = ws_cap.cell(8, 6).value
            withdrawals = float(e8) if e8 and not _is_formula(str(e8)) else 0
            profit = 0  # formula, skip
            d8_val = capital_amt - opening + withdrawals - profit
            if _safe_set(ws_cap, 8, 4, max(0, d8_val)):
                injected.append(f"capital!D8 (Capital Introduced) = {d8_val:,.2f} → closing balance will compute to {capital_amt:,.2f}")
            else:
                skipped.append(f"capital!D8 is a formula — could not inject capital {capital_amt:,.2f}")
        else:
            if _safe_set(ws_cap, 10, 7, capital_amt):
                injected.append(f"capital!G10 = {capital_amt:,.2f}")
            elif _safe_set(ws_cap, 8, 7, capital_amt):
                injected.append(f"capital!G8 = {capital_amt:,.2f}")
            else:
                skipped.append(f"capital: could not find writable cell for {capital_amt:,.2f}")

    # ────────────────────────────────────────────────────────────────
    # 2. NOTES TO BS — inject by scanning plain-value cells
    # ────────────────────────────────────────────────────────────────
    if "notes to bs" in wb.sheetnames:
        ws_n = wb["notes to bs"]

        def inject_notes_row(row, col, amount, label):
            if _safe_set(ws_n, row, col, amount):
                injected.append(f"notes to bs!{_col_letter(col-1)}{row} ({label}) = {amount:,.2f}")
                return True
            else:
                skipped.append(f"notes to bs R{row}C{col} ({label}) is formula")
                return False

        # Long-term borrowings → D8 (ICICI Bank or first bank row)
        ltb_amt = aggregated_values.get("lt_borrowings", 0)
        if ltb_amt:
            # Find writable bank row in notes to bs rows 7-10
            placed = False
            for r in range(7, 11):
                d = ws_n.cell(r, 4).value
                if d is None or (not _is_formula(str(d))):
                    placed = inject_notes_row(r, 4, ltb_amt, "Long-term borrowings")
                    break
            if not placed:
                skipped.append(f"long_term_borrowings {ltb_amt:,.2f}: no writable row found in notes to bs")

        # Short-term borrowings → D26 (bank CC row)
        stb_amt = aggregated_values.get("st_borrowings", 0)
        if stb_amt:
            placed = False
            for r in range(26, 32):
                d = ws_n.cell(r, 4).value
                if d is None or (not _is_formula(str(d))):
                    placed = inject_notes_row(r, 4, stb_amt, "Short-term borrowings")
                    break
            if not placed:
                skipped.append(f"short_term_borrowings {stb_amt:,.2f}: no writable row in notes to bs")

        # ── Cheques Issued But Not Cleared → D68 (OCL) ─────────────────
        # These are liabilities, not assets — classifier may put them in stla
        if individual_accounts:
            cheque_accounts = [a for a in individual_accounts
                               if "cheque" in a.get("name","").lower()
                               or "chq" in a.get("name","").lower()]
            total_cheques = sum(abs(a["net"]) for a in cheque_accounts)
            if total_cheques > 0:
                if _safe_set(ws_n, 68, 4, total_cheques):
                    injected.append(f"notes to bs!D68 (Cheques issued) = {total_cheques:,.2f}")

        # Other current liabilities → D64:D68 (Audit, Legal, Salary, TDS, Cheque)
        ocl_amt = aggregated_values.get("other_cl", 0)
        if ocl_amt:
            # Collect all writable OCL rows (plain-value or empty, in OCL section)
            ocl_rows = []
            for r in range(60, 82):
                d = ws_n.cell(r, 4).value
                b = ws_n.cell(r, 2).value
                if b is None:
                    continue
                b_str = str(b).strip().lower()
                if any(kw in b_str for kw in ['total', 'particular', 'amount', 'header',
                                               'short-term', 'short term', 'borrowing',
                                               'payable\n', 'current liabilit']):
                    continue
                if len(b_str) < 2:
                    continue
                if d is None or (not _is_formula(str(d))):
                    ocl_rows.append((r, b_str, d))

            if individual_accounts:
                ocl_accounts = [a for a in individual_accounts
                                if a.get("bs_head") == "other_cl"
                                and abs(a.get("net", 0)) > 0]
                row_idx = 0  # sequential pointer into ocl_rows
                for acct in ocl_accounts:
                    if row_idx >= len(ocl_rows):
                        skipped.append(f"OCL '{acct['name']}' = {abs(acct['net']):,.2f}: no more rows")
                        continue
                    r, label, _ = ocl_rows[row_idx]
                    inject_notes_row(r, 4, abs(acct["net"]), f"OCL: {acct['name']}")
                    row_idx += 1
            elif ocl_rows:
                inject_notes_row(ocl_rows[0][0], 4, ocl_amt, "Other current liabilities")

        # Cash → D109 (Cash in hand), HDFC → D113, ICICI → D114
        cash_bank_amt = aggregated_values.get("cash_bank", 0)
        if individual_accounts:
            cash_accounts = [a for a in individual_accounts
                             if a.get("bs_head") in ("cash_bank", "cash_bank")]
            cash_only  = [a for a in cash_accounts if "cash" in a["name"].lower()]
            bank_only  = [a for a in cash_accounts if "cash" not in a["name"].lower()]

            # Cash in hand → D109 directly (plain cell)
            if cash_only:
                total_cash = sum(abs(a["net"]) for a in cash_only)
                if _safe_set(ws_n, 109, 4, total_cash):
                    injected.append(f"notes to bs!D109 (Cash in hand) = {total_cash:,.2f}")

            # Individual bank accounts → D113, D114 by name match
            for acct in bank_only:
                name_l = acct["name"].lower()
                amt = abs(acct["net"])
                if "hdfc" in name_l:
                    if _safe_set(ws_n, 113, 4, amt):
                        injected.append(f"notes to bs!D113 (HDFC) = {amt:,.2f}")
                elif "icici" in name_l:
                    if _safe_set(ws_n, 114, 4, amt):
                        injected.append(f"notes to bs!D114 (ICICI) = {amt:,.2f}")
                elif "pnb" in name_l or "punjab national" in name_l:
                    if _safe_set(ws_n, 113, 4, amt):
                        injected.append(f"notes to bs!D113 (PNB) = {amt:,.2f}")
                elif "sbi" in name_l or "state bank" in name_l:
                    if _safe_set(ws_n, 113, 4, amt):
                        injected.append(f"notes to bs!D113 (SBI) = {amt:,.2f}")
                else:
                    # First empty bank row
                    for r in range(113, 116):
                        if ws_n.cell(r, 4).value is None:
                            if _safe_set(ws_n, r, 4, amt):
                                injected.append(f"notes to bs!D{r} ({acct['name']}) = {amt:,.2f}")
                            break
        elif cash_bank_amt:
            inject_notes_row(109, 4, cash_bank_amt, "Cash and bank (lump)")

        # Short-term loans & advances → D128:D131 (GST, TDS etc)
        stla_amt = aggregated_values.get("stla", 0)
        if stla_amt and individual_accounts:
            stla_accounts = [a for a in individual_accounts
                             if a.get("bs_head") == "stla"
                             and abs(a.get("net", 0)) > 0
                             and "cheque" not in a.get("name","").lower()
                             and "chq" not in a.get("name","").lower()]
            for acct in stla_accounts:
                name_l = acct["name"].lower()
                placed = False
                # Match GST refund → D128
                if "gst" in name_l or "igst" in name_l or "cgst" in name_l or "sgst" in name_l:
                    if _safe_set(ws_n, 128, 4, abs(acct["net"])):
                        injected.append(f"notes to bs!D128 (GST input) = {abs(acct['net']):,.2f}")
                        placed = True
                # Match TCS/TDS → D129/D130
                if "tcs" in name_l:
                    if _safe_set(ws_n, 129, 4, abs(acct["net"])):
                        injected.append(f"notes to bs!D129 (TCS GST) = {abs(acct['net']):,.2f}")
                        placed = True
                if ("tds" in name_l or "excess tds" in name_l) and not placed:
                    if _safe_set(ws_n, 130, 4, abs(acct["net"])):
                        injected.append(f"notes to bs!D130 (Excess TDS) = {abs(acct['net']):,.2f}")
                        placed = True
                if not placed:
                    # Find first empty row in 125-132
                    for r in range(125, 133):
                        d = ws_n.cell(r, 4).value
                        if d is None:
                            if _safe_set(ws_n, r, 4, abs(acct["net"])):
                                injected.append(f"notes to bs!D{r} ({acct['name']}) = {abs(acct['net']):,.2f}")
                            break
        elif stla_amt:
            inject_notes_row(128, 4, stla_amt, "Short-term loans & advances")

    # ────────────────────────────────────────────────────────────────
    # 3. DETAILS — individual creditor/debtor amounts (CACHE-BASED)
    # ────────────────────────────────────────────────────────────────
    if "Details" in wb.sheetnames and individual_accounts:
        ws_det  = wb["Details"]
        det_cache = _cache.get("Details", {})

        # Build name→row map from cache (instant, no cell access)
        cred_row_map = {}
        for r in range(21, 63):
            b = det_cache.get((r, 2))
            if b and str(b).strip():
                cred_row_map[str(b).strip()] = r

        written_rows = set()
        creditor_accounts = [a for a in individual_accounts
                             if a.get("bs_head") == "trade_payables"
                             and abs(a.get("net", 0)) > 0]
        unmatched_creditors = []
        for acct in creditor_accounts:
            best_row = None
            for tmpl_name, row in cred_row_map.items():
                if row not in written_rows and _fuzzy_match_name(acct["name"], tmpl_name):
                    best_row = row; break
            if best_row:
                ws_det.cell(best_row, 4).value = round(abs(acct["net"]), 2)
                injected.append(f"Details!D{best_row} ({acct['name']}) = {abs(acct['net']):,.2f}")
                written_rows.add(best_row)
            else:
                unmatched_creditors.append(acct)

        # Blank rows for unmatched (from cache)
        blank_rows = [r for r in range(21, 63)
                      if r not in written_rows
                      and det_cache.get((r, 4)) is None
                      and det_cache.get((r, 2)) is None]
        for i, acct in enumerate(unmatched_creditors):
            if i < len(blank_rows):
                r = blank_rows[i]
                ws_det.cell(r, 2).value = acct["name"]
                ws_det.cell(r, 4).value = round(abs(acct["net"]), 2)
                injected.append(f"Details!D{r} (new: {acct['name']}) = {abs(acct['net']):,.2f}")
            else:
                skipped.append(f"Trade payable '{acct['name']}': no row in Details")

        # Trade Receivables D74:D90
        receivable_accounts = [a for a in individual_accounts
                               if a.get("bs_head") == "trade_rec"
                               and abs(a.get("net", 0)) > 0]
        recv_written = set()
        for acct in receivable_accounts:
            placed = False
            for r in range(74, 90):
                b = det_cache.get((r, 2))
                if b and _fuzzy_match_name(acct["name"], str(b)) and r not in recv_written:
                    ws_det.cell(r, 4).value = round(abs(acct["net"]), 2)
                    injected.append(f"Details!D{r} ({acct['name']}) = {abs(acct['net']):,.2f}")
                    recv_written.add(r); placed = True; break
            if not placed:
                for r in range(74, 90):
                    if det_cache.get((r, 4)) is None and det_cache.get((r, 2)) is None and r not in recv_written:
                        ws_det.cell(r, 2).value = acct["name"]
                        ws_det.cell(r, 4).value = round(abs(acct["net"]), 2)
                        injected.append(f"Details!D{r} (recv: {acct['name']}) = {abs(acct['net']):,.2f}")
                        recv_written.add(r); break

        # LT borrowings D7:D10
        ltb_accounts = [a for a in individual_accounts
                        if a.get("bs_head") == "lt_borrowings"
                        and abs(a.get("net", 0)) > 0]
        for acct in ltb_accounts:
            placed = False
            for r in range(7, 12):
                b = det_cache.get((r, 2))
                if b and _fuzzy_match_name(acct["name"], str(b)):
                    ws_det.cell(r, 4).value = round(abs(acct["net"]), 2)
                    injected.append(f"Details!D{r} (loan: {acct['name']}) = {abs(acct['net']):,.2f}")
                    placed = True; break
            if not placed:
                for r in range(7, 12):
                    if det_cache.get((r, 4)) is None:
                        ws_det.cell(r, 2).value = acct["name"]
                        ws_det.cell(r, 4).value = round(abs(acct["net"]), 2)
                        injected.append(f"Details!D{r} (loan: {acct['name']}) = {abs(acct['net']):,.2f}")
                        break

    # ────────────────────────────────────────────────────────────────
    # 4. CLOSING STOCK / INVENTORIES
    # ────────────────────────────────────────────────────────────────
    # notes to p&l!D24 is a formula (='GROSS PROFIT'!E17) which itself
    # is computed from purchases. For TB injection we must override it
    # directly with the TB closing stock value.
    inventory_amt = aggregated_values.get("inventories", 0)
    if inventory_amt:
        placed = False
        if "notes to p&l" in wb.sheetnames:
            ws_npl = wb["notes to p&l"]
            # Override D24 directly (even if formula) — closing stock must come from TB
            ws_npl.cell(24, 4).value = round(inventory_amt, 2)
            injected.append(f"notes to p&l!D24 (Closing stock override) = {inventory_amt:,.2f}")
            placed = True
        if not placed and "GROSS PROFIT" in wb.sheetnames:
            ws_gp = wb["GROSS PROFIT"]
            ws_gp.cell(17, 5).value = round(inventory_amt, 2)
            injected.append(f"GROSS PROFIT!E17 (Closing stock override) = {inventory_amt:,.2f}")

    # ────────────────────────────────────────────────────────────────
    # 5. FIXED ASSETS — write WDV from TB into Fixed Assets C. Yr.
    # ────────────────────────────────────────────────────────────────
    fixed_assets_amt = aggregated_values.get("fixed_assets", 0)
    if fixed_assets_amt and "Fixed Assets C. Yr." in wb.sheetnames:
        ws_fa = wb["Fixed Assets C. Yr."]
        if individual_accounts:
            fa_accounts = [a for a in individual_accounts
                           if a.get("bs_head") == "fixed_assets"
                           and abs(a.get("net", 0)) > 0]
            # Match by name to FA rows
            for acct in fa_accounts:
                for r in range(10, 35):
                    a_val = ws_fa.cell(r, 1).value
                    if a_val and _fuzzy_match_name(acct["name"], str(a_val)):
                        # Col C = additions > 180 days
                        c_val = ws_fa.cell(r, 3).value
                        if c_val is not None and not _is_formula(str(c_val)):
                            # TB net = WDV after depreciation
                            # Write to col C (additions > 180 days) if opening (col B) is zero
                            b_val = ws_fa.cell(r, 2).value
                            opening_fa = float(b_val) if b_val and not _is_formula(str(b_val)) else 0
                            if opening_fa == 0:
                                if _safe_set(ws_fa, r, 3, abs(acct["net"])):
                                    injected.append(f"FA C. Yr.!C{r} ({acct['name']}) = {abs(acct['net']):,.2f}")
                        break
        log.append(f"Fixed assets total from TB: {fixed_assets_amt:,.2f} (written to individual FA rows)")

    # ────────────────────────────────────────────────────────────────
    # 6. SHORT TERM PROVISIONS
    # ────────────────────────────────────────────────────────────────
    stp_amt = aggregated_values.get("st_provisions", 0)
    if stp_amt and individual_accounts:
        prov_accounts = [a for a in individual_accounts
                         if a.get("bs_head") == "st_provisions"
                         and abs(a.get("net", 0)) > 0]
        if "notes to bs" in wb.sheetnames:
            ws_n = wb["notes to bs"]
            # TCS GST A/C → row 73 area  
            for acct in prov_accounts:
                name_l = acct["name"].lower()
                if "tcs" in name_l and _safe_set(ws_n, 73, 4, abs(acct["net"])):
                    injected.append(f"notes to bs!D73 (TCS provision) = {abs(acct['net']):,.2f}")

    # ────────────────────────────────────────────────────────────────
    # 7. P&L NOTES INJECTION (notes to p&l + GROSS PROFIT)
    # ────────────────────────────────────────────────────────────────
    # P&L sheet cells are ALL formulas pulling from notes to p&l and
    # GROSS PROFIT. We must write to the source sheets, never p&l directly.
    #
    # Chain:
    #  p&l!E7  (Revenue)      ← notes to p&l!D7  ← notes to p&l!D6
    #                                                ← SUM(GROSS PROFIT!E11:E14)
    #  p&l!E12 (Cost Mat.)    ← notes to p&l!D26 ← D17+D20+D22-D24
    #                              D17=E24 (formula←prev yr closing)
    #                              D20=SUM(GROSS PROFIT!B14:B18)
    #                              D24='GROSS PROFIT'!E17 (formula=auto)
    #  p&l!E13 (Employee)     ← notes to p&l!D34 ← SUM(D31:D33)  → write D31
    #  p&l!E14 (Finance cost) ← notes to p&l!D40 ← SUM(D38:D39)  → write D38
    #  p&l!E15 (Depreciation) ← notes to p&l!D53 ← D52
    #                              D52='Fixed Assets C. Yr.'!H31 (formula=auto)
    #  p&l!E16 (Other exp)    ← notes to p&l!D79 ← SUM(D57:D78)  → write D57:D78

    if individual_accounts and "notes to p&l" in wb.sheetnames:
        ws_npl = wb["notes to p&l"]

        # ── A. Sales → GROSS PROFIT!E11:E14 ─────────────────────────
        # GROSS PROFIT!E11 = Sale GST 12% Interstate
        # GROSS PROFIT!E12 = Sale GST 12% Within State
        # GROSS PROFIT!E13 = Sale GST 5% Interstate
        # GROSS PROFIT!E14 = Sale GST 5% Within State
        if "GROSS PROFIT" in wb.sheetnames:
            ws_gp = wb["GROSS PROFIT"]

            sale_accounts = [a for a in individual_accounts
                             if a.get("bs_head") == "revenue"
                             and abs(a.get("net", 0)) > 0]

            # Map TB sale account names → GROSS PROFIT rows
            # Row 11=12% Interstate, 12=12% WS, 13=5% Interstate, 14=5% WS
            SALE_ROW_MAP = {
                "12% interstate": 11, "12% intertate": 11,
                "12% within": 12,
                "5% interstate": 13, "5% intertate": 13,
                "5% within": 14,
            }
            sale_row_totals = {11: 0, 12: 0, 13: 0, 14: 0}
            unmatched_sale = 0

            for acct in sale_accounts:
                name_l = acct["name"].lower()
                matched = False
                for pattern, row in SALE_ROW_MAP.items():
                    if pattern in name_l:
                        sale_row_totals[row] += abs(acct["net"])
                        matched = True
                        break
                if not matched:
                    # Dump into row 14 (5% within state = largest category)
                    sale_row_totals[14] += abs(acct["net"])

            for row, amt in sale_row_totals.items():
                if amt > 0:
                    e_val = ws_gp.cell(row, 5).value
                    if not _is_formula(str(e_val or "")):
                        ws_gp.cell(row, 5).value = round(amt, 2)
                        injected.append(f"GROSS PROFIT!E{row} (Sale) = {amt:,.2f}")
                    else:
                        # Formula cell — override it
                        ws_gp.cell(row, 5).value = round(amt, 2)
                        injected.append(f"GROSS PROFIT!E{row} (Sale override) = {amt:,.2f}")

        # ── B. Purchases → GROSS PROFIT!B14:B18 ─────────────────────
        # Row 14=Purchase GST 12% Interstate, 15=12% WS, 16=18% WS
        # Row 17=5% Interstate, 18=5% WS
        if "GROSS PROFIT" in wb.sheetnames:
            purchase_accounts = [a for a in individual_accounts
                                  if a.get("bs_head") == "purchases"
                                  and abs(a.get("net", 0)) > 0]

            PURCH_ROW_MAP = {
                "12% interstate": 14, "12% intertate": 14,
                "12% within": 15,
                "18% within": 16,
                "5% interstate": 17, "5% intertate": 17,
                "5% within": 18,
            }
            purch_row_totals = {14: 0, 15: 0, 16: 0, 17: 0, 18: 0}

            for acct in purchase_accounts:
                name_l = acct["name"].lower()
                matched = False
                for pattern, row in PURCH_ROW_MAP.items():
                    if pattern in name_l:
                        purch_row_totals[row] += abs(acct["net"])
                        matched = True
                        break
                if not matched:
                    purch_row_totals[18] += abs(acct["net"])

            for row, amt in purch_row_totals.items():
                if amt > 0:
                    b_val = ws_gp.cell(row, 2).value
                    if not _is_formula(str(b_val or "")):
                        ws_gp.cell(row, 2).value = round(amt, 2)
                    else:
                        ws_gp.cell(row, 2).value = round(amt, 2)
                    injected.append(f"GROSS PROFIT!B{row} (Purchase) = {amt:,.2f}")

        # ── C. Opening Stock → GROSS PROFIT!B9 ──────────────────────
        # B9 = "='notes to p&l'!D17" which = "=E24" (prev yr closing)
        # Opening stock should come from TB opening stock account
        opening_stock_accounts = [a for a in individual_accounts
                                   if "opening stock" in a.get("name","").lower()
                                   and abs(a.get("net", 0)) > 0]
        if opening_stock_accounts and "GROSS PROFIT" in wb.sheetnames:
            total_opening = sum(abs(a["net"]) for a in opening_stock_accounts)
            # notes to p&l!D17 is "=E24" — override it directly
            ws_npl.cell(17, 4).value = round(total_opening, 2)
            injected.append(f"notes to p&l!D17 (Opening stock) = {total_opening:,.2f}")

        # ── D. Employee Expenses → notes to p&l!D31 ─────────────────
        # D34 = SUM(D31:D33), p&l!E13 ← notes to p&l!D34
        # Include accounts classified as employee_expenses OR named "salary"
        salary_accounts = [a for a in individual_accounts
                           if (a.get("bs_head") == "employee_expenses"
                               or "salary" in a.get("name","").lower()
                               or "wage" in a.get("name","").lower())
                           and "payable" not in a.get("name","").lower()
                           and abs(a.get("net", 0)) > 0]
        if salary_accounts:
            total_salary = sum(abs(a["net"]) for a in salary_accounts)
            ws_npl.cell(31, 4).value = round(total_salary, 2)
            injected.append(f"notes to p&l!D31 (Salaries) = {total_salary:,.2f}")

        # ── E. Finance Cost → notes to p&l!D38 ──────────────────────
        # D40 = SUM(D38:D39), finance cost rows
        finance_accounts = [a for a in individual_accounts
                            if a.get("bs_head") in ("other_expenses",)
                            and any(kw in a.get("name","").lower()
                                    for kw in ["bank interest","bank cc intt","intt paid",
                                               "interest on loan","interest paid","finance"])
                            and abs(a.get("net", 0)) > 0]
        if finance_accounts:
            total_finance = sum(abs(a["net"]) for a in finance_accounts)
            d38 = ws_npl.cell(38, 4).value
            if not _is_formula(str(d38 or "")):
                ws_npl.cell(38, 4).value = round(total_finance, 2)
                injected.append(f"notes to p&l!D38 (Finance cost) = {total_finance:,.2f}")

        # ── F. Other Expenses → notes to p&l!D57:D78 ─────────────────
        # Each row in D57:D78 corresponds to a specific expense item.
        # Read the template labels (col B) and match TB accounts.
        # D79 = SUM(D57:D78), p&l!E16 ← notes to p&l!D79

        # Finance-cost keywords to exclude from other_expenses
        FINANCE_KEYWORDS = {"bank interest","bank cc intt","intt paid",
                            "interest on loan","interest paid"}

        # Also exclude salary and depreciation — handled separately
        EXCLUDE_FROM_OTHER = FINANCE_KEYWORDS | {"salary","depreciation","dep on","amort"}

        other_exp_accounts = [a for a in individual_accounts
                              if a.get("bs_head") == "other_expenses"
                              and abs(a.get("net", 0)) > 0
                              and not any(kw in a.get("name","").lower()
                                          for kw in EXCLUDE_FROM_OTHER)
                              and "salary" not in a.get("name","").lower()
                              and "depreciation" not in a.get("name","").lower()]

        # Build template label → row map for D57:D78
        # Key = normalized label from col B, value = row number
        exp_row_map = {}
        npl_cache = _cache.get("notes to p&l", {})
        for r in range(57, 79):
            b_val = npl_cache.get((r, 2))
            if b_val:
                exp_row_map[r] = str(b_val).strip().lower()

        # For each TB account, find the best matching template row
        # Strategy: for each template row label, find the TB account whose
        # name best matches — not the other way around.
        # This prevents mismatches from greedy matching.

        # Build reverse: for each TB account → best matching row
        written_exp_rows = set()
        account_row_assignments = {}  # acct_key → row

        for acct in other_exp_accounts:
            name_l = acct["name"].lower().strip()
            best_row = None
            best_score = 0

            for r, lbl in exp_row_map.items():
                if r in written_exp_rows:
                    continue
                score = 0
                # Score based on common significant words
                name_words = set(w for w in name_l.replace("."," ").split() if len(w) > 2)
                lbl_words  = set(w for w in lbl.replace("."," ").split() if len(w) > 2)
                common = name_words & lbl_words
                if common:
                    score = sum(len(w) for w in common)
                # Bonus for exact substring
                if lbl in name_l or name_l in lbl:
                    score += 20
                # Key abbreviation matches
                abbrev_map = {
                    "exp": "expenses",
                    "exp.": "expenses",
                    "adda": "",
                    "advertisment": "advertisement",
                    "advertisement": "advertisement",
                    "stationery": "stationary",
                    "stationary": "stationary",
                }
                for short, full in abbrev_map.items():
                    if short in name_l and (short in lbl or full in lbl):
                        score += 5

                if score > best_score:
                    best_score = score
                    best_row = r

            if best_row and best_score > 2:
                account_row_assignments[acct["name"]] = best_row
                written_exp_rows.add(best_row)

        # Write matched accounts
        for acct in other_exp_accounts:
            amt = abs(acct["net"])
            row = account_row_assignments.get(acct["name"])
            if row:
                d_val = ws_npl.cell(row, 4).value
                if not _is_formula(str(d_val or "")):
                    ws_npl.cell(row, 4).value = round(amt, 2)
                    injected.append(f"notes to p&l!D{row} ({acct['name']}) = {amt:,.2f}")
            else:
                # No template match — find first empty D row in 57:78
                placed = False
                for r in range(57, 79):
                    if r not in written_exp_rows:
                        d_val = _cache.get("notes to p&l", {}).get((r, 4))
                        if d_val is None or d_val == 0:
                            ws_npl.cell(r, 4).value = round(amt, 2)
                            written_exp_rows.add(r)
                            injected.append(f"notes to p&l!D{r} (unmatched: {acct['name']}) = {amt:,.2f}")
                            placed = True
                            break
                if not placed:
                    skipped.append(f"Other expense '{acct['name']}' = {amt:,.2f}: no row in notes to p&l")

    # ── Depreciation note (D52) comes from Fixed Assets C. Yr.!H31 ──
    # H31 = SUM(H10:H30) which is computed from individual asset rows.
    # If FA rows are properly filled (section 5 above), this auto-calculates.
    # For TB-based filing, if depreciation is explicitly in TB:
    dep_accounts = []
    if individual_accounts:
        dep_accounts = [a for a in individual_accounts
                        if a.get("bs_head") == "depreciation"
                        and abs(a.get("net", 0)) > 0]
    if dep_accounts and "Fixed Assets C. Yr." in wb.sheetnames:
        total_dep = sum(abs(a["net"]) for a in dep_accounts)
        ws_fa = wb["Fixed Assets C. Yr."]
        # H31 = SUM formula — override only if sum doesn't match TB
        h31 = ws_fa.cell(31, 8).value
        if _is_formula(str(h31 or "")):
            # H31 is a formula (=SUM(H10:H30)).
            # Write TB depreciation total to H10 as a lump if H10:H30 are all zero
            fa_dep_sum = sum(
                float(ws_fa.cell(r, 8).value or 0)
                for r in range(10, 31)
                if not _is_formula(str(ws_fa.cell(r, 8).value or ""))
            )
            if abs(fa_dep_sum - total_dep) > 1:
                # Find first non-formula H row and write the difference
                for r in range(10, 31):
                    h_val = ws_fa.cell(r, 8).value
                    if h_val is not None and not _is_formula(str(h_val)):
                        ws_fa.cell(r, 8).value = float(h_val) + (total_dep - fa_dep_sum)
                        injected.append(f"Fixed Assets C. Yr.!H{r} (dep adjustment) = {float(h_val) + (total_dep - fa_dep_sum):,.2f}")
                        break
    for s in skipped:
        log.append(f"⚠ SKIPPED: {s}")
    for inj_msg in injected:
        log.append(f"✓ {inj_msg}")

    wb.save(output_path)
    wb.close()

    total_assets = sum(
        aggregated_values.get(k, 0)
        for k in aggregated_values
        if BS_HEADS.get(k, {}).get("side") == "asset"
    )
    total_liabilities = sum(
        aggregated_values.get(k, 0)
        for k in aggregated_values
        if BS_HEADS.get(k, {}).get("side") == "liability"
    )
    diff = abs(total_assets - total_liabilities)

    log.append(f"\n📊 Tally Check (TB Aggregates):")
    log.append(f"  Total Assets:      {total_assets:>15,.2f}")
    log.append(f"  Total Liabilities: {total_liabilities:>15,.2f}")
    if diff < 1:
        log.append(f"  ✅ Balance Sheet TALLIES!")
    else:
        log.append(f"  ❌ Difference: {diff:,.2f}")
        if total_assets > total_liabilities:
            log.append(f"  Assets exceed Liabilities by {diff:,.2f}")
        else:
            log.append(f"  Liabilities exceed Assets by {diff:,.2f}")

    return {
        "status": "success",
        "output": output_path,
        "log": log,
        "injected_count": len(injected),
        "skipped_count": len(skipped),
        "tally_ok": diff < 1,
        "total_assets": total_assets,
        "total_liabilities": total_liabilities,
    }


def detect_bs_template(file_path):
    """
    Detect BS template structure — returns info for the UI.
    Primary sheet is always the bs/balance sheet sheet.
    """
    wb = load_workbook(file_path, read_only=True, data_only=True)
    result = {"sheets": []}

    for sname in wb.sheetnames:
        ws = wb[sname]
        rows_data = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            rows_data.append(list(row))
            if i >= 100:
                break
        sheet_info = _detect_bs_columns(rows_data, sname)
        if sheet_info:
            result["sheets"].append(sheet_info)

    wb.close()

    if not result["sheets"]:
        return {"error": "Could not detect Balance Sheet template structure."}

    best = max(result["sheets"], key=lambda s: len(s.get("head_rows", {})))
    result["primary"] = best
    return result



# ═══════════════════════════════════════════════════════════════════════
# MAIN PIPELINE (called from app.py routes)
# ═══════════════════════════════════════════════════════════════════════

def analyze_trial_balance(tb_path):
    """
    Step 1: Read and analyze the trial balance.
    Returns structure info + classified accounts.
    """
    detection = detect_tb_structure(tb_path)
    if "error" in detection:
        return detection

    classified = classify_accounts(detection["accounts"])

    # Separate by confidence
    high_conf = [a for a in classified if a["confidence"] == "high"]
    low_conf = [a for a in classified if a["confidence"] == "low"]
    unclassified = [a for a in classified if a["confidence"] == "none"]

    return {
        "status": "success",
        "detection": {
            "format_type": detection["format_type"],
            "sheet_name": detection["sheet_name"],
            "header_row": detection["header_row"],
            "data_start_row": detection["data_start_row"],
            "account_col": detection["account_col"],
            "debit_col": detection["debit_col"],
            "credit_col": detection["credit_col"],
            "net_col": detection["net_col"],
        },
        "accounts": classified,
        "summary": {
            "total_accounts": len(classified),
            "high_confidence": len(high_conf),
            "low_confidence": len(low_conf),
            "unclassified": len(unclassified),
        },
    }


def process_tb_to_bs(tb_path, bs_template_path, output_path, user_mapping=None):
    """
    Full pipeline: analyze TB → classify → inject into BS.

    user_mapping: dict of overrides. Supports TWO formats:
      1. {"ACCOUNT NAME": "bs_head"}   — matched by name (case-insensitive)
      2. {"ACCOUNT NAME_rownum": "bs_head"} — matched by unique key

    This is the SINGLE SOURCE OF TRUTH for user overrides.
    The Flask /tb-process route must pass user_mapping here.
    """
    # Step 1: Analyze TB
    analysis = analyze_trial_balance(tb_path)
    if "error" in analysis:
        return analysis

    accounts = analysis["accounts"]

    # Step 2: Apply user overrides — match by unique key first, then by name
    if user_mapping:
        name_map = {}
        key_map  = {}
        for raw_key, head in user_mapping.items():
            if not head or head in ("auto", "", None):
                continue
            rk = str(raw_key).strip()
            # Key format: "ACCOUNT NAME_123"  (ends with underscore + digits)
            parts = rk.rsplit("_", 1)
            if len(parts) == 2 and parts[1].isdigit():
                key_map[rk] = head
            else:
                name_map[rk.upper()] = head

        for acct in accounts:
            acct_key  = str(acct.get("key", "")).strip()
            acct_name = str(acct.get("name", "")).strip().upper()
            if acct_key in key_map:
                acct["bs_head"]    = key_map[acct_key]
                acct["confidence"] = "user"
            elif acct_name in name_map:
                acct["bs_head"]    = name_map[acct_name]
                acct["confidence"] = "user"

    # Step 3: Aggregate by bs_head (uses overridden heads)
    aggregated = get_aggregated_values(accounts)

    # Step 4: Inject into BS template
    result = inject_into_bs(
        bs_template_path, output_path, aggregated,
        mapping_overrides=None,
        individual_accounts=accounts,
    )

    result["analysis"]            = analysis
    result["aggregated"]          = aggregated
    result["classified_accounts"] = accounts   # full list with final bs_head applied
    return result
