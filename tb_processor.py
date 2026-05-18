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
    "long_term_borrowings": {
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
    "short_term_borrowings": {
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
    "other_current_liabilities": {
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
    "short_term_provisions": {
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
    "trade_receivables": {
        "label": "Trade Receivables",
        "side": "asset",
        "keywords": [
            "trade receivable", "debtor", "sundry debtor",
            "accounts receivable", "bills receivable",
            "trade debtor", "receivable from customer",
        ],
        "negative_keywords": [],
    },
    "cash_and_bank": {
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
    "short_term_loans_advances": {
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
    "trade_payables", "trade_receivables",
    "cash_and_bank", "inventories",
    "short_term_provisions",
    "short_term_borrowings", "long_term_borrowings",
    "employee_expenses", "purchases", "revenue",
    "fixed_assets", "non_current_investments",
    "short_term_loans_advances",
    "capital",
    "other_current_liabilities", "other_current_assets",
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
            if i >= 60:
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

            # Debit column
            if val in ("dr", "debit", "dr balance", "debit balance",
                       "dr amount", "debit amount", "dr bal", "debit bal"):
                dr_col = ci

            # Credit column
            if val in ("cr", "credit", "cr balance", "credit balance",
                       "cr amount", "credit amount", "cr bal", "credit bal"):
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
    if dr_col is not None and cr_col is not None:
        format_type = 1  # Dr/Cr separate columns
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

    # Data starts after header
    data_start = header_row + 1

    # Extract accounts
    accounts = []
    total_keywords = {"total", "grand total", "difference", "net total",
                      "closing balance", "opening balance total",
                      "balance c/d", "balance b/d"}

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
            net_val = row[net_col] if net_col < len(row) else None
            net_amt = _to_float(net_val)
            if net_amt > 0:
                dr_amt = net_amt
            else:
                cr_amt = abs(net_amt)

        if dr_amt == 0 and cr_amt == 0 and net_amt == 0:
            continue

        accounts.append({
            "row": ri,
            "name": acct_name,
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

def classify_accounts(accounts):
    """
    Classify each TB account under a BS/P&L head.
    Returns list of dicts with added 'bs_head' and 'confidence' keys.
    """
    classified = []
    for acct in accounts:
        name = acct["name"]
        bs_head, confidence = _classify_single(name, acct["net"])
        acct_copy = dict(acct)
        acct_copy["bs_head"] = bs_head
        acct_copy["confidence"] = confidence
        classified.append(acct_copy)
    return classified


def _classify_single(name, net_amount):
    """Classify a single account name. Returns (head_key, confidence)."""
    name_lower = name.lower().strip()

    # Try each head in priority order
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

    # If no match: use net amount sign as hint
    # Debit balance (positive net) → likely asset or expense
    # Credit balance (negative net) → likely liability or income
    if net_amount > 0:
        return "other_current_assets", "low"
    elif net_amount < 0:
        return "other_current_liabilities", "low"
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
# BS TEMPLATE INJECTION
# ═══════════════════════════════════════════════════════════════════════

# Mapping from BS head keys to common labels found in BS templates
BS_LABEL_MAP = {
    "capital": [
        "capital", "partner", "proprietor", "owner", "equity",
        "reserves", "surplus", "shareholders fund",
    ],
    "long_term_borrowings": [
        "long term borrowing", "long-term borrowing", "secured loan",
        "term loan",
    ],
    "short_term_borrowings": [
        "short term borrowing", "short-term borrowing",
        "overdraft", "cash credit", "bank borrowing",
    ],
    "trade_payables": [
        "trade payable", "creditor", "sundry creditor",
    ],
    "other_current_liabilities": [
        "other current liabilit", "statutory",
    ],
    "short_term_provisions": [
        "short term provision", "provision",
    ],
    "fixed_assets": [
        "fixed asset", "property", "plant", "ppe",
        "tangible", "intangible",
    ],
    "non_current_investments": [
        "non-current investment", "non current investment",
        "investment", "long term investment",
    ],
    "inventories": [
        "inventor", "stock", "closing stock",
    ],
    "trade_receivables": [
        "trade receivable", "debtor", "sundry debtor",
    ],
    "cash_and_bank": [
        "cash and bank", "cash & bank", "cash at bank",
        "bank balance",
    ],
    "short_term_loans_advances": [
        "short term loan", "loan and advance", "loans & advance",
        "advance",
    ],
    "other_current_assets": [
        "other current asset",
    ],
    "revenue": [
        "revenue", "sale", "income from operation",
    ],
    "purchases": [
        "purchase", "cost of material", "cost of goods",
    ],
    "employee_expenses": [
        "employee", "salary", "staff",
    ],
    "other_expenses": [
        "other expense", "indirect expense", "administrative",
    ],
    "depreciation": [
        "depreciation", "amortisation",
    ],
}


def detect_bs_template(file_path):
    """
    Scan the BS template to find:
    - Which column is CY (current year)
    - Which column is PY (previous year)
    - Where each BS head label appears (row mapping)
    Returns dict with cy_col, py_col, head_rows, sheet_name
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

    # Return the best sheet (most head matches)
    best = max(result["sheets"], key=lambda s: len(s.get("head_rows", {})))
    result["primary"] = best
    return result


def _detect_bs_columns(rows, sheet_name):
    """Detect CY/PY columns and head label rows in a BS template."""
    cy_col = None
    py_col = None
    head_rows = {}

    # Scan for year headers in first 15 rows
    current_year_patterns = [
        r"31[\.\s/-]03[\.\s/-]20\d{2}",
        r"31st?\s+march[,\s]+20\d{2}",
        r"current\s+year",
        r"as\s+at.*20\d{2}",
    ]

    for ri, row in enumerate(rows[:15]):
        if not row:
            continue
        for ci, val in enumerate(row):
            if not val or not isinstance(val, str):
                continue
            val_lower = val.strip().lower()

            for pat in current_year_patterns:
                if re.search(pat, val_lower):
                    # Determine if this is CY or PY based on position
                    # CY is usually the first/left date column
                    year_match = re.search(r"20(\d{2})", val)
                    if year_match:
                        yr = int("20" + year_match.group(1))
                        if cy_col is None:
                            cy_col = ci
                        elif ci != cy_col and py_col is None:
                            py_col = ci
                            # The later year is CY, earlier is PY
                            cy_match = re.search(r"20(\d{2})", str(rows[ri][cy_col]) if rows[ri][cy_col] else "")
                            py_match = re.search(r"20(\d{2})", str(rows[ri][py_col]) if rows[ri][py_col] else "")
                            if cy_match and py_match:
                                cy_yr = int("20" + cy_match.group(1))
                                py_yr = int("20" + py_match.group(1))
                                if py_yr > cy_yr:
                                    cy_col, py_col = py_col, cy_col

    if cy_col is None:
        return None

    # Scan for BS head labels
    for ri, row in enumerate(rows):
        if not row:
            continue
        for ci, val in enumerate(row):
            if not val or not isinstance(val, str):
                continue
            val_lower = val.strip().lower()

            for head_key, labels in BS_LABEL_MAP.items():
                for label in labels:
                    if label in val_lower:
                        if head_key not in head_rows:
                            head_rows[head_key] = {
                                "row": ri,
                                "col": ci,
                                "label_found": val.strip(),
                            }
                        break

    # Also look for Note numbers (common in Indian BS templates)
    note_rows = {}
    for ri, row in enumerate(rows):
        if not row:
            continue
        for ci, val in enumerate(row):
            if val is None:
                continue
            val_str = str(val).strip()
            # Note numbers like "1", "2", "2.1" etc
            if re.match(r'^\d{1,2}(\.\d{1,2})?$', val_str):
                note_rows[val_str] = {"row": ri, "col": ci}

    return {
        "sheet_name": sheet_name,
        "cy_col": cy_col,
        "py_col": py_col,
        "head_rows": head_rows,
        "note_rows": note_rows,
    }


def inject_into_bs(bs_template_path, output_path, aggregated_values, mapping_overrides=None):
    """
    Inject aggregated TB values into the BS template CY column.
    Never touches PY, formulas, formatting.
    mapping_overrides: dict {head_key: {"row": r, "col": c}} to override auto-detection.
    Returns dict with status and log.
    """
    log = []

    # Detect template structure
    template_info = detect_bs_template(bs_template_path)
    if "error" in template_info:
        return {"status": "error", "message": template_info["error"], "log": log}

    primary = template_info["primary"]
    target_sheet = primary["sheet_name"]
    cy_col = primary["cy_col"]
    head_rows = primary["head_rows"]

    log.append(f"Target sheet: {target_sheet}")
    log.append(f"CY column detected: {_col_letter(cy_col)}")

    # Open workbook for writing (preserves formatting)
    wb = load_workbook(bs_template_path)
    ws = wb[target_sheet]

    injected = []
    skipped = []

    for head_key, amount in aggregated_values.items():
        if amount == 0:
            continue

        # Check for override
        if mapping_overrides and head_key in mapping_overrides:
            target_row = mapping_overrides[head_key]["row"]
            target_col = mapping_overrides[head_key].get("col", cy_col)
        elif head_key in head_rows:
            # Use the row where the label was found
            target_row = head_rows[head_key]["row"]
            # Write to the CY column on the same row
            target_col = cy_col
        else:
            skipped.append(f"{head_key}: {BS_HEADS.get(head_key, {}).get('label', head_key)} = {amount:,.2f} (no target row found)")
            continue

        # Convert 0-indexed to 1-indexed for openpyxl
        cell_row = target_row + 1
        cell_col = target_col + 1

        cell = ws.cell(row=cell_row, column=cell_col)

        # Skip formula cells
        if isinstance(cell.value, str) and cell.value.startswith("="):
            log.append(f"Skipped formula cell at ({cell_row}, {_col_letter(target_col)}): {cell.value}")
            continue

        # Skip merged cells
        if isinstance(cell, MergedCell):
            log.append(f"Skipped merged cell at ({cell_row}, {_col_letter(target_col)})")
            continue

        cell.value = round(amount, 2)
        label = BS_HEADS.get(head_key, {}).get("label", head_key)
        injected.append(f"{label} → ({cell_row}, {_col_letter(target_col)}) = {amount:,.2f}")

    for s in skipped:
        log.append(f"⚠ SKIPPED: {s}")
    for inj in injected:
        log.append(f"✓ {inj}")

    # Save
    wb.save(output_path)
    wb.close()

    # Tally check
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

    log.append(f"\n📊 Tally Check:")
    log.append(f"  Total Assets:      {total_assets:>15,.2f}")
    log.append(f"  Total Liabilities: {total_liabilities:>15,.2f}")
    diff = abs(total_assets - total_liabilities)
    if diff < 1:
        log.append(f"  ✅ Balance Sheet TALLIES!")
    else:
        log.append(f"  ❌ Difference: {diff:,.2f}")
        # Find which side is more
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


def _col_letter(col_idx):
    """Convert 0-indexed column to letter."""
    from openpyxl.utils import get_column_letter
    return get_column_letter(col_idx + 1)


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
    user_mapping: dict {account_name: head_key} for user overrides.
    Returns result dict.
    """
    # Step 1: Analyze TB
    analysis = analyze_trial_balance(tb_path)
    if "error" in analysis:
        return analysis

    accounts = analysis["accounts"]

    # Step 2: Apply user overrides if any
    if user_mapping:
        for acct in accounts:
            if acct["name"] in user_mapping:
                acct["bs_head"] = user_mapping[acct["name"]]
                acct["confidence"] = "user"

    # Step 3: Aggregate
    aggregated = get_aggregated_values(accounts)

    # Step 4: Inject into BS template
    result = inject_into_bs(bs_template_path, output_path, aggregated)

    result["analysis"] = analysis
    result["aggregated"] = aggregated
    return result
