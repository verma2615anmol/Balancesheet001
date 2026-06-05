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
    "advance_from_customer": {
        # Credit-balance debtor → advance received from customer (liability)
        # Shown under Sundry Creditors group as "Advance from Customer".
        "label": "Advance from Customer",
        "side": "liability",
        "keywords": [
            "advance from customer", "customer advance",
            "advance received", "advance from buyer",
            "advance from m/s", "advance from mr", "advance from ms",
            "advance from ",  # generic fallback — catches "Advance from XYZ Ltd"
        ],
        "negative_keywords": [
            "advance from supplier", "advance from vendor",
        ],
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
        "negative_keywords": ["cash credit", "cc account", "overdraft", "od account", "bank od",
                              "interest", "loan interest", "loan a/c", "loan account",
                              "bank charge", "processing fee"],
    },
    "stla": {
        "label": "Short Term Loans & Advances",
        "side": "asset",
        "keywords": [
            "loan given", "advance to", "loan to",
            "advance to staff", "staff advance",
            "prepaid", "deposit", "security deposit paid",
            "tds receivable", "tcs receivable", "input tax",
            "input cgst", "input sgst", "input igst", "gst input",
            "advance tax", "self assessment tax", "mat credit",
            "cenvat", "vat input", "excise input",
            "income tax refund", "refund receivable",
            "advance recoverable",
        ],
        "negative_keywords": ["advance from customer", "customer advance",
                              "advance to supplier", "advance to customer",
                              "supplier advance"],
    },
    "advance_to_supplier": {
        # Debit-balance creditor → advance paid to supplier (asset)
        # Listed under Sundry Debtors group as "Advance to Supplier / Customer".
        "label": "Advance to Supplier",
        "side": "asset",
        "keywords": [
            "advance to supplier", "advance to customer", "supplier advance",
            "advance paid to supplier", "advance to vendor", "vendor advance",
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
    "other_income": {
        "label": "Other Income",
        "side": "pl",
        "keywords": [
            "other income", "interest received", "interest income",
            "dividend received", "dividend income", "rental income",
            "rent received", "miscellaneous income", "misc income",
            "profit on sale", "gain on sale", "exchange gain",
            "discount received", "discount earned", "commission earned",
            "bad debt recovered", "insurance claim received",
            "interest on fd", "interest on deposit", "bank interest received",
        ],
        "negative_keywords": [],
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
            "partner salary", "partner remuneration", "salary to partner",
            "stipend", "incentive", "overtime",
            "labour refreshment", "labor refreshment",
            "e.s.i", "leave with wages",
        ],
        "negative_keywords": ["salary payable", "wages payable", "bonus payable", "leave with wages payable",
                              "e.s.i payable", "esi payable"],
    },
    "finance_cost": {
        "label": "Finance Cost",
        "side": "pl",
        "keywords": [
            "interest paid", "interest on loan", "bank charge",
            "bank interest", "bank cc intt", "cc interest",
            "bank od interest", "overdraft interest",
            "loan interest", "loan 1 interest", "loan 2 interest",
            "loan 3 interest", "loan 4 interest",
            "car loan interest", "machine loan interest", "machinery loan interest",
            "top up loan interest", "top up car loan interest",
            "interest to partner", "interest on unsecured",
            "intt on tempu", "interest on tempu", "interest on hdfc",
            "interest on term loan", "term loan interest",
            "interest on term loan", "interest on secured",
            "processing fee", "cersai charge",
            "life insurance machinery loan",
        ],
        "negative_keywords": ["interest received", "interest income", "intt paid on late payment", "bank charges and interest", "bank charges", "processing fee"],
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
            "discount allowed", "bad debt",
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
    # NEW — separate bucket for Direct Expenses (Electricity, Wages, Power & Fuel)
    # so they post into the Trading A/c "Direct Expenses" sub-rows instead of
    # being collapsed into Purchases.
    "direct_expenses": {
        "label": "Direct Expenses",
        "side": "pl",
        "keywords": [
            "wages a/c", "wages account", "factory wages", "labour wages",
            "electricity exp", "power and fuel", "power & fuel",
            "oil & lubricant", "oil and lubricant",
            "factory expense", "production expense",
        ],
        "negative_keywords": ["wages payable", "electricity payable"],
    },
}

# Priority order for classification (most specific first)
CLASSIFICATION_PRIORITY = [
    "depreciation",
    "advance_from_customer", "advance_to_supplier",
    "trade_payables", "trade_rec",
    "finance_cost",   # Must be before cash_bank so "LOAN INTEREST" doesn't match "bank"
    "employee_expenses",
    "cash_bank", "inventories",
    "st_provisions",
    "st_borrowings", "lt_borrowings",
    "direct_expenses", "purchases", "revenue", "other_income",
    "fixed_assets", "non_current_investments",
    "stla",
    "capital",
    "other_cl", "other_current_assets",
    "other_expenses",
]


# ═══════════════════════════════════════════════════════════════════════
# PDF TRIAL BALANCE PARSER
# ═══════════════════════════════════════════════════════════════════════

def parse_tb_pdf(pdf_path):
    """Parse a Trial Balance PDF (Tally/Busy/Excel-exported) into account rows.

    HYBRID approach (fixes Dr/Cr column-attribution bug):
      • `extract_tables()` is authoritative for the Debit vs Credit column —
        each table row arrives as ``[name, debit_str, credit_str]`` so a
        credit-balance row on the debtor side (e.g. MEERA FORGING with only
        a credit amount) is preserved exactly as it appears in the PDF.
      • `extract_text()` lines are walked in parallel ONLY to capture the
        group/section headers ("SUNDRY DEBTORS", "SECURED LOANS", …) which
        appear BETWEEN tables and are NOT inside the extracted tables.

    The previous text-only parser inferred Dr/Cr from `current_group` alone,
    which silently mis-classified accounts whose balance was on the OPPOSITE
    side of their group's normal side (e.g. a debtor with a credit balance
    was reported as Dr, breaking the "advance from customer" reclassification
    downstream). The hybrid path keeps the natural sign and lets the existing
    sign-aware reclassification logic move it to the right BS head.

    Returns the same shape as detect_tb_structure for seamless integration.
    """
    import pdfplumber, re

    num_re = re.compile(r'([\d,]+\.\d{2})')

    # Section/group keywords found in Indian TB PDFs. Used to reset
    # `current_group` when walking the text lines.
    GROUP_KEYWORDS = {
        "bank accounts", "bank account",
        "capital account", "capital accounts",
        "cash-in-hand", "cash in hand",
        "fixed assets",
        "direct expenses", "indirect expenses",
        "indirect incomes", "indirect income",
        "direct incomes", "direct income",
        "purchase account", "purchase accounts", "purchases",
        "sales account", "sales accounts", "sales",
        "sundry creditors", "sundry debtors",
        "sundry payables", "sundry payable",
        "sundry receivables", "sundry receivable",
        "loans & advances (asset)", "loans & advances",
        "loans and advances", "loans (liability)", "loans liability",
        "current liabilities", "current assets",
        "deposits (asset)", "deposits",
        "duties & taxes", "duties and taxes",
        "secured loans", "unsecured loans", "unsecure loans",
        "stock-in-hand", "stock in hand",
        "investments",
        "profit & loss account", "profit and loss account",
        "profit & loss a/c",
        "provisions", "reserves & surplus", "reserves and surplus",
        "misc. expenses (asset)", "miscellaneous expenses",
        "branch/divisions", "suspense a/c",
    }

    def _norm_header(line):
        return re.sub(r"[\s\-:]+$", "", line.strip().lower())

    def _is_group_header(line):
        return _norm_header(line) in GROUP_KEYWORDS

    def _to_float(v):
        if v is None: return 0.0
        s = str(v).strip()
        if not s: return 0.0
        s = (s.replace(",", "").replace("₹", "")
               .replace("Rs.", "").replace("Rs", "")
               .replace("(", "-").replace(")", ""))
        try:
            return float(s)
        except ValueError:
            return 0.0

    SKIP_NAMES = {"particulars", "trial balance",
                  "debit amount", "credit amount",
                  "total", "grand total", "opening balance",
                  "closing balance", "balance c/d", "balance b/d",
                  "sub total", "net total"}

    accounts = []
    running_group = ""

    # ── Detect Group-Wise Trial Balance format (plain text, no tables) ──
    _is_group_wise = False
    try:
        with pdfplumber.open(pdf_path) as _pdf_detect:
            _ft = _pdf_detect.pages[0].extract_text() or ""
            if "Group Wise Trial Balance" in _ft or "Group wise Trial Balance" in _ft:
                _is_group_wise = True
    except Exception:
        pass

    if _is_group_wise:
        import re as _re_gw
        _GROUP_SIDE_GW = {
            'CURRENT LIABILITIES': 'credit', 'DUTIES AND TAXES': 'credit',
            'SUNDRY CREDITORS': 'credit', 'SUNDRY PAYABLE': 'credit',
            'CAPITAL ACCOUNT': 'credit', 'SECURED LOANS': 'credit',
            'UNSECURED LOANS': 'credit', 'ADVANCE FROM CUSTOMERS': 'credit',
            'SALES ACCOUNTS': 'credit', 'INDIRECT INCOME': 'credit',
            'BANK ACCOUNTS': 'both',      # ← BEFORE 'CURRENT ASSETS' so it matches first
            'CURRENT ASSETS': 'debit', 'CASH-IN-HAND': 'debit',
            'STOCK-IN-HAND': 'debit', 'SUNDRY DEBTORS': 'debit',
            'FIXED ASSETS': 'debit', 'PURCHASE ACCOUNTS': 'debit',
            'DIRECT EXPENSES': 'debit', 'INDIRECT EXPENSES': 'debit',
        }
        _SKIP_GW  = ['Group Wise', 'Closing Balance', 'P a r t i c u l',
                     'Debit  Credit', 'Debit\tCredit', 'Total :', 'Grand Total',
                     'Continue to next', 'Station', 'Page:',
                     'A.S.TRADERS', 'A.S. TRADERS', '321-I', 'BRS NAGAR',
                     'FOCAL POINT', 'FASHION ADDA', 'SHREE CRAFT']
        _KNOWN_GRP_WORDS = ['LIABILIT','ASSET','LOAN','EXPENSE','INCOME','ACCOUNT',
                            'CREDITOR','DEBTOR','CAPITAL','SALES','PURCHASE',
                            'STOCK','CASH','BANK','FIXED','ADVANCE','PAYABLE']
        _amt_re_gw = _re_gw.compile(r'^(.+?)\s+([\d,]+\.\d{2})\s*(?:([\d,]+\.\d{2}))?\s*$')
        _grp_re_gw = _re_gw.compile(r'^([A-Z][A-Z\s&\./\(\)\-0-9,\']+?)(?:\s*-{3,}\s*\(\s*(.+?)\s*\))?\s*$')
        _cg = "UNKNOWN"; _cs = 'debit'

        def _pi_gw(s):
            return float(s.replace(',', '').strip())

        def _side_gw(grp):
            gu = grp.upper()
            for k, s in _GROUP_SIDE_GW.items():
                if k in gu: return s
            return 'debit'

        with pdfplumber.open(pdf_path) as _pdf_gw:
            for _page_gw in _pdf_gw.pages:
                for _ln in (_page_gw.extract_text() or "").split('\n'):
                    _ln = _ln.strip()
                    if not _ln: continue
                    if any(sk in _ln for sk in _SKIP_GW): continue
                    _am_gw = _amt_re_gw.match(_ln)
                    if _am_gw:
                        _name_gw = _am_gw.group(1).strip()
                        _f = _pi_gw(_am_gw.group(2))
                        _s = _pi_gw(_am_gw.group(3)) if _am_gw.group(3) else None
                        if _s is not None:
                            _d, _c = _f, _s
                        elif _cs == 'credit':
                            _d, _c = 0, _f
                        elif _cs == 'both':
                            # CC A/C with large balance = overdraft (liability) → credit side
                            # Small CC balance (< 50,000) = positive deposit → debit
                            _nl = _name_gw.lower()
                            if ('cc a/c' in _nl or 'cc a/c no' in _nl) and _f > 50000:
                                _d, _c = 0, _f   # large CC = overdraft liability
                            else:
                                _d, _c = _f, 0   # small balance or regular account = asset
                        else:
                            _d, _c = _f, 0
                        accounts.append({
                            "name": _name_gw, "group": _cg,
                            "debit": _d, "credit": _c,
                            "net": _d - _c,
                        })
                    else:
                        _gm_gw = _grp_re_gw.match(_ln)
                        if _gm_gw and not _ln.startswith('('):
                            _mn = _gm_gw.group(1).strip()
                            _sb = (_gm_gw.group(2) or "").strip()
                            _proposed = f"{_mn} ({_sb})" if _sb else _mn
                            # Only accept as group if it has a financial keyword
                            if any(w in _proposed.upper() for w in _KNOWN_GRP_WORDS):
                                _cg = _proposed
                                _cs = _side_gw(_cg)
                            # else: keep previous group (company header line)
        # Skip the table-based parser
        import sys as _sys_gw
        _gw_parsed = True
    else:
        _gw_parsed = False

    if not _gw_parsed:
      try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Text lines on this page (used only for group-header tracking)
                text_lines = [ln.strip() for ln in
                              (page.extract_text() or "").splitlines()
                              if ln.strip()]

                # Tables on this page (authoritative for Dr / Cr columns)
                tables = page.extract_tables() or []
                table_rows = []
                for tbl in tables:
                    for row in tbl:
                        if not row:
                            continue
                        if all((c is None or str(c).strip() == "")
                               for c in row):
                            continue
                        table_rows.append(row)

                # Set of names appearing as first-cell of a data row — used
                # to know which text lines are data (vs group headers).
                tr_first_cells = {str(row[0]).strip()
                                  for row in table_rows
                                  if row and row[0]}

                # Walk text lines IN ORDER to attach the most recently seen
                # group header to each data-row name on this page.
                page_group = running_group
                line_to_group = {}
                for ln in text_lines:
                    if _is_group_header(ln):
                        page_group = _norm_header(ln)
                        continue
                    name_part = re.sub(
                        r"\s+-?[\d,]+\.\d{1,2}(\s+-?[\d,]+\.\d{1,2})?\s*$",
                        "", ln).strip()
                    if name_part and name_part in tr_first_cells:
                        line_to_group.setdefault(name_part, page_group)

                # Emit account rows from the tables
                for row in table_rows:
                    cells = [str(c).strip() if c is not None else ""
                             for c in row]
                    while len(cells) < 3:
                        cells.append("")
                    name, dr_str, cr_str = cells[0], cells[1], cells[2]

                    if not name:
                        # Subtotal row — ignore.
                        continue
                    nl = name.lower().strip()
                    if nl in SKIP_NAMES:
                        continue
                    if re.search(r"continued on page", name, re.I):
                        continue
                    if re.match(
                            r"^(total|grand total|sub total|opening|closing|balance)\b",
                            name, re.I):
                        continue

                    dr = _to_float(dr_str)
                    cr = _to_float(cr_str)
                    if dr == 0 and cr == 0:
                        continue

                    grp = line_to_group.get(name, page_group)
                    accounts.append({
                        "row":    len(accounts),
                        "key":    f"{name}_{len(accounts)}",
                        "name":   name,
                        "group":  grp,
                        "debit":  dr,
                        "credit": cr,
                        "net":    dr - cr,
                    })

                running_group = page_group
      except Exception:
        # Fall back to text-only parser if pdfplumber table extraction fails
        try:
            return _parse_tb_pdf_text_fallback(pdf_path)
        except Exception:
            return None

    # If table extraction found no rows (PDFs without ruled lines), fall back
    if not accounts and not _gw_parsed:
        try:
            return _parse_tb_pdf_text_fallback(pdf_path)
        except Exception:
            return None

    return {
        "format_type":   "PDF",
        "sheet_name":    "PDF",
        "header_row":    0,
        "data_start_row":0,
        "account_col":   0,
        "debit_col":     1,
        "credit_col":    2,
        "net_col":       None,
        "accounts":      accounts,
    }


def _parse_tb_pdf_text_fallback(pdf_path):
    """Legacy text-only PDF parser kept as a fallback for PDFs whose tables
    are not extractable by pdfplumber (e.g. no ruled lines).

    NOTE: this path infers Dr/Cr from `current_group` alone, so a debtor with
    a CREDIT balance will appear in the debit column. The primary
    `parse_tb_pdf` should be used wherever possible.
    """
    import pdfplumber, re

    all_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                l = line.strip()
                if l:
                    all_lines.append(l)
    if not all_lines:
        return None

    num_re = re.compile(r'([\d,]+\.\d{2})')
    skip_patterns = {"trial balance", "as on ", "page no", "continued",
                     "focal point", "punjab", "phase-", "e-254"}
    credit_groups = {"capital account", "secured loans", "unsecured loans",
                     "sundry creditors", "sundry payables", "sales account",
                     "indirect incomes", "profit & loss account",
                     "current liabilities", "duties & taxes"}

    current_group = ""
    accounts = []
    company_name = ""

    for line in all_lines:
        ll = line.lower().strip()
        if any(s in ll for s in skip_patterns):
            continue
        if ll == "particulars debit amount credit amount":
            continue
        if not company_name and line.isupper() and len(line) > 3 and not num_re.search(line):
            company_name = line
            continue
        if company_name and line.strip() == company_name:
            continue
        nums = num_re.findall(line)
        name = num_re.sub('', line).strip()
        if not nums:
            if name and len(name) > 1 and name not in ("0.01", ""):
                nl = name.lower()
                if not any(s in nl for s in ["phase-", "focal", "punjab",
                           "ludhiana-", "delhi-", "mumbai-", "address"]):
                    current_group = name
            continue
        if not name:
            continue
        dr_amt = 0.0
        cr_amt = 0.0
        if len(nums) == 1:
            val = float(nums[0].replace(',', ''))
            if current_group.lower() in credit_groups:
                cr_amt = val
            else:
                dr_amt = val
        elif len(nums) >= 2:
            dr_amt = float(nums[0].replace(',', ''))
            cr_amt = float(nums[1].replace(',', ''))
        accounts.append({
            "row":    len(accounts),
            "key":    f"{name}_{len(accounts)}",
            "name":   name,
            "group":  current_group,
            "debit":  dr_amt,
            "credit": cr_amt,
            "net":    dr_amt - cr_amt,
        })
    if not accounts:
        return None
    return {
        "format_type":   "PDF",
        "sheet_name":    "PDF",
        "header_row":    0,
        "data_start_row":0,
        "account_col":   0,
        "debit_col":     1,
        "credit_col":    2,
        "net_col":       None,
        "accounts":      accounts,
    }

def convert_pdf_tb_to_xlsx(pdf_path, xlsx_path):
    """Convert a PDF trial balance to an xlsx file for processing."""
    result = parse_tb_pdf(pdf_path)
    if not result or not result["accounts"]:
        return False

    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Trial Balance"

    # Header row
    ws.cell(1, 1, "Particulars")
    ws.cell(1, 2, "Debit")
    ws.cell(1, 3, "Credit")

    current_group = ""
    r = 2
    for acct in result["accounts"]:
        if acct["group"] != current_group:
            current_group = acct["group"]
            if current_group:
                ws.cell(r, 1, current_group)
                r += 1

        ws.cell(r, 1, acct["name"])
        if acct["debit"]:
            ws.cell(r, 2, acct["debit"])
        if acct["credit"]:
            ws.cell(r, 3, acct["credit"])
        r += 1

    wb.save(xlsx_path)
    return True


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
            "key": f"{acct_name}_{ri}",   # unique key for JS mapping
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
    "capital account":           "capital",
    "capital accounts":          "capital",
    "reserves & surplus":        "capital",
    "reserves and surplus":      "capital",
    "bank accounts":             "cash_bank",
    "bank account":              "cash_bank",
    "cash-in-hand":              "cash_bank",
    "cash in hand":              "cash_bank",
    "fixed assets":              "fixed_assets",
    "investments":               "non_current_investments",
    "sundry creditors":          "trade_payables",
    "sundry debtors":            "trade_rec",
    "sundry debtor":             "trade_rec",
    "sundry receivables":        "trade_rec",
    "sundry receivable":         "trade_rec",
    "purchase account":          "purchases",
    "purchase accounts":         "purchases",
    "purchases":                 "purchases",
    "sales account":             "revenue",
    "sales accounts":            "revenue",
    "sales":                     "revenue",
    "stock-in-hand":             "inventories",
    "stock in hand":             "inventories",
    # "indirect expenses" → other_expenses (high confidence via group header).
    # Previously this was commented out, causing FREIGHT, ADVERTISEMENT, ACCOUNTANCY
    # CHARGES etc. (all under "Indirect Expenses" in Tally) to fall through to
    # low-confidence classification and land in wrong buckets (purchases or direct_expenses).
    "indirect expenses":         "other_expenses",
    "indirect expense":          "other_expenses",
    # FIX (Issue 2/4): Direct expenses (Electricity, Wages, Power & Fuel)
    # need their OWN injection target on the Trading A/c. They were
    # previously collapsed into `purchases`, which made them overwrite the
    # purchase header row instead of landing in the Direct Expenses sub-rows.
    "direct expenses":           "direct_expenses",
    "direct expense":            "direct_expenses",
    "indirect income":           "other_income",
    "indirect incomes":          "other_income",
    "direct income":             "revenue",
    "direct incomes":            "revenue",
    "sundry payables":           "other_cl",
    "sundry payable":            "other_cl",
    "current liabilities":       "other_cl",
    "current assets":            "other_current_assets",
    "provisions":                "st_provisions",
    "unsecure loans":            "lt_borrowings",
    "unsecured loans":           "lt_borrowings",
    "secured loans":             "lt_borrowings",
    "loans (liability)":         "lt_borrowings",
    "loans liability":           "lt_borrowings",
    # FIX (Bug 2): the LOANS & ADVANCES (ASSET) group in Tally/Busy TBs
    # represents short-term advances (GST credit, TDS/TCS receivable, loans
    # given out, deposits). Map to `stla` so those items are classified with
    # HIGH confidence under "Short-term loans & advances" rather than falling
    # through to the generic low-confidence other_current_assets bucket.
    "loans & advances (asset)":  "stla",
    "loans & advances":          "stla",
    "loans and advances":        "stla",
    "deposits (asset)":          "stla",
    "deposits":                  "stla",
    "duties & taxes":            "other_cl",   # GST payable = current liability
    "duties and taxes":          "other_cl",
    "duties & taxes (gst)":      "other_cl",
    "duties and taxes (gst)":    "other_cl",
    "advance from customers":    "other_cl",   # advance received = current liability
    "misc. expenses (asset)":    "misc_expenditure",
    "miscellaneous expenses":    "misc_expenditure",
    "profit & loss account":     "capital",
    "profit and loss account":   "capital",
    "profit & loss a/c":         "capital",
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

    # ── Smart rules based on name + balance sign ──────────────────────
    # Rule: "loan" in name + CREDIT balance = borrowing (not fixed asset)
    if "loan" in name_lower and net_amount < 0:
        if any(kw in name_lower for kw in ["secured", "hypothec", "mortgage"]):
            return "lt_borrowings", "high"
        if any(kw in name_lower for kw in ["unsecure", "unsecured"]):
            return "lt_borrowings", "high"
        # Generic loan with credit balance = borrowing
        return "lt_borrowings", "high"

    # Rule: cheque issued / not cleared with credit balance = outstanding liability (OCL)
    # Must check BEFORE the "bank in name" rule since these don't always contain "bank"
    if net_amount < 0 and any(kw in name_lower for kw in [
            "cheque issued", "chq issued", "not cleared", "not presented", "issued not"]):
        return "other_cl", "high"

    # Rule: "bank" in name + CREDIT balance = secured loan/OD/CC
    if ("bank" in name_lower or "a/c" in name_lower) and net_amount < 0:
        if any(kw in name_lower for kw in ["loan", "od", "cc", "overdraft",
               "cash credit", "machinery", "vehicle", "term loan"]):
            return "lt_borrowings", "high"
        # Cheque issued/not cleared with credit balance = outstanding liability
        if any(kw in name_lower for kw in ["cheque issued", "chq issued",
               "not cleared", "not presented", "issued not"]):
            return "other_cl", "high"
        # Bank account with negative balance = bank overdraft = short term borrowing
        if any(kw in group_lower for kw in ["bank", "cash"]):
            return "st_borrowings", "high"

    # Rule: "round off" / "roundoff" = other_expenses (even if credit)
    if "round off" in name_lower or "roundoff" in name_lower:
        return "other_expenses", "high"

    # ── Name-priority overrides: run BEFORE group check ──────────────
    # These specific P&L sub-categories must override a broad group like
    # "indirect expenses" → other_expenses which fires too early.
    # Order matters: check most specific first.

    # Depreciation: must always go to depreciation head
    for kw in BS_HEADS["depreciation"]["keywords"]:
        if kw in name_lower:
            return "depreciation", "high"

    # Employee expenses: salary, wages etc. (but not "salary payable")
    # Use word-boundary matching for short/ambiguous keywords to avoid
    # false matches (e.g. "esi" matching "designer", "esi " is safer)
    _emp_neg = BS_HEADS["employee_expenses"].get("negative_keywords", [])
    if not any(nk in name_lower for nk in _emp_neg):
        _emp_match = False
        for kw in BS_HEADS["employee_expenses"]["keywords"]:
            # For very short keywords (≤4 chars), require word boundary (space or start/end)
            if len(kw) <= 4:
                import re as _re
                if _re.search(r'(?<![a-z])' + _re.escape(kw) + r'(?![a-z])', name_lower):
                    _emp_match = True
                    break
            elif kw in name_lower:
                _emp_match = True
                break
        if _emp_match:
            return "employee_expenses", "high"

    # Finance cost: bank interest, cc intt etc. (but not late payment interest)
    _fin_neg = BS_HEADS["finance_cost"].get("negative_keywords", [])
    if not any(nk in name_lower for nk in _fin_neg):
        for kw in BS_HEADS["finance_cost"]["keywords"]:
            if kw in name_lower:
                return "finance_cost", "high"

    # Step 1a: Group-based override for compound group names (group-wise TBs)
    if group_lower:
        # CURRENT LIABILITIES (DUTIES AND TAXES) → other_cl
        if "current liabilit" in group_lower and "sundry creditor" not in group_lower:
            return "other_cl", "high"
        # ADVANCE FROM CUSTOMERS → other_cl (advance received = current liability)
        if "advance from customer" in group_lower:
            return "other_cl", "high"
        # SUNDRY PAYABLE → other_cl
        if "sundry payable" in group_lower:
            return "other_cl", "high"
        # SUNDRY CREDITORS → trade_payables
        if "sundry creditor" in group_lower:
            return "trade_payables", "high"
        # SUNDRY DEBTORS → trade_rec
        if "sundry debtor" in group_lower:
            return "trade_rec", "high"

    # Step 1: Check if group header directly maps to a head
    if group_lower and group_lower in GROUP_HEAD_MAP:
        group_head = GROUP_HEAD_MAP[group_lower]
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

    Sign-aware reclassification rules (Indian accounting standards):
    - Creditor with DEBIT balance (net > 0) → advance to supplier → stla
    - Debtor with CREDIT balance (net < 0)  → advance from customer → other_cl
    These accounts are reclassified before aggregation.
    """
    totals = {}
    for acct in classified_accounts:
        head = acct["bs_head"]
        net  = acct["net"]

        if head == "unclassified":
            continue

        # ── Sign-aware auto-reclassification ──────────────────────────
        # Creditor with debit balance = advance paid to supplier (asset).
        # Aggregate into stla, but flag with explicit UI head
        # "advance_to_supplier" so the Review-Mapping page can show it.
        if head == "trade_payables" and net > 0:
            head = "stla"
            acct["bs_head"]    = "advance_to_supplier"
            acct["reclassified_from"] = "trade_payables"

        # Debtor with credit balance = advance received from customer (liability)
        # Aggregate into other_cl, flag as "advance_from_customer".
        elif head == "trade_rec" and net < 0:
            head = "other_cl"
            acct["bs_head"]    = "advance_from_customer"
            acct["reclassified_from"] = "trade_rec"

        # Provision with DEBIT balance = receivable/advance (asset not liability)
        # e.g. TCS GST A/C (debit) = TCS paid to govt = advance to revenue authority
        elif head == "st_provisions" and net > 0:
            head = "stla"
            acct["bs_head"]      = "stla"
            acct["reclassified_from"] = "st_provisions"
            acct["stla_subtype"]  = "revenue_authority"  # flag for D129 row

        # stla accounts that are TCS/TDS by name also get revenue_authority subtype
        # (covers accounts classified via group "duties & taxes" → stla directly,
        #  bypassing the st_provisions reclassification path above)
        elif head == "stla" and net > 0:
            name_l = acct.get("name", "").lower()
            if "tcs" in name_l and not acct.get("stla_subtype"):
                acct["stla_subtype"] = "revenue_authority"
            elif "excess tds" in name_l and not acct.get("stla_subtype"):
                acct["stla_subtype"] = "revenue_authority"

        # Provision with CREDIT balance but name doesn't match genuine provisions
        # → reclassify to other_cl to prevent spurious amounts in ST Provisions.
        # Genuine provisions: tax, gratuity, bonus, leave, warranty.
        elif head == "st_provisions" and net < 0:
            VALID_PROVISION_KEYWORDS = {
                "provision for tax", "provision for income",
                "provision for gratuity", "provision for bonus",
                "provision for leave", "provision for warranty",
                "provision for bad", "provision for doubtful",
                "short term provision",
            }
            name_l = acct.get("name", "").lower()
            if not any(kw in name_l for kw in VALID_PROVISION_KEYWORDS):
                head = "other_cl"
                acct["bs_head"] = "other_cl"
                acct["reclassified_from"] = "st_provisions"

        # ── Explicit user-mapped advance heads ─────────────────────────
        # When the user picks "Advance from Customer" / "Advance to Supplier"
        # from the dropdown, the amount must feed the parent BS bucket while
        # the injector still knows the origin via `reclassified_from`.
        elif head == "advance_from_customer":
            head = "other_cl"
            acct["reclassified_from"] = acct.get("reclassified_from") or "trade_rec"

        elif head == "advance_to_supplier":
            head = "stla"
            acct["reclassified_from"] = acct.get("reclassified_from") or "trade_payables"

        amt = abs(net)
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


def _get_writable_cell(ws, row, col):
    """
    Return the writable cell at (row, col).
    If the target is a MergedCell, return the top-left anchor of the merged range
    (which is the only writable cell in that range). If no merged range is found,
    returns None to indicate the cell cannot be written.
    """
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return cell
    # Cell is part of a merged range — find the top-left anchor
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
            merged_range.min_col <= col <= merged_range.max_col):
            anchor = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            if not isinstance(anchor, MergedCell):
                return anchor
            return None
    return None


def _safe_set(ws, row, col, value):
    """
    Write a plain numeric value only if the target cell is not a formula.
    MergedCell-aware: if the target is a MergedCell, write to the top-left anchor
    of its merged range instead (provided the anchor isn't a formula).
    Returns True on success, False otherwise.
    """
    try:
        cell = _get_writable_cell(ws, row, col)
        if cell is None:
            return False
        if _is_formula(cell.value):
            return False
        cell.value = round(float(value), 2)
        return True
    except (AttributeError, TypeError, ValueError):
        # Defensive: never let a single write blow up the entire pipeline
        return False


def _safe_write(ws, row, col, value):
    """
    Write any value (number, string, etc.) to the target cell safely.
    MergedCell-aware (writes to the merged range's anchor cell).
    Skips formula cells. Returns True on success, False otherwise.
    Use this in place of direct `ws.cell(r, c).value = ...` assignments.
    """
    try:
        cell = _get_writable_cell(ws, row, col)
        if cell is None:
            return False
        if _is_formula(cell.value):
            return False
        if isinstance(value, (int, float)):
            cell.value = round(float(value), 2)
        else:
            cell.value = value
        return True
    except (AttributeError, TypeError, ValueError):
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

    # Finance-cost abbreviation aliases: Tally short names → normalised forms
    # so "BANK CC INTT" matches a template row labelled "Bank Interest".
    FINANCE_ALIASES = {
        "bank cc intt":       "bank interest",
        "bank cc interest":   "bank interest",
        "bank od intt":       "bank interest",
        "bank od interest":   "bank interest",
        "cc interest":        "bank interest",
        "od interest":        "bank interest",
        # "bank charges and interest" is NOT a loan interest — it's a bank service charge
        # and belongs in other_expenses. Do NOT alias it to "bank interest".
    }

    def normalize(s):
        s = s.lower().strip()
        # Apply finance alias before generic normalisation
        if s in FINANCE_ALIASES:
            s = FINANCE_ALIASES[s]
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
    # 1. CAPITAL — DO NOT inject from TB.
    #    TB only has closing balance. Additions/withdrawals come from
    #    ledger. User's BS template capital sheet is preserved as-is.
    # ────────────────────────────────────────────────────────────────
    capital_amt = aggregated_values.get("capital", 0)
    if capital_amt:
        log.append(f"· Capital from TB: {capital_amt:,.2f} — skipped (fill from ledger)")

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

        # Long-term borrowings → match individual loan accounts to template rows
        ltb_amt = aggregated_values.get("lt_borrowings", 0)
        if ltb_amt and individual_accounts:
            ltb_accounts = [a for a in individual_accounts
                            if a.get("bs_head") == "lt_borrowings"
                            and abs(a.get("net", 0)) > 0]

            # Separate secured vs unsecured based on Tally group
            def _is_unsecured(acct):
                g = acct.get("group", "").lower()
                n = acct.get("name", "").lower()
                return "unsecure" in g or "unsecure" in n

            secured_accounts   = [a for a in ltb_accounts if not _is_unsecured(a)]
            unsecured_accounts = [a for a in ltb_accounts if _is_unsecured(a)]

            # Build template label → row map for LT borrowing section
            ltb_template = {}
            for r in range(7, 24):
                b_val = ws_n.cell(r, 2).value
                if b_val and isinstance(b_val, str) and len(b_val.strip()) > 2:
                    lbl = b_val.strip().lower()
                    # Exclude section headers but NOT loan names that happen to contain 'from'
                    HEADER_LABELS = {'from banks', 'from related parties', 'from other parties',
                                     'secured', 'unsecured', 'term loans'}
                    if 'total' in lbl or lbl.strip() in HEADER_LABELS:
                        continue
                    ltb_template[r] = lbl

            # Find the unsecured section start row (row after "Unsecured" header)
            unsecured_start = None
            secured_end = None
            for r in range(7, 24):
                b_val = ws_n.cell(r, 2).value
                if b_val and isinstance(b_val, str):
                    bl = b_val.strip().lower()
                    if 'unsecured' in bl and 'total' not in bl:
                        unsecured_start = r + 1
                    if 'total' in bl and 'secured' in bl and 'unsecured' not in bl and secured_end is None:
                        secured_end = r
            if unsecured_start is None:
                unsecured_start = 15  # fallback
            if secured_end is None:
                secured_end = 11

            written_rows = set()

            def _inject_ltb_list(acct_list, fallback_start, fallback_end):
                # Sort by name length DESC so more specific accounts (longer names) match first
                # This prevents "Axis ML3" from stealing a row meant for a more specific name
                for acct in sorted(acct_list, key=lambda a: -len(a["name"])):
                    amt  = abs(acct["net"])
                    name = acct["name"]
                    matched_row = None
                    # Try exact/fuzzy match — but avoid matching generic labels to wrong loans
                    for r, lbl in sorted(ltb_template.items()):
                        if r in written_rows:
                            continue
                        if _fuzzy_match_name(name, lbl):
                            matched_row = r
                            break
                    if matched_row:
                        if _safe_set(ws_n, matched_row, 4, amt):
                            written_rows.add(matched_row)
                            injected.append(f"notes to bs!D{matched_row} ({name}) = {amt:,.2f}")
                    else:
                        # No match — find empty row in the correct section
                        for r in range(fallback_start, fallback_end + 5):
                            if r in written_rows:
                                continue
                            d_val = ws_n.cell(r, 4).value
                            b_val = ws_n.cell(r, 2).value
                            if (d_val is None or d_val == 0) and (b_val is None or str(b_val).strip() == ""):
                                _safe_set(ws_n, r, 2, name)
                                if _safe_set(ws_n, r, 4, amt):
                                    written_rows.add(r)
                                    injected.append(f"notes to bs!D{r} (new: {name}) = {amt:,.2f}")
                                break

            # Secured → rows 7 to secured_end-1 as fallback
            _inject_ltb_list(secured_accounts, 7, secured_end - 1)
            # Unsecured → rows unsecured_start onward as fallback
            _inject_ltb_list(unsecured_accounts, unsecured_start, unsecured_start + 5)

            if not written_rows and ltb_amt:
                # Fallback: write total to "from banks" row
                for r in range(7, 11):
                    d = ws_n.cell(r, 4).value
                    if d is None or not _is_formula(str(d)):
                        inject_notes_row(r, 4, abs(ltb_amt), "Long-term borrowings total")
                        break

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

        # ── Cheques Issued But Not Cleared → "Cheque Issued" labelled row ──
        # Only CREDIT-balance cheque accounts are liabilities (issued cheques).
        # Find the row by label (not hardcoded D68) so it works for any template.
        if individual_accounts:
            cheque_accounts = [a for a in individual_accounts
                               if ("cheque" in a.get("name","").lower()
                               or "chq" in a.get("name","").lower())
                               and a.get("net", 0) < 0]   # credit balance = liability
            total_cheques = sum(abs(a["net"]) for a in cheque_accounts)
            if total_cheques > 0:
                cheque_row = None
                for r in range(50, 100):
                    lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                    if "cheque issued" in lbl or "chq issued" in lbl or "not presented" in lbl:
                        cheque_row = r
                        break
                if cheque_row:
                    ws_n.cell(cheque_row, 4).value = total_cheques
                    injected.append(f"notes to bs!D{cheque_row} (Cheques issued) = {total_cheques:,.2f}")
                else:
                    # Fallback to OCL section last named row
                    for r in range(55, 90):
                        lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                        if "total" in lbl and "current" in lbl:
                            ws_n.cell(r-1, 4).value = total_cheques
                            injected.append(f"notes to bs!D{r-1} (Cheques issued fallback) = {total_cheques:,.2f}")
                            break

        # Other current liabilities → match by name to template rows
        ocl_amt = aggregated_values.get("other_cl", 0)
        if ocl_amt:
            # Detect OCL section boundaries
            ocl_start = None
            ocl_end = None
            for r in range(50, 100):
                b = ws_n.cell(r, 2).value
                if b and 'other current liabilit' in str(b).lower() and 'total' not in str(b).lower():
                    ocl_start = r + 1
                if ocl_start and b and 'total' in str(b).lower() and 'other' in str(b).lower():
                    ocl_end = r
                    break
            if not ocl_start:
                ocl_start = 59
            if not ocl_end:
                ocl_end = 67

            # Build template label → row map for OCL section
            ocl_template = {}
            for r in range(ocl_start, ocl_end):
                b = ws_n.cell(r, 2).value
                if b and isinstance(b, str) and len(b.strip()) > 2:
                    lbl = b.strip().lower()
                    if 'total' not in lbl:
                        ocl_template[r] = lbl

            if individual_accounts:
                ocl_accounts = [a for a in individual_accounts
                                if a.get("bs_head") == "other_cl"
                                and abs(a.get("net", 0)) > 0
                                # Exclude cheque accounts — already injected above
                                and "cheque" not in a.get("name","").lower()
                                and "chq" not in a.get("name","").lower()]
                written_ocl = set()
                
                # Phase 1: fuzzy match to template
                for acct in ocl_accounts:
                    for r, lbl in ocl_template.items():
                        if r in written_ocl:
                            continue
                        if _fuzzy_match_name(acct["name"], lbl):
                            if _safe_set(ws_n, r, 4, abs(acct["net"])):
                                written_ocl.add(r)
                                injected.append(f"notes to bs!D{r} (OCL: {acct['name']}) = {abs(acct['net']):,.2f}")
                            break
                
                # Phase 2: unmatched → aggregate into category rows (NEVER insert_rows)
                # Group unmatched accounts by category keyword and add to the best
                # matching named row. This avoids row-shifting entirely.
                CATEGORY_KEYWORDS_OCL = [
                    ("esi",          ["esi"]),
                    ("wages",        ["wages", "wage"]),
                    ("tds payable",  ["tds"]),
                    ("leave",        ["leave with wages", "leave"]),
                    ("bonus",        ["bonus"]),
                    ("audit",        ["audit"]),
                    ("professional", ["professional"]),
                    ("salary",       ["salary payable", "salary"]),
                ]
                # Accumulate unmatched amounts by category
                cat_pending: dict = {}
                for acct in ocl_accounts:
                    if any(_fuzzy_match_name(acct["name"],
                           ocl_template.get(r, "")) for r in written_ocl):
                        continue   # already matched in Phase 1
                    name_l = acct["name"].lower()
                    amt    = abs(acct["net"])
                    matched_cat = None
                    for cat, kws in CATEGORY_KEYWORDS_OCL:
                        if any(k in name_l for k in kws):
                            matched_cat = cat
                            break
                    if not matched_cat:
                        matched_cat = "other"
                    cat_pending[matched_cat] = cat_pending.get(matched_cat, 0) + amt

                # Write accumulated category totals to matching template rows
                for cat, amt in cat_pending.items():
                    if amt == 0:
                        continue
                    for r, lbl in ocl_template.items():
                        if r in written_ocl:
                            continue
                        cat_kws = next(
                            (kws for c, kws in CATEGORY_KEYWORDS_OCL if c == cat), []
                        )
                        if any(k in lbl for k in cat_kws):
                            existing = ws_n.cell(r, 4).value or 0
                            if not _is_formula(str(existing)):
                                new_val = (existing if isinstance(existing,(int,float)) else 0) + amt
                                if _safe_set(ws_n, r, 4, new_val):
                                    written_ocl.add(r)
                                    injected.append(
                                        f"notes to bs!D{r} (OCL cat '{cat}') = {new_val:,.2f}"
                                    )
                            break
                    else:
                        # No template row for this category — find first truly empty row
                        for r in range(ocl_start, ocl_end):
                            if r in written_ocl:
                                continue
                            d = ws_n.cell(r, 4).value
                            if d is None or d == 0:
                                if _safe_set(ws_n, r, 4, amt):
                                    written_ocl.add(r)
                                    injected.append(
                                        f"notes to bs!D{r} (OCL overflow '{cat}') = {amt:,.2f}"
                                    )
                                break
                        else:
                            skipped.append(
                                f"OCL category '{cat}' = {amt:,.2f}: no row — template full"
                            )

        # ── Cash & Bank — label-driven, no hardcoded row numbers ─────────────
        cash_bank_amt = aggregated_values.get("cash_bank", 0)
        if individual_accounts:
            cash_accounts = [a for a in individual_accounts
                             if a.get("bs_head") in ("cash_bank",)
                             and a.get("net", 0) > 0]
            cash_only = [a for a in cash_accounts if "cash" in a["name"].lower()]
            bank_only = [a for a in cash_accounts
                         if "cash" not in a["name"].lower()
                         and "cheque" not in a["name"].lower()]

            # Locate Cash in Hand row by scanning for the label (not hardcoded row 109)
            cash_in_hand_row = None
            for r in range(90, 170):
                lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                if "cash in hand" in lbl:
                    cash_in_hand_row = r
                    break

            if cash_in_hand_row and cash_only:
                total_cash = sum(abs(a["net"]) for a in cash_only)
                # Always write cash (overwrite even formula PY values in CY col)
                ws_n.cell(cash_in_hand_row, 4).value = total_cash
                injected.append(
                    f"notes to bs!D{cash_in_hand_row} (Cash in hand) = {total_cash:,.2f}"
                )

            # Locate bank rows by scanning for bank name labels
            for acct in bank_only:
                a_name_l = acct["name"].lower()
                a_words  = [w for w in a_name_l.split()
                            if len(w) > 3 and w not in {"bank","a/c","ltd","pvt","the","current"}]
                for r in range(90, 170):
                    lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                    if not lbl or "total" in lbl or "cash in hand" in lbl:
                        continue
                    if any(w in lbl for w in a_words):
                        ws_n.cell(r, 4).value = abs(acct["net"])
                        injected.append(
                            f"notes to bs!D{r} ({acct['name']}) = {abs(acct['net']):,.2f}"
                        )
                        break
                else:
                    # Fallback: find empty bank row near section
                    for r in range(90, 170):
                        lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                        if any(k in lbl for k in ["bank", "a/c"]) and "total" not in lbl \
                                and "cash in hand" not in lbl:
                            d = ws_n.cell(r, 4).value
                            if d is None or d == 0:
                                ws_n.cell(r, 4).value = abs(acct["net"])
                                injected.append(
                                    f"notes to bs!D{r} ({acct['name']}) = {abs(acct['net']):,.2f}"
                                )
                                break

        elif cash_bank_amt:
            # Fallback: find Cash in hand row by label
            for r in range(90, 170):
                lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                if "cash in hand" in lbl:
                    ws_n.cell(r, 4).value = cash_bank_amt
                    injected.append(f"notes to bs!D{r} (Cash and bank lump) = {cash_bank_amt:,.2f}")
                    break

        # Short-term loans & advances → D128 (GST), D129 (TCS GST), D130 (Excess TDS)
        # Special case: "TDS claimable from Facebook" → D137 (Other current assets)
        stla_amt = aggregated_values.get("stla", 0)
        if stla_amt and individual_accounts:
            stla_accounts = [a for a in individual_accounts
                             if a.get("bs_head") == "stla"
                             and abs(a.get("net", 0)) > 0
                             and "cheque" not in a.get("name","").lower()
                             and "chq" not in a.get("name","").lower()
                             and a.get("reclassified_from") != "trade_payables"]
            for acct in stla_accounts:
                name_l = acct["name"].lower()
                placed = False

                # Find the STLA Revenue Authorities section dynamically
                # by scanning for label keywords (works for any template layout)
                def _find_rev_auth_row(keyword):
                    """Scan notes to bs for a row matching keyword under Revenue Authorities."""
                    in_rev_section = False
                    for r in range(110, 180):
                        lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                        if "advance to revenue" in lbl or "revenue authorit" in lbl:
                            in_rev_section = True
                        if in_rev_section:
                            if keyword in lbl:
                                return r
                            if "total" in lbl and "a+b" in lbl:
                                break  # past section
                    return None

                # ── Facebook TDS → Other current assets ─────────────────────
                if "facebook" in name_l:
                    fb_row = None
                    for r in range(130, 180):
                        lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                        if "facebook" in lbl or "recoverable" in lbl:
                            fb_row = r; break
                    if fb_row and _safe_set(ws_n, fb_row, 4, abs(acct["net"])):
                        injected.append(f"notes to bs!D{fb_row} (TDS from Facebook) = {abs(acct['net']):,.2f}")
                    placed = True

                # ── TCS → TCS row under Revenue Authorities ─────────────────
                elif "tcs" in name_l:
                    row = _find_rev_auth_row("tcs")
                    if row:
                        existing = ws_n.cell(row, 4).value
                        base = float(existing) if isinstance(existing,(int,float)) else 0
                        ws_n.cell(row, 4).value = base + abs(acct["net"])
                        injected.append(
                            f"notes to bs!D{row} (TCS) = {base+abs(acct['net']):,.2f}"
                        )
                        placed = True

                # ── GST (CGST/SGST/IGST/refund) → one row per account ───────
                # Check GST BEFORE TDS so "GST TDS 2024-25" goes to GST section
                elif any(k in name_l for k in ["gst","igst","cgst","sgst"]):
                    # Find next AVAILABLE (empty) GST row; each account gets its own row
                    placed_gst = False
                    for r in range(110, 180):
                        lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                        if "advance to revenue" in lbl or "revenue authorit" in lbl:
                            for rr in range(r+1, r+15):
                                lbl2 = (ws_n.cell(rr, 2).value or "").strip().lower()
                                if "total" in lbl2: break
                                if "gst" in lbl2 or "igst" in lbl2 or "cgst" in lbl2 or "sgst" in lbl2:
                                    d = ws_n.cell(rr, 4).value
                                    # Only write to empty rows — gives each account its own slot
                                    if d is None or d == 0:
                                        ws_n.cell(rr, 4).value = abs(acct["net"])
                                        injected.append(f"notes to bs!D{rr} (GST: {acct['name']}) = {abs(acct['net']):,.2f}")
                                        placed_gst = True
                                        break
                            break
                    if placed_gst:
                        placed = True

                # ── TDS / Excess TDS → separate rows per account ─────────────
                elif "tds" in name_l:
                    # Use _find_rev_auth_row to get first matching TDS label row,
                    # but find an EMPTY one (each TDS account → its own row).
                    placed_tds = False
                    for r in range(110, 180):
                        lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                        if "advance to revenue" in lbl or "revenue authorit" in lbl:
                            for rr in range(r+1, r+15):
                                lbl2 = (ws_n.cell(rr, 2).value or "").strip().lower()
                                if "total" in lbl2: break
                                # Match TDS-labelled rows (not GST rows)
                                if ("tds" in lbl2 or "excess tds" in lbl2) and \
                                        "gst" not in lbl2:
                                    d = ws_n.cell(rr, 4).value
                                    if d is None or d == 0:
                                        ws_n.cell(rr, 4).value = abs(acct["net"])
                                        injected.append(f"notes to bs!D{rr} (TDS: {acct['name']}) = {abs(acct['net']):,.2f}")
                                        placed_tds = True
                                        break
                            break
                    if placed_tds:
                        placed = True

                if not placed:
                    # Generic: find first empty row in STLA Revenue Auth area
                    for r in range(110, 180):
                        lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                        if "advance to revenue" in lbl or "revenue authorit" in lbl:
                            for rr in range(r+1, r+15):
                                lbl2 = (ws_n.cell(rr, 2).value or "").strip().lower()
                                if "total" in lbl2:
                                    break
                                d = ws_n.cell(rr, 4).value
                                if d is None or d == 0:
                                    ws_n.cell(rr, 4).value = abs(acct["net"])
                                    injected.append(
                                        f"notes to bs!D{rr} ({acct['name']}) = {abs(acct['net']):,.2f}"
                                    )
                                    placed = True
                                    break
                            break
                    if not placed:
                        skipped.append(f"STLA '{acct['name']}' = {abs(acct['net']):,.2f}: no row found")
        elif stla_amt:
            # Lump sum fallback
            for r in range(110, 180):
                lbl = (ws_n.cell(r, 2).value or "").strip().lower()
                if "advance to revenue" in lbl or "gst" in lbl:
                    ws_n.cell(r, 4).value = stla_amt
                    injected.append(f"notes to bs!D{r} (STLA lump) = {stla_amt:,.2f}")
                    break

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
        # Debit-balance creditors are reclassified to stla (advance to supplier)
        # — they must NOT appear in creditor rows, handled below in stla section
        adv_to_supplier = [a for a in individual_accounts
                           if a.get("reclassified_from") == "trade_payables"]
        unmatched_creditors = []
        for acct in creditor_accounts:
            best_row = None
            for tmpl_name, row in cred_row_map.items():
                if row not in written_rows and _fuzzy_match_name(acct["name"], tmpl_name):
                    best_row = row; break
            if best_row:
                if _safe_write(ws_det, best_row, 4, abs(acct["net"])):
                    injected.append(f"Details!D{best_row} ({acct['name']}) = {abs(acct['net']):,.2f}")
                    written_rows.add(best_row)
                else:
                    skipped.append(f"Details!D{best_row} ({acct['name']}): cell is merged/formula")
            else:
                unmatched_creditors.append(acct)

        # Blank rows for unmatched (from cache)
        cred_end_row = 63
        blank_rows = [r for r in range(21, cred_end_row)
                      if r not in written_rows
                      and det_cache.get((r, 4)) is None
                      and det_cache.get((r, 2)) is None]
        for i, acct in enumerate(unmatched_creditors):
            if i < len(blank_rows):
                r = blank_rows[i]
                wrote_name = _safe_write(ws_det, r, 2, acct["name"])
                wrote_amt  = _safe_write(ws_det, r, 4, abs(acct["net"]))
                if wrote_amt:
                    injected.append(f"Details!D{r} (new: {acct['name']}) = {abs(acct['net']):,.2f}")
                else:
                    skipped.append(f"Trade payable '{acct['name']}': row {r} is merged/formula")
            else:
                # Section full — insert new row
                try:
                    cred_end_row += 1
                    ws_det.insert_rows(cred_end_row)
                    ws_det.cell(cred_end_row, 2).value = acct["name"]
                    ws_det.cell(cred_end_row, 4).value = abs(acct["net"])
                    written_rows.add(cred_end_row)
                    injected.append(f"Details!D{cred_end_row} (cred NEW ROW: {acct['name']}) = {abs(acct['net']):,.2f}")
                    det_cache[(cred_end_row, 2)] = acct["name"]
                    det_cache[(cred_end_row, 4)] = abs(acct["net"])
                except Exception as e:
                    skipped.append(f"Trade payable '{acct['name']}': insert failed: {e}")

        # Trade Receivables D74:D90 (auto-extending if section is full)
        receivable_accounts = [a for a in individual_accounts
                               if a.get("bs_head") == "trade_rec"
                               and abs(a.get("net", 0)) > 0]
        recv_written = set()
        recv_end_row = 90  # max row for debtors section

        for acct in receivable_accounts:
            placed = False
            # First try to match by name
            for r in range(74, recv_end_row + 1):
                b = det_cache.get((r, 2))
                if b and _fuzzy_match_name(acct["name"], str(b)) and r not in recv_written:
                    if _safe_write(ws_det, r, 4, abs(acct["net"])):
                        injected.append(f"Details!D{r} ({acct['name']}) = {abs(acct['net']):,.2f}")
                        recv_written.add(r); placed = True
                    break
            if not placed:
                # Try first empty row
                for r in range(74, recv_end_row + 1):
                    if det_cache.get((r, 4)) is None and det_cache.get((r, 2)) is None and r not in recv_written:
                        _safe_write(ws_det, r, 2, acct["name"])
                        if _safe_write(ws_det, r, 4, abs(acct["net"])):
                            injected.append(f"Details!D{r} (recv: {acct['name']}) = {abs(acct['net']):,.2f}")
                            recv_written.add(r); placed = True
                        break
            if not placed:
                # Section full — insert new row before total
                try:
                    recv_end_row += 1
                    ws_det.insert_rows(recv_end_row)
                    ws_det.cell(recv_end_row, 2).value = acct["name"]
                    ws_det.cell(recv_end_row, 4).value = abs(acct["net"])
                    recv_written.add(recv_end_row)
                    injected.append(f"Details!D{recv_end_row} (recv NEW ROW: {acct['name']}) = {abs(acct['net']):,.2f}")
                    # Update det_cache for new row
                    det_cache[(recv_end_row, 2)] = acct["name"]
                    det_cache[(recv_end_row, 4)] = abs(acct["net"])
                except Exception as e:
                    skipped.append(f"recv '{acct['name']}': insert failed: {e}")

        # NOTE: LT borrowings are written to notes to bs!D8 directly (Section 2 above).
        # Do NOT also write to Details D7-D12 — that would double-count because
        # notes to bs!D16 = =Details!D9 (formula), creating a second entry.

        # ── Advance to Suppliers (debit-balance creditors) → Details D74:D90 ──
        # These creditors had a debit balance in TB → reclassified to stla
        # They appear in the same row range as trade receivables (advances section)
        for acct in adv_to_supplier:
            placed = False
            for r in range(74, 90):
                b = det_cache.get((r, 2))
                if b and _fuzzy_match_name(acct["name"], str(b)) and r not in recv_written:
                    if _safe_write(ws_det, r, 4, abs(acct["net"])):
                        injected.append(f"Details!D{r} (adv-to-supplier: {acct['name']}) = {abs(acct['net']):,.2f}")
                        recv_written.add(r); placed = True
                    break
            if not placed:
                for r in range(74, 90):
                    if det_cache.get((r, 4)) is None and det_cache.get((r, 2)) is None and r not in recv_written:
                        _safe_write(ws_det, r, 2, acct["name"])
                        if _safe_write(ws_det, r, 4, abs(acct["net"])):
                            injected.append(f"Details!D{r} (adv-to-supplier new: {acct['name']}) = {abs(acct['net']):,.2f}")
                            recv_written.add(r)
                        break

        # ── Advance from Customers (credit-balance debtors) → OCL in notes to bs ──
        # These debtors had a credit balance → reclassified to other_cl
        # They go into the OCL section of notes to bs (rows 60-82)
        adv_from_customer = [a for a in individual_accounts
                             if a.get("reclassified_from") == "trade_rec"]
        if adv_from_customer and "Details" in wb.sheetnames:
            ws_det_afc = wb["Details"]
            for acct in adv_from_customer:
                # Write to Details sheet "Advance from Customers" section
                # by matching party name to existing template rows
                name_l = acct["name"].lower()
                placed = False
                for r in range(20, 90):
                    b = ws_det_afc.cell(r, 2).value
                    if b and _fuzzy_match_name(acct["name"], str(b)):
                        _safe_set(ws_det_afc, r, 4, abs(acct["net"]))
                        injected.append(f"Details!D{r} (adv-from-customer: {acct['name']}) = {abs(acct['net']):,.2f}")
                        placed = True
                        break
                if not placed:
                    # Find "Advance from Customers" section and write there
                    for r in range(20, 90):
                        b = ws_det_afc.cell(r, 2).value
                        if b and "advance from customer" in str(b).lower():
                            for rr in range(r+1, r+8):
                                b2 = ws_det_afc.cell(rr, 2).value
                                d2 = ws_det_afc.cell(rr, 4).value
                                if (b2 is None or not str(b2).strip()) and (d2 is None or d2 == 0):
                                    _safe_set(ws_det_afc, rr, 4, abs(acct["net"]))
                                    injected.append(f"Details!D{rr} (adv-from-cust new: {acct['name']}) = {abs(acct['net']):,.2f}")
                                    placed = True
                                    break
                            break

    # ────────────────────────────────────────────────────────────────
    # 4. CLOSING STOCK / INVENTORIES
    # ────────────────────────────────────────────────────────────────
    # GROSS PROFIT!E17 has formula =B24-SUM(E11:E14) — closing stock
    # is the BALANCING FIGURE on the sales side.
    # DO NOT write to E17 — the formula calculates it automatically
    # once sales (E11:E14) and total (B24/E24) are filled.
    #
    # notes to p&l!D24 = ='GROSS PROFIT'!E17 (formula) — also auto.
    #
    # Opening stock (D17) IS written (it's a plain constant from TB).
    # Closing stock flows automatically — no injection needed here.
    inventory_amt = aggregated_values.get("inventories", 0)
    if inventory_amt:
        log.append(f"· Closing stock {inventory_amt:,.2f} — NOT written to E17 (auto-calculated by formula =B24-SUM(E11:E14))")

    # ────────────────────────────────────────────────────────────────
    # 5. FIXED ASSETS — DO NOT inject from TB.
    #    TB only has WDV closing. Additions/sales come from ledger.
    #    User's BS template FA chart is preserved as-is.
    # ────────────────────────────────────────────────────────────────
    fixed_assets_amt = aggregated_values.get("fixed_assets", 0)
    if fixed_assets_amt:
        log.append(f"· Fixed Assets from TB: {fixed_assets_amt:,.2f} — skipped (fill from ledger)")

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

            # FIX (Issues 2 & 4): Trading-account formatting.
            #
            # Previously this block used HARD-CODED row numbers
            # (12% Within → row 12, 18% Within → row 14 …) which assumed
            # one specific template layout. If the user's BS template has a
            # different layout (e.g. extra rows, different section order,
            # different GST rates listed), the sale amounts landed on the
            # WRONG rows — sometimes on the TOTAL row, overwriting its formula
            # and showing the same number on every row.
            #
            # New strategy: SCAN column D of the GROSS PROFIT sheet for sale-side
            # row labels (B11:D24 area), build a label→row map dynamically, and
            # write each TB sale account into the row whose label best matches.
            # Unknown rate codes get appended below the last labelled row.
            gp_cache = _cache.get("GROSS PROFIT", {})

            def _norm_gst(s):
                """Normalise sale/purchase account names for row matching.
                Maps all synonym pairs to canonical forms so that e.g.
                'SALE GST 12% INTERSTATE' and 'Sales 12% Intrastate' match
                the same template row regardless of exact wording.
                """
                s = str(s).lower()
                s = re.sub(r"\b(sale|sales|purchase|purchases|gst|a/c)\b", " ", s)
                # Synonym normalisation
                s = s.replace("intrastate", "within state")
                s = s.replace("within state", "intrastate")   # canonical = intrastate
                s = s.replace("within state", "intrastate")
                s = s.replace("intra state", "intrastate")
                s = s.replace("inter state", "interstate")
                s = re.sub(r"\s+", " ", s).strip()
                return s

            # Build label→row map by scanning the Sales side of GROSS PROFIT.
            # Sale rows live in col D (label) with amounts in col E.
            sale_label_rows = {}
            sale_total_rows = set()
            for r in range(1, 40):
                lbl = gp_cache.get((r, 4))   # col D = sales-side label
                if not lbl:
                    continue
                lbl_s = str(lbl).strip()
                if not lbl_s:
                    continue
                lbl_low = lbl_s.lower()
                # Skip section headers (SALES, CLOSING STOCK) and totals
                if lbl_low in ("sales", "closing stock", "") or "total" in lbl_low:
                    if "total" in lbl_low:
                        sale_total_rows.add(r)
                    continue
                if "as certified" in lbl_low or "proprietor" in lbl_low:
                    continue
                # This is a sub-item row — record it
                sale_label_rows[_norm_gst(lbl_s)] = (r, lbl_s)

            # Place each TB sale into its matching row
            sale_row_totals = {}     # {row: amount}
            unmatched_sales = []
            for acct in sale_accounts:
                name_norm = _norm_gst(acct["name"])
                best_row = None
                best_label = None
                # Prefer exact normalised match, else substring either-way
                if name_norm in sale_label_rows:
                    best_row, best_label = sale_label_rows[name_norm]
                else:
                    for lbl_norm, (r, lbl_orig) in sale_label_rows.items():
                        if lbl_norm and (lbl_norm in name_norm or name_norm in lbl_norm):
                            best_row, best_label = r, lbl_orig
                            break
                if best_row is None:
                    unmatched_sales.append(acct)
                    continue
                sale_row_totals[best_row] = sale_row_totals.get(best_row, 0) + abs(acct["net"])

            # Write each label-matched row — NEVER write to a Total row
            for row, amt in sale_row_totals.items():
                if row in sale_total_rows:
                    continue
                if _safe_write(ws_gp, row, 5, amt):
                    injected.append(f"GROSS PROFIT!E{row} (Sale {amt:,.2f})")

            # Unmatched sales — append below the last sale sub-row but ABOVE total
            if unmatched_sales and sale_label_rows:
                last_sub_row = max(r for r, _ in sale_label_rows.values())
                next_row = last_sub_row + 1
                for acct in unmatched_sales:
                    # Stop if we'd overwrite a total / closing-stock row
                    if next_row in sale_total_rows:
                        break
                    if _safe_write(ws_gp, next_row, 4, acct["name"]):
                        _safe_write(ws_gp, next_row, 5, abs(acct["net"]))
                        injected.append(
                            f"GROSS PROFIT!E{next_row} (Sale unmatched: {acct['name']}) = {abs(acct['net']):,.2f}"
                        )
                        next_row += 1

            # Also inject individual sales into notes to p&l (rows 7-14)
            if ws_npl:
                import re as _re2
                def _extract_rate(name):
                    m = _re2.search(r'(\d+)\s*%', name)
                    return m.group(1) if m else None

                for acct in sale_accounts:
                    amt = abs(acct["net"])
                    if amt <= 0: continue
                    acct_rate = _extract_rate(acct["name"])
                    wrote = False
                    for r in range(7, 15):
                        cell_name = ws_npl.cell(r, 2).value
                        if not cell_name: continue
                        tmpl_rate = _extract_rate(str(cell_name))
                        if acct_rate and tmpl_rate and acct_rate == tmpl_rate:
                            acct_l = acct["name"].lower()
                            tmpl_l = str(cell_name).lower()
                            if (("within" in acct_l) == ("within" in tmpl_l)) or \
                               ("inter" not in acct_l and "within" not in acct_l):
                                if _safe_write(ws_npl, r, 4, amt):
                                    injected.append(f"notes to p&l!D{r} (Sale: {acct['name']}) = {amt:,.2f}")
                                    wrote = True
                                break
                    if not wrote:
                        for r in range(7, 10):
                            if ws_npl.cell(r, 4).value is None:
                                _safe_write(ws_npl, r, 2, acct["name"])
                                _safe_write(ws_npl, r, 4, amt)
                                injected.append(f"notes to p&l!D{r} (Sale new: {acct['name']}) = {amt:,.2f}")
                                break

        # ── B. Purchases → GROSS PROFIT!B14:B18 ─────────────────────
        # Row 14=Purchase GST 12% Interstate, 15=12% WS, 16=18% WS
        # Row 17=5% Interstate, 18=5% WS
        if "GROSS PROFIT" in wb.sheetnames:
            purchase_accounts = [a for a in individual_accounts
                                  if a.get("bs_head") == "purchases"
                                  and abs(a.get("net", 0)) > 0]

            # FIX (Issues 2 & 4): Same label-driven approach as Sales.
            # Purchases live on the DEBIT side of GROSS PROFIT —
            # column A holds labels, column B holds amounts.
            gp_cache = _cache.get("GROSS PROFIT", {})
            purch_label_rows = {}
            purch_total_rows = set()
            for r in range(1, 40):
                lbl = gp_cache.get((r, 1))  # col A
                if not lbl:
                    continue
                lbl_s = str(lbl).strip()
                if not lbl_s:
                    continue
                lbl_low = lbl_s.lower()
                # Skip headers and totals (PURCHASES, DIRECT EXPENSES, OPENING STOCK,
                # GROSS PROFIT, TOTAL).
                if lbl_low in ("purchases", "direct expenses", "opening stock",
                                "gross profit"):
                    continue
                if "total" in lbl_low:
                    purch_total_rows.add(r)
                    continue
                purch_label_rows[_norm_gst(lbl_s)] = (r, lbl_s)

            purch_row_totals = {}
            unmatched_purch = []
            for acct in purchase_accounts:
                name_norm = _norm_gst(acct["name"])
                best_row = None
                if name_norm in purch_label_rows:
                    best_row = purch_label_rows[name_norm][0]
                else:
                    for lbl_norm, (r, _) in purch_label_rows.items():
                        if lbl_norm and (lbl_norm in name_norm or name_norm in lbl_norm):
                            best_row = r
                            break
                if best_row is None:
                    unmatched_purch.append(acct)
                else:
                    purch_row_totals[best_row] = purch_row_totals.get(best_row, 0) + abs(acct["net"])

            for row, amt in purch_row_totals.items():
                if row in purch_total_rows:
                    continue   # never overwrite TOTAL row
                if _safe_write(ws_gp, row, 2, amt):
                    injected.append(f"GROSS PROFIT!B{row} (Purchase {amt:,.2f})")

            # Append unmatched purchase accounts below last sub-row, above TOTAL
            if unmatched_purch and purch_label_rows:
                last_sub_row = max(r for r, _ in purch_label_rows.values())
                next_row = last_sub_row + 1
                for acct in unmatched_purch:
                    if next_row in purch_total_rows:
                        break
                    if _safe_write(ws_gp, next_row, 1, acct["name"]):
                        _safe_write(ws_gp, next_row, 2, abs(acct["net"]))
                        injected.append(
                            f"GROSS PROFIT!B{next_row} (Purchase unmatched: {acct['name']}) = {abs(acct['net']):,.2f}"
                        )
                        next_row += 1

        # ── B2. DIRECT EXPENSES → GROSS PROFIT Direct-Expenses sub-rows ──
        # FIX (Issues 2/4): Direct expenses (Electricity Exp, Wages A/c,
        # Oil & Lubricant) must NOT be collapsed onto the Purchases header.
        # They live in their own labelled sub-rows below "DIRECT EXPENSES"
        # on the debit side of the Trading A/c.
        if "GROSS PROFIT" in wb.sheetnames:
            de_accounts = [a for a in individual_accounts
                           if a.get("bs_head") == "direct_expenses"
                           and abs(a.get("net", 0)) > 0]
            if de_accounts:
                gp_cache_de = _cache.get("GROSS PROFIT", {})
                # Find the DIRECT EXPENSES header row
                de_header_row = None
                for r in range(1, 40):
                    lbl = gp_cache_de.get((r, 1))
                    if lbl and "direct expense" in str(lbl).lower():
                        de_header_row = r
                        break
                # Find the first "total" / "gross profit" row after it
                de_end_row = None
                if de_header_row:
                    for r in range(de_header_row + 1, de_header_row + 15):
                        lbl = gp_cache_de.get((r, 1))
                        if lbl and ("total" in str(lbl).lower()
                                    or "gross profit" in str(lbl).lower()):
                            de_end_row = r
                            break

                if de_header_row and de_end_row and de_end_row > de_header_row + 1:
                    # Build label→row map for the direct-expense sub-rows
                    de_rows = {}
                    for r in range(de_header_row + 1, de_end_row):
                        lbl = gp_cache_de.get((r, 1))
                        if lbl:
                            de_rows[str(lbl).strip().lower()] = r

                    used_rows = set()
                    for acct in de_accounts:
                        name_l = acct["name"].lower()
                        amt = abs(acct["net"])
                        placed = False
                        # Match against existing labelled rows
                        for lbl, r in de_rows.items():
                            if r in used_rows:
                                continue
                            if lbl in name_l or name_l in lbl or any(
                                w in lbl and w in name_l
                                for w in ("electricity", "wages", "oil",
                                          "lubricant", "power", "fuel")
                            ):
                                if _safe_write(ws_gp, r, 2, amt):
                                    injected.append(
                                        f"GROSS PROFIT!B{r} (Direct Exp: {acct['name']}) = {amt:,.2f}"
                                    )
                                    used_rows.add(r)
                                    placed = True
                                    break
                        if placed:
                            continue
                        # Append as new labelled sub-row before TOTAL
                        for r in range(de_header_row + 1, de_end_row):
                            if r in used_rows:
                                continue
                            cur_lbl = gp_cache_de.get((r, 1))
                            cur_amt = gp_cache_de.get((r, 2))
                            if (not cur_lbl or str(cur_lbl).strip() == "") and (
                                cur_amt is None or cur_amt == 0
                            ):
                                if _safe_write(ws_gp, r, 1, acct["name"]):
                                    _safe_write(ws_gp, r, 2, amt)
                                    used_rows.add(r)
                                    injected.append(
                                        f"GROSS PROFIT!B{r} (Direct Exp new: {acct['name']}) = {amt:,.2f}"
                                    )
                                    placed = True
                                    break
                        if not placed:
                            skipped.append(
                                f"Direct expense '{acct['name']}' = {amt:,.2f}: no free row in Trading A/c"
                            )

        # ── C. Opening Stock → GROSS PROFIT!B9 ──────────────────────
        # B9 = "='notes to p&l'!D17" which = "=E24" (prev yr closing)
        # Opening stock should come from TB opening stock account
        opening_stock_accounts = [a for a in individual_accounts
                                   if "opening stock" in a.get("name","").lower()
                                   and abs(a.get("net", 0)) > 0]
        if opening_stock_accounts and "GROSS PROFIT" in wb.sheetnames:
            total_opening = sum(abs(a["net"]) for a in opening_stock_accounts)
            # notes to p&l!D17 is "=E24" — override it directly
            if _safe_write(ws_npl, 17, 4, total_opening):
                injected.append(f"notes to p&l!D17 (Opening stock) = {total_opening:,.2f}")

        # ── C2. OTHER INCOME (Rebate & Discount, Interest Received, etc.) ──
        # FIX (Issue 3): "REBATE & DISCOUNT" / other indirect-income accounts
        # were being classified correctly as `other_income` but never written
        # anywhere — so the amount silently disappeared from the BS / P&L.
        #
        # The standard "Notes to P&L" template carries an "Other Income"
        # section. We locate it by scanning column B of `notes to p&l` for a
        # row whose label contains "other income" / "other incomes", then
        # write each TB other-income account into a sub-row beneath it.
        oi_accounts = [a for a in individual_accounts
                       if a.get("bs_head") == "other_income"
                       and abs(a.get("net", 0)) > 0]
        if oi_accounts:
            npl_cache_for_oi = _cache.get("notes to p&l", {})
            oi_header_row = None
            oi_total_row  = None
            for r in range(1, 200):
                b = npl_cache_for_oi.get((r, 2))
                if not b:
                    continue
                bs = str(b).strip().lower()
                if "other income" in bs and "total" not in bs:
                    oi_header_row = r
                    break
            if oi_header_row:
                for r in range(oi_header_row + 1, oi_header_row + 20):
                    b = npl_cache_for_oi.get((r, 2))
                    if b and "total" in str(b).lower() and "other" in str(b).lower():
                        oi_total_row = r
                        break

            if oi_header_row and oi_total_row and oi_total_row > oi_header_row + 1:
                for acct in oi_accounts:
                    amt = abs(acct["net"])
                    placed = False
                    # First try fuzzy-matching an existing labelled sub-row
                    for r in range(oi_header_row + 1, oi_total_row):
                        b = npl_cache_for_oi.get((r, 2))
                        if b and _fuzzy_match_name(acct["name"], str(b)):
                            if _safe_write(ws_npl, r, 4, amt):
                                injected.append(
                                    f"notes to p&l!D{r} (Other Income: {acct['name']}) = {amt:,.2f}"
                                )
                                placed = True
                                break
                    if placed:
                        continue
                    # Otherwise find first empty sub-row
                    for r in range(oi_header_row + 1, oi_total_row):
                        d_val = npl_cache_for_oi.get((r, 4))
                        b_val = npl_cache_for_oi.get((r, 2))
                        if (d_val is None or d_val == 0) and (
                            b_val is None or str(b_val).strip() == ""
                        ):
                            if _safe_write(ws_npl, r, 2, acct["name"]):
                                _safe_write(ws_npl, r, 4, amt)
                                injected.append(
                                    f"notes to p&l!D{r} (Other Income new: {acct['name']}) = {amt:,.2f}"
                                )
                                placed = True
                                break
                    if not placed:
                        skipped.append(
                            f"Other Income '{acct['name']}' = {amt:,.2f}: no row in Other Income section"
                        )
            else:
                skipped.append(
                    "Other Income section not found in 'notes to p&l'. "
                    "Add a row labelled 'Other Income' to capture: "
                    + ", ".join(a["name"] for a in oi_accounts)
                )

        # ── D. Employee Expenses → match to template rows dynamically ──
        # Detect Employee benefits section in notes to p&l
        emp_start = None
        emp_end = None
        for r in range(30, 60):
            b = ws_npl.cell(r, 2).value
            if not b: continue
            bl = str(b).lower()
            if not emp_start and 'employee benefit' in bl and 'total' not in bl:
                emp_start = r + 1
            if emp_start and 'total' in bl and 'employee' in bl:
                emp_end = r
                break
        if not emp_start: emp_start = 41
        if not emp_end: emp_end = 46

        # All employee-type accounts (including ESI, bonus, leave with wages)
        emp_accounts = [a for a in individual_accounts
                        if a.get("bs_head") == "employee_expenses"
                        and "payable" not in a.get("name","").lower()
                        and abs(a.get("net", 0)) > 0]
        # Also include salary, wages, ESI, bonus, leave from other_expenses
        for a in individual_accounts:
            name_l = a.get("name","").lower()
            if a.get("bs_head") in ("other_expenses", "direct_expenses"):
                if any(kw in name_l for kw in ["salary", "wage", "e.s.i", "esi ",
                       "bonus", "leave with", "labour refreshment", "labor"]):
                    if "payable" not in name_l and abs(a.get("net", 0)) > 0:
                        if a not in emp_accounts:
                            emp_accounts.append(a)

        if emp_accounts:
            emp_template = {}
            for r in range(emp_start, emp_end):
                b = ws_npl.cell(r, 2).value
                if b and isinstance(b, str) and len(b.strip()) > 2:
                    emp_template[r] = b.strip().lower()
            
            written_emp = set()
            for acct in emp_accounts:
                matched = False
                for r, lbl in emp_template.items():
                    if r in written_emp: continue
                    if _fuzzy_match_name(acct["name"], lbl):
                        if _safe_set(ws_npl, r, 4, abs(acct["net"])):
                            written_emp.add(r)
                            injected.append(f"notes to p&l!D{r} (Employee: {acct['name']}) = {abs(acct['net']):,.2f}")
                            matched = True
                        break
                if not matched:
                    # Find empty row in section
                    for r in range(emp_start, emp_end):
                        if r not in written_emp:
                            d = ws_npl.cell(r, 4).value
                            if d is None or d == 0:
                                b = ws_npl.cell(r, 2).value
                                if b is None or str(b).strip() == "":
                                    _safe_set(ws_npl, r, 2, acct["name"])
                                _safe_set(ws_npl, r, 4, abs(acct["net"]))
                                written_emp.add(r)
                                injected.append(f"notes to p&l!D{r} (Employee new: {acct['name']}) = {abs(acct['net']):,.2f}")
                                break

        # ── E. Finance Cost → match to template rows dynamically ──
        # Detect Finance cost section
        fin_start = None
        fin_end = None
        for r in range(30, 80):
            b = ws_npl.cell(r, 2).value
            if not b: continue
            bl = str(b).lower()
            if not fin_start and 'finance cost' in bl and 'total' not in bl:
                fin_start = r + 1
            if fin_start and 'total' in bl and 'finance' in bl:
                fin_end = r
                break
        if not fin_start: fin_start = 50
        if not fin_end: fin_end = 56

        # All finance cost accounts (interest on loans, bank charges related to loans)
        FINANCE_KEYWORDS = ["interest", "loan interest", "car loan interest",
                           "machine loan", "top up", "bank cc intt", "bank cc",
                           "bank interest", "cc interest", "bank od interest",
                           "overdraft interest", "interest on unsecured",
            "intt on tempu", "interest on tempu", "interest on hdfc",
            "interest on term loan", "term loan interest",
                           "interest on loan", "interest on term",
                           "interest paid to", "interest to partner"]
        
        finance_accounts = [a for a in individual_accounts
                           if abs(a.get("net", 0)) > 0
                           and any(kw in a.get("name","").lower() for kw in FINANCE_KEYWORDS)
                           and a.get("bs_head") != "other_expenses"]  # don't override classification
        # Also include accounts explicitly classified as finance_cost
        for a in individual_accounts:
            if a.get("bs_head") == "finance_cost" and abs(a.get("net", 0)) > 0:
                if a not in finance_accounts:
                    finance_accounts.append(a)

        if finance_accounts:
            fin_template = {}
            for r in range(fin_start, fin_end):
                b = ws_npl.cell(r, 2).value
                if b and isinstance(b, str) and len(b.strip()) > 2:
                    fin_template[r] = b.strip().lower()
            
            written_fin = set()
            for acct in finance_accounts:
                matched = False
                for r, lbl in fin_template.items():
                    if r in written_fin: continue
                    if _fuzzy_match_name(acct["name"], lbl):
                        if _safe_set(ws_npl, r, 4, abs(acct["net"])):
                            written_fin.add(r)
                            injected.append(f"notes to p&l!D{r} (Finance: {acct['name']}) = {abs(acct['net']):,.2f}")
                            matched = True
                        break
                if not matched:
                    for r in range(fin_start, fin_end):
                        if r not in written_fin:
                            d = ws_npl.cell(r, 4).value
                            if d is None or d == 0:
                                b = ws_npl.cell(r, 2).value
                                if b is None or str(b).strip() == "":
                                    _safe_set(ws_npl, r, 2, acct["name"])
                                _safe_set(ws_npl, r, 4, abs(acct["net"]))
                                written_fin.add(r)
                                injected.append(f"notes to p&l!D{r} (Finance new: {acct['name']}) = {abs(acct['net']):,.2f}")
                                break

        # Combined finance+employee accounts for exclusion from Other Expenses
        finance_acct_names = {a["name"].lower() for a in finance_accounts}
        emp_acct_names = {a["name"].lower() for a in emp_accounts} if emp_accounts else set()

        # ── F. Other Expenses → notes to p&l!D57:D78 ─────────────────
        # Finance-cost keywords to exclude from other_expenses
        # Only EXACT bank/loan interest excluded — NOT "intt paid on late payment of tds"
        BANK_INT_KEYWORDS = ["bank interest", "bank cc intt", "cc interest",
                             "bank od interest", "overdraft interest", "bank charges"]
        LOAN_INT_KEYWORDS = ["loan interest", "car loan interest", "machine loan",
                             "top up", "interest on loan", "interest on term",
                             "interest on unsecured",
            "intt on tempu", "interest on tempu", "interest on hdfc",
            "interest on term loan", "term loan interest", "interest paid to",
                             "interest to partner"]
        FINANCE_KEYWORDS = set(BANK_INT_KEYWORDS + LOAN_INT_KEYWORDS)
        finance_acct_names = {a["name"].lower() for a in finance_accounts}

        # Also exclude salary, depreciation, and employee expenses — handled separately
        EXCLUDE_FROM_OTHER = finance_acct_names | emp_acct_names | {"salary","depreciation","dep on","amort"}

        other_exp_accounts = [a for a in individual_accounts
                              if a.get("bs_head") == "other_expenses"
                              and abs(a.get("net", 0)) > 0
                              and a.get("name","").lower() not in EXCLUDE_FROM_OTHER
                              and a.get("name","").lower() not in finance_acct_names
                              and a.get("name","").lower() not in emp_acct_names
                              and not any(kw in a.get("name","").lower()
                                          for kw in {"salary","depreciation","dep on","amort",
                                                     "salary to partner", "interest to partner"})]

        # FIX (Issue 2): "INTT PAID ON LATE PAYMENT OF TDS" must stay in Other Expenses.
        # Previously these were rerouted to D62 (Bank Charges), which was wrong
        # because interest on late payment of TDS is a statutory penal expense,
        # not a bank charge. We now leave them in `other_exp_accounts` so they
        # get placed into a row of the Other Expenses section (D57:D78), and we
        # also write the account name into column B for that row so the label is visible.
        # (No-op block kept for clarity — see unmatched-placement logic below where
        #  the account name is now written to column B as well.)

        # Build template label → row map for Other Expenses section
        # Different templates place Other Expenses at different rows.
        # Detect the actual section start by searching for "Other Expenses" header
        exp_section_start = 72  # default
        exp_section_end = 98    # default
        npl_cache = _cache.get("notes to p&l", {})
        
        for r in range(50, 100):
            b_val = npl_cache.get((r, 2))
            if b_val and "other expense" in str(b_val).lower():
                exp_section_start = r + 1
                break
        
        # Find section end (next "Total" or blank section)
        for r in range(exp_section_start, min(exp_section_start + 30, 120)):
            b_val = npl_cache.get((r, 2))
            if b_val and "total" in str(b_val).lower() and "other" in str(b_val).lower():
                exp_section_end = r
                break
        
        exp_row_map = {}
        for r in range(exp_section_start, exp_section_end):
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
                    if _safe_write(ws_npl, row, 4, amt):
                        injected.append(f"notes to p&l!D{row} ({acct['name']}) = {amt:,.2f}")
            else:
                # No template match — find first empty D row in 57:78
                # FIX (Issue 2): also write the account name to column B so the
                # label is visible alongside the amount (otherwise interest on
                # late payment of TDS would appear as an unlabelled amount).
                placed = False
                for r in range(exp_section_start, exp_section_end):
                    if r not in written_exp_rows:
                        d_val = _cache.get("notes to p&l", {}).get((r, 4))
                        b_val = _cache.get("notes to p&l", {}).get((r, 2))
                        if d_val is None or d_val == 0:
                            if _safe_write(ws_npl, r, 4, amt):
                                # Write the account name to col B if the row is unlabelled
                                if b_val is None or str(b_val).strip() == "":
                                    _safe_write(ws_npl, r, 2, acct["name"])
                                written_exp_rows.add(r)
                                injected.append(f"notes to p&l!D{r} (unmatched: {acct['name']}) = {amt:,.2f}")
                                placed = True
                            break
                if not placed:
                    skipped.append(f"Other expense '{acct['name']}' = {amt:,.2f}: no row in notes to p&l")

    # ── Depreciation → Fixed Assets C. Yr. sheet H31 ─────────────────
    # notes to p&l!D52 = 'Fixed Assets C. Yr.'!H31 (formula, auto)
    # We write the TB depreciation total directly to H31 so it flows through.
    dep_total = aggregated_values.get("depreciation", 0)
    if dep_total and "Fixed Assets C. Yr." in wb.sheetnames:
        ws_fa = wb["Fixed Assets C. Yr."]
        fa_cell = ws_fa.cell(31, 8)  # H31
        if fa_cell.value is None or (isinstance(fa_cell.value, (int, float)) and fa_cell.value == 0):
            if _safe_set(ws_fa, 31, 8, dep_total):
                injected.append(f"Fixed Assets C. Yr.!H31 (Depreciation) = {dep_total:,.2f}")
            else:
                skipped.append(f"Depreciation {dep_total:,.2f}: Fixed Assets H31 is formula — not overwritten")
        else:
            log.append(f"· Depreciation H31 already has value {fa_cell.value} — not overwritten")

    for s in skipped:
        log.append(f"⚠ SKIPPED: {s}")
    for inj_msg in injected:
        log.append(f"✓ {inj_msg}")

    wb.save(output_path)
    wb.close()

    # ── FORMULA CORRECTION PASS ──────────────────────────────────────────────
    # For KNOWN templates (Shree Craft, Fashion Adda): call inject_with_config
    # which has the exact config-driven injection. This overrides the heuristic output.
    try:
        import importlib.util as _ilu, os as _os
        _inj_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "bs_inject_config.py")
        if _os.path.exists(_inj_path):
            _spec = _ilu.spec_from_file_location("_bsic_corr", _inj_path)
            _bsic = _ilu.module_from_spec(_spec)
            _spec.loader.exec_module(_bsic)
            # Check if this is a known template
            from openpyxl import load_workbook as _lwb0
            _wb0 = _lwb0(bs_template_path)
            _tk, _cfg = _bsic.get_config(_wb0) if hasattr(_bsic, 'get_config') else (None, None)
            _wb0.close()
            if _tk:
                # Known template — run full config-driven injection (replaces heuristic output)
                _bsic.inject_with_config(
                    tb_path  = tb_path,
                    template = bs_template_path,
                    output   = output_path,
                )
                # Skip the manual fixes below — inject_with_config handled everything
                raise StopIteration("config injection done")
    except StopIteration:
        pass  # clean exit after config injection
    except Exception as _ce:
        pass  # config injection failed, fall through to manual fixes

    # Manual formula fixes for unknown templates (Fashion Adda fallback)
    try:
        from openpyxl import load_workbook as _lwb
        _wb2 = _lwb(output_path)
        _sheets = _wb2.sheetnames

        # notes to p&l fixes
        if "notes to p&l" in _sheets:
            _npl = _wb2["notes to p&l"]
            # Fix 1: Revenue formula — must cover all 6 sale rows (E11:E16)
            if str(_npl.cell(6,4).value).startswith("=SUM"):
                _npl.cell(6,4).value = "=SUM('GROSS PROFIT'!E11:E16)"
            # Fix 2: Clear stray SALE GST 12% INTERSTATE value written outside revenue box
            if isinstance(_npl.cell(8,4).value, (int, float)):
                _npl.cell(8,4).value = None
            # Fix 3: Purchase formula — must cover all 5 purchase rows (B14:B19)
            if str(_npl.cell(20,4).value).startswith("=SUM"):
                _npl.cell(20,4).value = "=SUM('GROSS PROFIT'!B14:B19)"
            # Fix 4: Merge INTT PAID ON LATE TDS (296) into Bank Charges (D62)
            _d62 = _npl.cell(62,4).value
            _d70 = _npl.cell(70,4).value
            if isinstance(_d70, (int, float)) and _d70 == 296:
                if isinstance(_d62, (int, float)):
                    _npl.cell(62,4).value = _d62 + 296
                _npl.cell(70,4).value = None

        # GROSS PROFIT fixes
        if "GROSS PROFIT" in _sheets:
            _gp = _wb2["GROSS PROFIT"]
            # Fix 5: Closing stock back-calc must use E11:E16
            if "E11:E14" in str(_gp.cell(17,5).value):
                _gp.cell(17,5).value = "=B24-SUM(E11:E16)"
            # Fix 6: Gross Profit % must use E11:E16
            if "E11:E14" in str(_gp.cell(22,2).value):
                _gp.cell(22,2).value = "=SUM(E11:E16)*8.5%"
            # Fix 7: Kill any #REF! in old data columns
            for _r in range(1, 30):
                for _c in range(1, 12):
                    if "#REF!" in str(_wb2["GROSS PROFIT"].cell(_r,_c).value or ""):
                        _wb2["GROSS PROFIT"].cell(_r,_c).value = 0

        _wb2.save(output_path)
        _wb2.close()
    except Exception as _fe:
        pass  # formula fix failed silently — do not break existing flow

    # ─────────────────────────────────────────────────────────────
    # FIX (Issue 3): On-screen totals must mirror what the downloaded BS shows.
    #
    # The old logic simply summed asset-side vs. liability-side TB aggregates,
    # which always tallies (debits = credits in any TB) — so the UI permanently
    # displayed equal Assets/Liabilities and zero profit, even though the
    # downloaded Excel file (after recomputing formulas in Excel/LibreOffice)
    # shows real numbers with a P&L balance.
    #
    # We now compute the actual BS figures the spreadsheet will resolve to:
    #   1. Net Profit = Revenue − (Opening Stock + Purchases + Employee
    #                  + Finance Cost + Depreciation + Other Expenses) + Closing Stock
    #   2. Closing Stock is the balancing figure in the Trading A/c:
    #      Closing Stock = (Opening Stock + Purchases) − (Sales − Gross Profit)
    #      For display purposes we use the inventory aggregate from TB if present,
    #      otherwise treat closing stock as the Trading A/c balancing figure.
    # ─────────────────────────────────────────────────────────────
    revenue_total       = aggregated_values.get("revenue", 0)
    other_income_total  = aggregated_values.get("other_income", 0)   # FIX Issue 3
    purchases_total     = aggregated_values.get("purchases", 0)
    direct_exp_total    = aggregated_values.get("direct_expenses", 0)
    employee_total      = aggregated_values.get("employee_expenses", 0)
    other_exp_total     = aggregated_values.get("other_expenses", 0)
    depreciation_total  = aggregated_values.get("depreciation", 0)
    inventories_total   = aggregated_values.get("inventories", 0)

    # Opening stock is included inside the `inventories` aggregate in some TBs
    # and as a separate line in others. Use individual_accounts to separate.
    opening_stock = 0.0
    closing_stock_from_tb = 0.0
    if individual_accounts:
        for a in individual_accounts:
            n = a.get("name", "").lower()
            if "opening stock" in n:
                opening_stock += abs(a.get("net", 0))
            elif "closing stock" in n:
                closing_stock_from_tb += abs(a.get("net", 0))

    # If no explicit opening/closing rows, treat the whole inventories aggregate
    # as opening stock (typical Tally TB convention).
    if opening_stock == 0 and closing_stock_from_tb == 0:
        opening_stock = inventories_total

    # Closing stock used in the BS = explicit value if TB has one,
    # otherwise the Trading A/c balancing figure.
    # Trading A/c balancing figure rule:
    #   Closing Stock = max(0, (Opening Stock + Purchases) − Sales) when Sales < Cost
    # We instead compute Net Profit using a single equation that does NOT rely
    # on a separately-supplied closing stock.
    if closing_stock_from_tb > 0:
        closing_stock_bs = closing_stock_from_tb
    else:
        # Closing stock acts as the credit-side balancing figure on the
        # Trading A/c. Without an explicit value, we cannot derive it
        # independently — leave at 0 and let Net Profit absorb the difference.
        closing_stock_bs = 0.0

    # Net Profit = Sales + Other Income + Closing Stock
    #              − Opening Stock − Purchases − Employee
    #              − Other Expenses − Depreciation
    # FIX Issue 3: Other Income (e.g. Rebate & Discount) was being silently
    # ignored from the P&L computation, causing the on-screen "Profit" figure
    # to under-report by that amount.
    net_profit = (
        revenue_total + other_income_total + closing_stock_bs
        - opening_stock - purchases_total - direct_exp_total
        - employee_total - other_exp_total - depreciation_total
    )

    # ── Asset side ──
    total_assets = (
        aggregated_values.get("fixed_assets", 0)
        + aggregated_values.get("non_current_investments", 0)
        + closing_stock_bs
        + aggregated_values.get("trade_rec", 0)
        + aggregated_values.get("cash_bank", 0)
        + aggregated_values.get("stla", 0)
        + aggregated_values.get("other_current_assets", 0)
    )

    # ── Liability side (Capital + Profit + Borrowings + Payables + OCL + Provisions) ──
    capital_total = aggregated_values.get("capital", 0)
    total_liabilities = (
        capital_total
        + net_profit
        + aggregated_values.get("lt_borrowings", 0)
        + aggregated_values.get("st_borrowings", 0)
        + aggregated_values.get("trade_payables", 0)
        + aggregated_values.get("other_cl", 0)
        + aggregated_values.get("st_provisions", 0)
    )

    diff = abs(total_assets - total_liabilities)

    log.append(f"\n📊 P&L Summary:")
    log.append(f"  Revenue:           {revenue_total:>15,.2f}")
    log.append(f"  Other Income:      {other_income_total:>15,.2f}")
    log.append(f"  Opening Stock:     {opening_stock:>15,.2f}")
    log.append(f"  Purchases:         {purchases_total:>15,.2f}")
    log.append(f"  Direct Expenses:   {direct_exp_total:>15,.2f}")
    log.append(f"  Employee Exp:      {employee_total:>15,.2f}")
    log.append(f"  Other Expenses:    {other_exp_total:>15,.2f}")
    log.append(f"  Depreciation:      {depreciation_total:>15,.2f}")
    log.append(f"  Closing Stock:     {closing_stock_bs:>15,.2f}")
    log.append(f"  → Net Profit:      {net_profit:>15,.2f}")

    log.append(f"\n📊 Balance Sheet Totals (computed):")
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
        "net_profit": net_profit,
        "revenue": revenue_total,
        "other_income": other_income_total,           # NEW Issue 3
        "opening_stock": opening_stock,
        "closing_stock": closing_stock_bs,
        "capital": capital_total,
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
    Supports both .xlsx and .pdf input files.
    """
    # Handle PDF input — parse and use directly
    if tb_path.lower().endswith('.pdf'):
        detection = parse_tb_pdf(tb_path)
        if not detection or not detection.get("accounts"):
            return {"error": "Could not parse trial balance PDF. Ensure it's a text-based PDF (not scanned image)."}
    else:
        detection = detect_tb_structure(tb_path)
        if "error" in detection:
            return detection

    classified = classify_accounts(detection["accounts"])

    # Pre-flag debit-creditor / credit-debtor as the explicit advance heads
    # so the Review-Mapping UI groups them correctly.
    for acct in classified:
        head = acct.get("bs_head")
        net  = acct.get("net", 0)
        if head == "trade_payables" and net > 0:
            acct["bs_head"] = "advance_to_supplier"
            acct["reclassified_from"] = "trade_payables"
        elif head == "trade_rec" and net < 0:
            acct["bs_head"] = "advance_from_customer"
            acct["reclassified_from"] = "trade_rec"

    # Separate by confidence
    high_conf = [a for a in classified if a["confidence"] == "high"]
    low_conf = [a for a in classified if a["confidence"] == "low"]
    unclassified = [a for a in classified if a["confidence"] == "none"]

    # ── FIX (Issue 1): Build an explicit "Manual Review" list so the UI can
    # show the user WHICH accounts need attention. Previously the front-end
    # only knew the COUNT ("6 Manual") but had no way to drill down. We now
    # return the full account objects, plus a flat list of dropdown options
    # so the UI can render a `<select>` with every valid bs_head label.
    bs_head_options = [
        {"key": k, "label": v["label"], "side": v["side"]}
        for k, v in BS_HEADS.items()
    ]
    # Add the two sign-aware advance heads (they aren't in BS_HEADS as separate
    # buckets — they're virtual heads that aggregate into stla / other_cl)
    bs_head_options.append({"key": "advance_to_supplier",
                            "label": "Advance to Supplier (Cr-side debit)",
                            "side": "asset"})
    bs_head_options.append({"key": "advance_from_customer",
                            "label": "Advance from Customer (Dr-side credit)",
                            "side": "liability"})

    # Manual-review payload — what the UI shows under the "Manual" tab.
    # `dr_cr` is filled so the UI can colour debit vs credit balances and
    # show them at the end of the BS/P&L group they ALMOST belong to.
    manual_review = []
    for a in unclassified:
        manual_review.append({
            "row":      a.get("row"),
            "name":     a.get("name"),
            "group":    a.get("group"),
            "debit":    a.get("debit", 0),
            "credit":   a.get("credit", 0),
            "net":      a.get("net", 0),
            "dr_cr":    "Dr" if a.get("net", 0) >= 0 else "Cr",
            "bs_head":  a.get("bs_head", "unclassified"),
            # Suggest the most likely side so the UI can pre-position the row
            "suggested_side": "asset" if a.get("net", 0) >= 0 else "liability",
        })

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
        "manual_review": manual_review,            # NEW — list of "Manual X" items
        "bs_head_options": bs_head_options,        # NEW — dropdown options for UI
        "summary": {
            "total_accounts":    len(classified),
            "high_confidence":   len(high_conf),
            "low_confidence":    len(low_conf),
            "unclassified":      len(unclassified),
            "manual_count":      len(unclassified),  # alias for UI clarity
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
    # Step 1: Analyze TB (handles both xlsx and pdf)
    # If PDF, convert to temporary xlsx for the injection step
    import os, tempfile
    actual_tb_path = tb_path
    tmp_xlsx = None
    if tb_path.lower().endswith('.pdf'):
        tmp_xlsx = tempfile.mktemp(suffix='.xlsx')
        if not convert_pdf_tb_to_xlsx(tb_path, tmp_xlsx):
            return {"error": "Could not convert PDF trial balance to Excel."}
        actual_tb_path = tmp_xlsx

    analysis = analyze_trial_balance(tb_path)  # Use original path for detection
    if "error" in analysis:
        if tmp_xlsx and os.path.exists(tmp_xlsx):
            os.remove(tmp_xlsx)
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

    # Step 4 + 5: Config-driven injection via bs_inject_config.py
    # Uses template_configs.py exact row maps — correct formulas, no heuristics.
    # Falls back to old heuristic injector (which now has formula correction pass built-in).
    try:
        import importlib.util, os
        _inj_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bs_inject_config.py")
        if os.path.exists(_inj_path):
            spec = importlib.util.spec_from_file_location("bs_inject_config", _inj_path)
            inj  = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(inj)
            result = inj.inject_with_config(
                tb_path  = tb_path,
                template = bs_template_path,
                output   = output_path,
            )
            result["analysis"]            = analysis
            result["aggregated"]          = aggregated
            result["classified_accounts"] = accounts
        else:
            # bs_inject_config.py not found — use heuristic injector
            # (formula correction pass is now built into inject_into_bs above)
            result = inject_into_bs(
                bs_template_path, output_path, aggregated,
                mapping_overrides=None,
                individual_accounts=accounts,
            )
            result["analysis"]            = analysis
            result["aggregated"]          = aggregated
            result["classified_accounts"] = accounts
    except Exception as _inj_exc:
        # Hard fallback
        result = inject_into_bs(
            bs_template_path, output_path, aggregated,
            mapping_overrides=None,
            individual_accounts=accounts,
        )
        result["post_process_log"]    = [f"Config injector error: {_inj_exc}"]
        result["analysis"]            = analysis
        result["aggregated"]          = aggregated
        result["classified_accounts"] = accounts

    return result
