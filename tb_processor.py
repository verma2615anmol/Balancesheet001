"""
Trial Balance → Balance Sheet Processor (Adaptive Scanner Edition)
- Dynamic layout scanning replaces hardcoded rows (Risk 1).
- Writes to 'notes to p&l' or 'GROSS PROFIT' via dual-route detector (Risk 2).
- Hard guards prevent writes below total rows (Risk 3).
- Memory-efficient: single workbook open, read_only where possible.
"""
import re
import os
import shutil
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.utils import get_column_letter
from copy import copy
import pdfplumber

# ═══════════════════════════════════════════════════════════════════════
# ACCOUNT CLASSIFICATION RULES
# ═══════════════════════════════════════════════════════════════════════
BS_HEADS = {
# --- LIABILITIES ---
 "capital ": {
 "label ":  "Owner's Capital / Partners Capital ",
 "side ":  "liability ",
 "keywords ": [
 "capital ",  "partner ",  "proprietor ",  "owner ",  "equity ",
 "share capital ",  "reserves ",  "surplus ",  "retained earning ",
 "general reserve ",  "capital reserve ",  "securities premium ",
 "profit  & loss ",  "profit and loss ",  "p &l appropriation ",
 "current account ",  "drawing ",  "partner current ",
 "capital account ",  "share application ",
],
 "negative_keywords ": [ "capital gain ",  "capital goods ",  "working capital loan "],
},
 "lt_borrowings ": {
 "label ":  "Long Term Borrowings ",
 "side ":  "liability ",
 "keywords ": [
 "long term loan ",  "term loan ",  "secured loan ",  "unsecured loan ",
 "mortgage ",  "debenture ",  "long term borrowing ",
 "loan from bank ",  "loan from director ",  "loan from partner ",
 "vehicle loan ",  "car loan ",  "home loan ",  "housing loan ",
 "hypothecation ",  "loan payable ",
],
 "negative_keywords ": [ "short term ",  "od ",  "overdraft ",  "cc limit "],
},
 "st_borrowings ": {
 "label ":  "Short Term Borrowings ",
 "side ":  "liability ",
 "keywords ": [
 "short term loan ",  "overdraft ",  "od account ",  "cash credit ",
 "cc limit ",  "cc account ",  "working capital ",  "packing credit ",
 "bank od ",  "bank overdraft ",  "short term borrowing ",
],
 "negative_keywords ": [],
},
 "trade_payables ": {
 "label ":  "Trade Payables ",
 "side ":  "liability ",
 "keywords ": [
 "trade payable ",  "creditor ",  "sundry creditor ",
 "accounts payable ",  "supplier ",  "purchase payable ",
 "bills payable ",  "trade creditor ",
],
 "negative_keywords ": [],
},
 "advance_from_customer ": {
 "label ":  "Advance from Customer ",
 "side ":  "liability ",
 "keywords ": [
 "advance from customer ",  "customer advance ",
 "advance received ",  "advance from buyer ",
 "advance from m/s ",  "advance from mr ",  "advance from ms ",
 "advance from  ",
],
 "negative_keywords ": [
 "advance from supplier ",  "advance from vendor ",
],
},
 "other_cl ": {
 "label ":  "Other Current Liabilities ",
 "side ":  "liability ",
 "keywords ": [
 "other current liabilit ",  "statutory ",  "tds payable ",
 "gst payable ",  "gst output ",  "output cgst ",  "output sgst ",
 "output igst ",  "tax payable ",  "duty payable ",  "cess payable ",
 "salary payable ",  "wages payable ",  "rent payable ",
 "interest payable ",  "expense payable ",  "outstanding ",
 "advance from customer ",  "customer advance ",  "security deposit received ",
 "audit fee payable ",  "professional fee payable ",
 "electricity payable ",  "telephone payable ",
 "provision for expense ",  "payable ",
 "income received in advance ",  "unearned ",
],
 "negative_keywords ": [ "provision for tax ",  "provision for depreciation ",  "provision for bad debt "],
},
 "st_provisions ": {
 "label ":  "Short Term Provisions ",
 "side ":  "liability ",
 "keywords ": [
 "provision for tax ",  "provision for income tax ",
 "provision for depreciation ",  "provision for bad debt ",
 "provision for doubtful ",  "short term provision ",
 "provision for gratuity ",  "provision for bonus ",
 "provision for leave ",  "provision for warranty ",
],
 "negative_keywords ": [],
},
# --- ASSETS ---
 "fixed_assets ": {
     "label ":  "Fixed Assets / PPE ",
     "side ":  "asset ",
     "keywords ": [
         "fixed asset ",  "property ",  "plant ",  "equipment ",  "ppe ",
         "land ",  "building ",  "furniture ",  "fixture ",  "vehicle ",
         "computer ",  "machinery ",  "office equipment ",  "electrical ",
         "air condition ",  "ac  ",  "motor car ",  "scooter ",  "bike ",
         "mobile ",  "telephone instrument ",  "printer ",  "laptop ",
         "intangible ",  "goodwill ",  "patent ",  "trademark ",  "copyright ",
         "software ",  "leasehold ",  "capital wip ",  "cwip ",
    ],
     "negative_keywords ": [
         "depreciation ",  "accumulated ",  "provision for ",
         "repair ",  "maintenance ",  "rent ",
    ],
},
 "non_current_investments ": {
     "label ":  "Non-Current Investments ",
     "side ":  "asset ",
     "keywords ": [
         "investment ",  "shares ",  "debenture held ",  "mutual fund ",
         "fixed deposit ",  "fdr ",  "fd  ",  "nsc  ",  "kvp ",
         "government securities ",  "bonds ",
         "investment in subsidiary ",  "investment in associate ",
    ],
     "negative_keywords ": [ "provision for investment "],
},
 "inventories ": {
     "label ":  "Inventories / Stock ",
     "side ":  "asset ",
     "keywords ": [
         "inventor ",  "stock ",  "closing stock ",  "opening stock ",
         "raw material ",  "finished good ",  "work in progress ",
         "wip ",  "stores ",  "spare ",  "packing material ",
         "stock in trade ",  "goods in transit ",
    ],
     "negative_keywords ": [ "stock broker "],
},
 "trade_rec ": {
     "label ":  "Trade Receivables ",
     "side ":  "asset ",
     "keywords ": [
         "trade receivable ",  "debtor ",  "sundry debtor ",
         "accounts receivable ",  "bills receivable ",
         "trade debtor ",  "receivable from customer ",
    ],
     "negative_keywords ": [],
},
 "cash_bank ": {
     "label ":  "Cash and Bank Balances ",
     "side ":  "asset ",
     "keywords ": [
         "cash ",  "bank ",  "cash in hand ",  "cash at bank ",
         "petty cash ",  "savings account ",  "current account bank ",
         "bank account ",  "bank balance ",  "cheque in hand ",
         "imprest ",
    ],
     "negative_keywords ": [ "cash credit ",  "cc account ",  "overdraft ",  "od account ",  "bank od ",
                           "interest ",  "loan interest ",  "loan a/c ",  "loan account ",
                           "bank charge ",  "processing fee "],
},
 "stla ": {
     "label ":  "Short Term Loans  & Advances ",
     "side ":  "asset ",
     "keywords ": [
         "loan given ",  "advance to ",  "loan to ",
         "advance to staff ",  "staff advance ",
         "prepaid ",  "deposit ",  "security deposit paid ",
         "tds receivable ",  "tcs receivable ",  "input tax ",
         "input cgst ",  "input sgst ",  "input igst ",  "gst input ",
         "advance tax ",  "self assessment tax ",  "mat credit ",
         "cenvat ",  "vat input ",  "excise input ",
         "income tax refund ",  "refund receivable ",
         "advance recoverable ",
    ],
     "negative_keywords ": [ "advance from customer ",  "customer advance ",
                           "advance to supplier ",  "advance to customer ",
                           "supplier advance "],
},
 "advance_to_supplier ": {
     "label ":  "Advance to Supplier ",
     "side ":  "asset ",
     "keywords ": [
         "advance to supplier ",  "advance to customer ",  "supplier advance ",
         "advance paid to supplier ",  "advance to vendor ",  "vendor advance ",
    ],
     "negative_keywords ": [ "advance from customer ",  "customer advance "],
},
 "other_current_assets ": {
     "label ":  "Other Current Assets ",
     "side ":  "asset ",
     "keywords ": [
         "other current asset ",  "accrued income ",  "interest accrued ",
         "accrued interest ",  "interest receivable ",
         "other receivable ",  "other asset ",
    ],
     "negative_keywords ": [],
},

# --- P &L HEADS ---
 "revenue ": {
     "label ":  "Revenue from Operations / Sales ",
     "side ":  "pl ",
     "keywords ": [
         "sale ",  "revenue ",  "income from operation ",
         "turnover ",  "gross receipt ",  "service income ",
         "service revenue ",  "consulting income ",  "fee received ",
         "commission received ",  "commission income ",
         "export sale ",  "domestic sale ",  "local sale ",
    ],
     "negative_keywords ": [ "sale return ",  "sales return ",  "sale of asset ",  "sale of investment "],
},
 "other_income ": {
     "label ":  "Other Income ",
     "side ":  "pl ",
     "keywords ": [
         "other income ",  "interest received ",  "interest income ",
         "dividend received ",  "dividend income ",  "rental income ",
         "rent received ",  "miscellaneous income ",  "misc income ",
         "profit on sale ",  "gain on sale ",  "exchange gain ",
         "discount received ",  "discount earned ",  "commission earned ",
         "bad debt recovered ",  "insurance claim received ",
         "interest on fd ",  "interest on deposit ",  "bank interest received ",
    ],
     "negative_keywords ": [],
},
 "purchases ": {
     "label ":  "Purchases / Cost of Material ",
     "side ":  "pl ",
     "keywords ": [
         "purchase ",  "cost of material ",  "cost of goods ",
         "import purchase ",  "local purchase ",  "domestic purchase ",
         "raw material consumed ",  "material consumed ",
         "sub contract ",  "job work ",  "labour charge ",
         "freight inward ",  "carriage inward ",  "octroi ",
         "custom duty ",  "clearing charge ",
    ],
     "negative_keywords ": [ "purchase return "],
},
 "employee_expenses ": {
     "label ":  "Employee / Salary Expenses ",
     "side ":  "pl ",
     "keywords ": [
         "salary ",  "wage ",  "bonus ",  "gratuity ",  "leave ",
         "staff welfare ",  "epf ",  "esi ",  "pf contribution ",
         "employee benefit ",  "director remuneration ",
         "partner salary ",  "partner remuneration ",  "salary to partner ",
         "stipend ",  "incentive ",  "overtime ",
         "labour refreshment ",  "labor refreshment ",
         "e.s.i ",  "leave with wages ",
    ],
     "negative_keywords ": [ "salary payable ",  "wages payable ",  "bonus payable ",  "leave with wages payable ",
                           "e.s.i payable ",  "esi payable "],
},
 "finance_cost ": {
     "label ":  "Finance Cost ",
     "side ":  "pl ",
     "keywords ": [
         "interest paid ",  "interest on loan ",  "bank charge ",
         "bank interest ",  "bank cc intt ",  "cc interest ",
         "bank od interest ",  "overdraft interest ",
         "loan interest ",  "loan 1 interest ",  "loan 2 interest ",
         "loan 3 interest ",  "loan 4 interest ",
         "car loan interest ",  "machine loan interest ",  "machinery loan interest ",
         "top up loan interest ",  "top up car loan interest ",
         "interest to partner ",  "interest on unsecured ",
         "interest on term loan ",  "interest on secured ",
         "processing fee ",  "cersai charge ",
         "life insurance machinery loan ",
    ],
     "negative_keywords ": [ "interest received ",  "interest income ",  "intt paid on late payment "],
},
 "other_expenses ": {
     "label ":  "Other Expenses / Indirect Expenses ",
     "side ":  "pl ",
     "keywords ": [
         "expense ",  "rent ",  "electricity ",  "telephone ",
         "internet ",  "travelling ",  "conveyance ",  "vehicle running ",
         "petrol ",  "diesel ",  "fuel ",  "repair ",  "maintenance ",
         "insurance ",  "audit fee ",  "professional fee ",  "legal fee ",
         "printing ",  "stationery ",  "postage ",  "courier ",
         "advertisement ",  "marketing ",  "donation ",
         "discount allowed ",  "bad debt ",
         "miscellaneous ",  "office expense ",  "general expense ",
         "entertainment ",  "subscription ",  "membership ",
         "rate ",  "tax ",  "municipal ",  "water charge ",
         "loading ",  "unloading ",  "packing ",
         "commission paid ",  "brokerage ",  "agency ",
         "foreign exchange loss ",  "exchange diff ",
         "penalty ",  "fine ",  "late fee ",  "round off ",
         "festival ",  "gift ",  "welfare ",
    ],
     "negative_keywords ": [
         "salary ",  "wage ",  "depreciation ",  "purchase ",
         "payable ",  "outstanding ",  "provision ",
    ],
},
 "depreciation ": {
     "label ":  "Depreciation ",
     "side ":  "pl ",
     "keywords ": [
         "depreciation ",  "amortisation ",  "amortization ",
         "dep on ",  "accumulated depreciation ",
    ],
     "negative_keywords ": [ "provision for depreciation "],
},
 "direct_expenses ": {
     "label ":  "Direct Expenses ",
     "side ":  "pl ",
     "keywords ": [
         "wages a/c ",  "wages account ",  "factory wages ",  "labour wages ",
         "electricity exp ",  "power and fuel ",  "power  & fuel ",
         "oil  & lubricant ",  "oil and lubricant ",
         "factory expense ",  "production expense ",
    ],
     "negative_keywords ": [ "wages payable ",  "electricity payable "],
},
}

CLASSIFICATION_PRIORITY = [
 "depreciation ",
 "advance_from_customer ",  "advance_to_supplier ",
 "trade_payables ",  "trade_rec ",
 "finance_cost ",   # Must be before cash_bank so  "LOAN INTEREST " doesn't match  "bank "
 "employee_expenses ",
 "cash_bank ",  "inventories ",
 "st_provisions ",
 "st_borrowings ",  "lt_borrowings ",
 "direct_expenses ",  "purchases ",  "revenue ",  "other_income ",
 "fixed_assets ",  "non_current_investments ",
 "stla ",
 "capital ",
 "other_cl ",  "other_current_assets ",
 "other_expenses ",
]

# ═══════════════════════════════════════════════════════════════════════
# PDF TRIAL BALANCE PARSER
# ═══════════════════════════════════════════════════════════════════════
def parse_tb_pdf(pdf_path):
    """Parse a Trial Balance PDF (Tally/Busy/Excel-exported) into account rows."""
    import pdfplumber, re

    num_re = re.compile(r'([\d,]+\.\d{2})')
    GROUP_KEYWORDS = {
         "bank accounts ",  "bank account ", "capital account ",  "capital accounts ",
         "cash-in-hand ",  "cash in hand ", "fixed assets ",
         "direct expenses ",  "indirect expenses ", "indirect incomes ",  "indirect income ",
         "direct incomes ",  "direct income ", "purchase account ",  "purchase accounts ",  "purchases ",
         "sales account ",  "sales accounts ",  "sales ",
         "sundry creditors ",  "sundry debtors ", "sundry payables ",  "sundry payable ",
         "sundry receivables ",  "sundry receivable ",
         "loans  & advances (asset) ",  "loans  & advances ", "loans and advances ",  "loans (liability) ",  "loans liability ",
         "current liabilities ",  "current assets ",
         "deposits (asset) ",  "deposits ", "duties  & taxes ",  "duties and taxes ",
         "secured loans ",  "unsecured loans ",  "unsecure loans ",
         "stock-in-hand ",  "stock in hand ", "investments ",
         "profit  & loss account ",  "profit and loss account ", "profit  & loss a/c ",
         "provisions ",  "reserves  & surplus ",  "reserves and surplus ",
         "misc. expenses (asset) ",  "miscellaneous expenses ",
         "branch/divisions ",  "suspense a/c ",
    }

    def _norm_header(line):
        return re.sub(r "[\s\-:]+$ ",  " ", line.strip().lower())

    def _is_group_header(line):
        return _norm_header(line) in GROUP_KEYWORDS

    def _to_float(v):
        if v is None: return 0.0
        s = str(v).strip()
        if  not s: return 0.0
        s = (s.replace( ", ",  " ").replace( "₹ ",  " ")
               .replace( "Rs. ",  " ").replace( "Rs ",  " ")
               .replace( "( ",  "-").replace( ") ",  " "))
        try:
            return float(s)
        except ValueError:
            return 0.0

    SKIP_NAMES = { "particulars ",  "trial balance ",
                   "debit amount ",  "credit amount ",
                   "total ",  "grand total ",  "opening balance ",
                   "closing balance ",  "balance c/d ",  "balance b/d ",
                   "sub total ",  "net total "}

    accounts = []
    running_group =  " "

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text_lines  = [ln.strip() for ln in
                              (page.extract_text() or  " ").splitlines()
                              if ln.strip()]

                tables = page.extract_tables() or []
                 table_rows = []
                for tbl in tables:
                    for row in tbl:
                        if not row: continue
                        if all( (c is None or str(c).strip() ==  " ")
                               for c in row): continue
                        table_rows.append(row)

                tr_first_cells = {str(row[0]).strip()
                                  for row in table_rows
                                  if row and row[0]}

                page_group = running_group
                line_to_group = {}
                for ln in text_lines:
                    if _is_group_header(ln):
                        page_g roup = _norm_header(ln)
                        continue
                    name_part = re.sub(
                        r "\s+-?[\d,]+\.\d{1,2}(\s+-?[\d,]+\.\d{1,2})?\s*$ ",
                         " ", ln).strip()
                    if name_part and name_part in tr_first_cells:
                        line_to_group.setdefault(name_part, page_group)

                for row in table_rows:
                    cells = [str(c).strip() if c is not None else  " "
                             for c in row]
                    while len(cells)  < 3: cells.append( " ")
                    name, dr_str, cr_str = cells[0], cells[1], cells[2]
                    if not name: continue
                    
                    nl = name.lower().strip()
                    if nl in SKIP_NAMES: continue
                    if re.search(r "continued on page ", name, re.I): continue
                    if re.match(r "^(total|grand total|sub total|opening|closing|balance)\b ", name, re.I): continue

                    dr = _to_float(dr_str)
                    cr = _to_float(cr_str)
                    if dr == 0 and c r == 0: continue

                    grp = line_to_group.get(name, page_group)
                    accounts.append({
                         "row ":    len(accounts),
                         "key ":    f "{name}_{len(accounts)} ",
                         "name ":   name,
                         "group ":  grp,
                         "debit ":  dr,
                         "credit ": cr,
                         "net ":    dr - cr,
                    })

                running_group = page_group
    except Exception:
        try: return _parse_tb_pdf_text_fallback(pdf_path)
        except Exception: return None

    if not accounts: 
        try: return _parse_tb_pdf_text_fallback(pdf_path)
        except Exception: return None

    return {
         "format_type ":    "PDF ",
         "sheet_name ":     "PDF ",
         "header_row ":    0,
         "data_start_row ":0,
         "account_col ":   0,
         "debit_col ":     1,
         "credit_col ":    2,
         "net_col ":       None,
         "accounts ":      accounts,
    }

def _parse_tb_pdf_text_fallback(pdf_path):
    """Legacy text-only PDF parser kept as a fallback."""
    import pdfplumber, re
    all_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or  " "
            for line in text.split( "\n "):
                l = line.strip()
                if l: all_lines.append(l)
    if not all_lines: return None

    num_re = re.compile(r'([\d,]+\.\d{2})')
    skip_patterns = { "trial balance ",  "as on  ",  "page no ",  "continued ",
                      "focal point ",  "punjab ",  "phase-",  "e-254 "}
    credit_groups = { "capital account ",  "secured loans ",  "unsecured loans ",
                      "sundry creditors ",  "sundry payables ",  "sales account ",
                      "indirect incomes ",  "profit  & loss account ",
                      "current liabilities ",  "duties  & taxes "}

    current_group =  " "
    accounts = []
    company_name =  " "

    for line in all_lines:
        ll = line.lower().strip()
        if any(s in ll for s in skip_patterns): continue
        if ll ==  "particulars debit amount credit amount ": continue
        if not company_name and line.isupper() and len(line)  > 3 and not num_re.search(line):
            company_name = line; continue
        if company_name and line.strip() == company_name: continue
        nums = num_re.findall(line) 
        name = num_re.sub('', line).strip()
        if not nums:
            if name and len(name)  > 1 and name not in ( "0.01 ",  " "):
                nl = name.lower()
                if not any(s in nl for s in [ "phase-",  "focal ",  "punjab ",
                            "ludhiana-",  "delhi-",  "mumbai-",  "address "]):
                    current_group = name
            continue
        if not name: continue
        dr_amt = 0.0
        cr_amt = 0.0
        if len(nums) == 1:
            val = float(nums[0].re place(',', ''))
            if current_group.lower() in credit_groups: cr_amt = val
            else: dr_amt = val
        elif len(nums)  >= 2:
            dr_amt = float(nums[0].replace(',', ''))
            cr_amt = float(nums[1].replace(',', ''))
        accounts.append({
             "row ":    len(accounts),
             "key ":    f "{name}_{len(accounts)} ",
             "name ":   name,
             "group ":  current_group,
             "debit ":  dr_amt,
             "credit ": cr_amt,
             "net ":    dr_amt - cr_amt,
        })
    if not accounts: return None
    return {
         "format_type ":    "PDF ",
         "sheet_name ":     "PDF ",
         "header_row ":    0,
         "data_start_row ":0,
         "account_col ":   0,
         "debit_col ":     1,
         "credit_col ":    2,
         "net_col ":       None,
         "accounts ":      accounts,
    }

def convert_pdf_tb_to_xlsx(pdf_path, xlsx_path):
    """Convert a PDF trial balance to an xlsx file for processing."""
    result = parse_tb_pdf(pdf_path)
    if not result or not result["accounts"]: return False
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title =  "Trial Balance "
    ws.cell(1, 1,  "Particulars ")
    ws.cell(1, 2,  "Debit ")
    ws.cell(1, 3,  "Credit ")
    current_group =  " "
    r = 2
    for acct in result[ "accounts "]:
        if acct[ "group "] != current_group:
            current_group = acct[ "group "]
            if current_group:
                ws.cell(r, 1, current_group)
                r += 1
        ws.cell(r, 1, acct[ "name "])
        if acct[ "debit "]: ws.cell(r, 2, acct[ "debit "])
        if acct[ "credit "]: ws.cell(r, 3, acct[ "credit "])
        r += 1
    wb.save(xlsx_path)
    return True

# ═══════════════════════════════════════════════════════════════════════
# TB AUTO-DETECTION ENGINE
# ═══════════════════════════════════════════════════════════════════════
def detect_tb_structure(file_path):
    wb = load_workbook(file_path, read_only=True, data_only=True)
    results = []
    for sname in wb.sheetnames:
        ws = wb[sname]
        rows_data = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            rows_data.append(list(row))
            if i >= 2000: break
        if len(rows_data) < 2: continue
        det = _detect_columns(rows_data, sname)
        if det: results.append(det)
    wb.close()
    if not results:
        return {"error": "Could not detect trial balance structure. Please check the file format."}
    best = max(results, key=lambda d: len(d.get("accounts", [])))
    return best

def _detect_columns(rows, sheet_name):
    header_row = None
    acct_col = None
    dr_col = None
    cr_col = None
    net_col = None
    format_type = None
    
    for ri, row in enumerate(rows[:15]):
        if not row: continue
        row_lower = [str(c).strip().lower() if c else  " " for c in row]

        for ci, val in enumerate(row_lower):
            if not val: continue
            if  any(k in val for k in [ "particular ",  "account ",  "ledger ",  "name ",  "head ", "description ",  "party name ", ]):
                if acct_col is None: acct_col = ci; header_row = ri
            if (val in ( "dr ",  "debit ",  "dr balance ",  "debit balance ", "dr amount ",  "debit amount ",  "dr bal ",  "debit bal ") or val.startswith( "debit ") or val.startswith( "dr  ") or val ==  "dr. " or  "debit " in val.split( "( ")[0].strip()):
                dr_col = ci
            if (val in ( "cr ",  "credit ",  "cr balance ",  "credit balance ", "cr amount ",  "credit amount ",  "cr bal ",  "credit bal ") or val.startswith( "credit ") or val.startswith( "cr  ") or val ==  "cr. " or  "credit " in val.split( "( ")[0].strip()):
                cr_col = ci
            if val in ( "amount ",  "net balance ",  "net amount ",  "balance ", "closing balance ",  "closing ",  "net "): net_col = ci

    if acct_col is None:
        for  ri, row in enumerate(rows[:15]):
            if not row: continue
            text_cols = []; num_cols = []
            for ci, val in enumerate(row):
                if val i s None: continue
                if isinstance(val, str) and len(val.strip())  > 2 and not _is_number_str(val): text_cols.append(ci)
                elif isinstance(val, (int, float)) or _is_number_str(str(val)): num_cols.append(ci )
            if len(text_cols)  >= 1 and len(num_cols)  >= 1: acct_col = text_cols[0]; header_row = ri; break

    if acct_col is None: return None

    if dr_col is not None and cr_col is not None and dr_col != cr_col: format_type = 1
    elif net_col is not None: format_typ e = 4
    else:
        for ri, row in enumerate(rows[header_row + 1: header_row + 10] , header_row + 1):
            if not row: continue
            num_cols_found = []
            for ci, val in enumerate(row):
                if ci == acct_col: contin ue
                if isinstance(val, (int, float)) and val != 0: num_cols_found.append(ci)
                elif isinstance(val, str) and _is_number_str(val): num_cols_found.append(ci)
            if len(num_cols_found)  >= 2: dr_col = num_cols_found[0]; cr_col = num_cols_found[1]; format_type = 1; break
            elif len(num_cols_found) == 1: net_col = num_cols_found[0]; format_type = 4; break

    if format_type is None: return None
    data_start = header_row + 1
    accounts = []
    total_ke ywords = { "total ",  "grand total ",  "difference ",  "net total ", "closing balance ",  "opening balance total ", "balance c/d ",  "balance b/d "}
    current_group = None

    for ri in range(data_start, len(rows)):
        row = rows[ri]
        if not row or ri  >= len(rows): continue
        acct_name = row[acct_col] if acct_col  < len(row) else None
        if not acct_name or not isinstance(acct_name, str): continue
        acct_name = acct_name.strip()
        if not acct_name or len(acct_name)  < 2: continue
        if acct_name.lower().strip() in total_keywords: continue
        if re.match(r'^(total|grand total|sub total|net total)\b ', acct_name, re.I): continue

        dr_amt = 0; cr_amt = 0; net_amt = 0
        if format_type == 1 and dr_col is not None and cr_col is not None:
            dr_amt = _to_float(row[dr_col] if dr_col  < len(row) else None)
            cr_amt = _to_float(row[cr_col] if cr_col  < len(row) else None)
            net_amt = dr_amt - cr_amt
        elif format_type == 4 and net_col is not None:
            if dr_col is None and cr_col is None:
                dr_try = row[1] if len(row)  > 1 else None
                cr_try = row[2] if len(row)  > 2 else None
                dr_f = _to_float(dr_try); cr_f = _to_float(cr_try)
                if dr_f != 0 or cr_f != 0: dr_amt = dr_f; cr_amt  = cr_f; net_amt = dr_amt - cr_amt
                else: net_val = row[net_col] if net_col  < len(row) else None; net_amt = _to_float(net_val)
            else:
                net_val = row[net_col] if net_col  < len(row) else None
                net_amt = _to_float(net_val)
            if net_amt  > 0: dr_amt = net_amt
            else: cr_amt = abs(net_amt)

        if dr_amt == 0 and cr_amt == 0 and net_amt == 0: current_group = acct_name; conti nue
        accounts.append({
             "row ": ri,
             "key ": f "{acct_name}_{ri} ",
             "name ": acct_name,
             "group ": current_group,
             "debit ": dr_amt,
             "credit ": cr_amt,
             "net ": net_amt,
        })

    return {
         "format_type ": format_type,
         "sheet_name ": sheet_name,
         "header_row ": header_row,
         "data_start_row ": data_start,
         "account_col ": acct_col,
         "debit_col ": dr_col,
         "credit_col ": cr_col,
         "net_col ": net_col,
         "accounts ": accounts,
    }

def _is_number_str(s):
    if not s: return False
    s = s.strip().replace( ", ",  " ").replace( "( ",  "-").replace( ") ",  " ")
    try: float(s); return True
    except (ValueError, TypeError): return False

def _to_float(val):
    if val is None: return 0.0
    if isinstance(val, (int, float)): return float(val)
    if isinstance(val, str):
        s = val.strip().replace( ", ",  " ").replace( "₹ ",  " ").replace( "Rs. ",  " ").replace( "Rs ",  " ")
        if s.startswith( "( ") and s.endswith( ") "): s =  "-" + s[1:-1]
        try: return float(s)
        except (ValueError, TypeError): return 0.0
    return 0.0

# ═══════════════════════════════════════════════════════════════════════
# ACCOUNT CLASSIFIER
# ═══════════════════════════════════════════════════════════════════════
GROUP_HEAD_MAP = {
 "capital account ":            "capital ",
 "capital accounts ":           "capital ",
 "reserves  & surplus ":         "capital ",
 "reserves and surplus ":       "capital ",
 "bank accounts ":              "cash_bank ",
 "bank account ":               "cash_bank ",
 "cash-in-hand ":               "cash_bank ",
 "cash in hand ":               "cash_bank ",
 "fixed assets ":               "fixed_assets ",
 "investments ":                "non_current_investments ",
 "sundry creditors ":           "trade_payables ",
 "sundry debtors ":             "trade_rec ",
 "sundry debtor ":              "trade_rec ",
 "sundry receivables ":         "trade_rec ",
 "sundry receivable ":          "trade_rec ",
 "purchase account ":           "purchases ",
 "purchase accounts ":          "purchases ",
 "purchases ":                  "purchases ",
 "sales account ":              "revenue ",
 "sales accounts ":             "revenue ",
 "sales ":                      "revenue ",
 "stock-in-hand ":              "inventories ",
 "stock in hand ":              "inventories ",
 "direct expenses ":            "direct_expenses ",
 "indirect income ":            "other_income ",
 "indirect incomes ":           "other_income ",
 "direct income ":              "revenue ",
 "direct incomes ":             "revenue ",
 "sundry payables ":            "other_cl ",
 "sundry payable ":             "other_cl ",
 "current liabilities ":        "other_cl ",
 "current assets ":             "other_current_assets ",
 "provisions ":                 "st_provisions ",
 "unsecure loans ":             "lt_borrowings ",
 "unsecured loans ":            "lt_borrowings ",
 "secured loans ":              "lt_borrowings ",
 "loans (liability) ":          "lt_borrowings ",
 "loans liability ":            "lt_borrowings ",
 "loans  & advances (asset) ":   "stla ",
 "loans  & advances ":           "stla ",
 "loans and advances ":         "stla ",
 "deposits (asset) ":           "stla ",
 "deposits ":                   "stla ",
 "duties  & taxes ":             "stla ",
 "duties and taxes ":           "stla ",
 "misc. expenses (asset) ":     "misc_expenditure ",
 "miscellaneous expenses ":     "misc_expenditure ",
 "profit  & loss account ":      "capital ",
 "profit and loss account ":    "capital ",
 "profit  & loss a/c ":          "capital ",
}

def classify_accounts(accounts):
    classified = []
    for acct in accounts:
        name = acct[ "name "]
        group = acct.get( "group ",  " ")
        bs_head, confidence = _classify_single(name, acct[ "net "], group)
        acct_copy = dict(acct)
        acct_copy[ "bs_head "] = bs_head
        acct_copy[ "confidence "] = confidence
        classified.append(acct_copy)
    return classified

def _classify_single(name, net_amount, group=None):
    name_lower = name.lower().strip()
    group_lower = (group or "").lower().strip()
    if  "loan " in name_lower and net_amount  < 0:
        if any(kw in name_lower for kw in [ "secured ",  "hypothec ",  "mortgage "]): return  "lt_borrowings ",  "high "
        if any(kw in name_lower for kw in [ "unsecure ",  "unsecured "]): return  "lt_borrowings ",  "high "
        return  "lt_borrowings ",  "high "
    if ( "bank " in name_lower or  "a/c " in name_lower) and net_amount  < 0:
        if any(kw in name_lower for kw in [ "loan ",  "od ",  "cc ",  "overdraft ", "cash credit ",  "machinery ",  "vehicle ",  "term loan "]): return  "lt_borrowings ",  "high "
        if any(kw in group_lower for kw in [ "bank ",  "cash "]): return  "st_borrowings ",  "high "
    if  "round off " in name_lower or  "roundoff " in name_lower: return  "other_expenses ",  "high "
    if group_lower and group_lower in GROUP_HEAD_MAP: return GROUP_HEAD_MAP[group_lower],  "high "
    for head_key in CLASSIFICATION_PRIORITY:
        head = BS_HEADS[head_key]
        excluded = False
        for nk in head.get( "negative_keywords ", []):
            if nk in name_lower: excluded = True; break
        if excluded: continue
        for kw in head[ "keywords "]:
            if kw in name_lower: return head_key,  "high "
    if group_lower:
        for grp_key, head_key in GROUP_HEAD_MAP.items():
            if grp_key in group_lower or group_lower in grp_key: return  head_key,  "low "
    if net_amount  > 0: return  "other_current_assets ",  "low "
    elif net_amount  < 0: return  "other_cl ",  "low "
    else: return  "unclassified ",  "none "

def get_aggregated_values(classified_accounts):
    totals = {}
    for acct in classified_accounts:
        head = acct[ "bs_head "]
        net  = acct[ "net "]
        if head ==  "unclassified ": continue
        if head ==  "trade_payables " and net  > 0:
            head =  "stla "; acct[ "bs_head "] =  "advance_to_supplier "; acct[ "reclassified_from "] =  "trade_payables "
        elif head ==  "trade_rec " and net  < 0:
            head =  "other_cl "; acct[ "bs_head "] =  "advance_from_customer "; acct[ "reclassified_from "] =  "trade_rec "
        elif head ==  "st_provisions " and net  > 0:
            head =  "stla "; acct[ "bs_head "] =  "stla "; acct[ "reclassified_from "] =  "st_provisions "; acct[ "stla_subtype "] =  "revenue_authority "
        elif head ==  "advance_from_customer ": head =  "other_cl "; acct[ "reclassified_from "] = acct.get( "reclassified_from ") or  "trade_rec "
        elif head ==  "advance_to_supplier ": head =  "stla "; acct[ "reclassified_from "] = acct.get( "reclassified_from ") or  "trade_payables "
        amt = abs(net)
        if head not in totals: totals[head] = 0
        totals[head] += amt
    return totals

# ═══════════════════════════════════════════════════════════════════════
# SAFE CELL WRITERS
# ═══════════════════════════════════════════════════════════════════════
def _is_formula(val):
    return isinstance(val, str) and val.strip().startswith("=")

def _get_writable_cell(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell): return cell
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col):
            anchor = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
            if not isinstance(anchor, MergedCell): return anchor
    return None

def _safe_set(ws, row, col, value):
    try:
        cell = _get_writable_cell(ws, row, col)
        if cell is None: return False
        if _is_formula(cell.value): return False
        cell.value = round(float(value), 2)
        return True
    except (AttributeError, TypeError, ValueError): return False

def _safe_write(ws, row, col, value):
    try:
        cell = _get_writable_cell(ws, row, col)
        if cell is None: return False
        if _is_formula(cell.value): return False
        if isinstance(value, (int, float)): cell.value = round(float(value), 2)
        else: cell.value = value
        return True
    except (AttributeError, TypeError, ValueError): return False

def _fuzzy_match_name(tb_name, template_name):
    import re
    def normalize(s):
        s = s.lower()
        s = re.sub(r'\bm/s.?\s*', '', s)
        s = re.sub(r'[^a-z0-9\s]', ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
        for city in ['ludhiana', 'delhi', 'jalandhar', 'surat', 'ahmedabad', 'ahemadabad', 'varanasi', 'mumbai', 'ambala', 'citi']:
            s = re.sub(r'\b' + city + r'\b', '', s).strip()
        return s
    a = normalize(tb_name)
    b = normalize(template_name)
    if not a or not b: return False
    if a == b: return True
    COMMON_WORDS = {'textiles', 'trading', 'enterprises', 'creation', 'fashion', 'fabrics', 'industries', 'pvt', 'ltd', 'co', 'and', 'sons'}
    if len(a) >= 6 and len(b) >= 6:
        shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
        if shorter in longer:
            unique_words = [w for w in shorter.split() if w not in COMMON_WORDS and len(w) > 3]
            if unique_words: return True
    words_a = [w for w in a.split() if len(w) > 3 and w not in COMMON_WORDS]
    words_b = [w for w in b.split() if len(w) > 3 and w not in COMMON_WORDS]
    if words_a and words_b:
        common = sum(1 for w in words_a if w in words_b)
        if common >= min(2, len(words_a), len(words_b)): return True
    if words_a and words_b:
        long_a = [w for w in words_a if len(w) >= 7]
        long_b = [w for w in words_b if len(w) >= 7]
        if any(w in long_b for w in long_a): return True
    return False

# ═══════════════════════════════════════════════════════════════════════
# DYNAMIC SCANNER + DUAL-ROUTE + HARD-GUARD INJECTION ENGINE
# ═══════════════════════════════════════════════════════════════════════

def _scan_template_layout(wb):
    """RISK 1 FIX: Dynamically scans template sheets to find section boundaries."""
    layout = {}
    section_keywords = {
        "revenue": ["revenue from operations", "sale", "turnover"],
        "purchases": ["purchase", "cost of material consumed"],
        "opening_stock": ["opening stock"],
        "closing_stock": ["closing stock", "inventories"],
        "direct_expenses": ["direct expense", "factory exp", "wages", "electricity exp", "power & fuel"],
        "employee_expenses": ["employee benefit", "salary exp", "staff cost", "wage"],
        "finance_cost": ["finance cost", "interest exp", "bank charges", "interest paid"],
        "depreciation": ["depreciation"],
        "other_expenses": ["other expense", "indirect exp"],
        "lt_borrowings": ["long term borrowings", "secured loan"],
        "st_borrowings": ["short term borrowings", "od ", "cc ", "cash credit"],
        "trade_payables": ["trade payable", "sundry creditor"],
        "other_cl": ["other current liabilit"],
        "st_provisions": ["short term provision"],
        "cash_bank": ["cash and bank", "cash in hand", "bank balance"],
        "stla": ["short term loan", "advance", "deposit", "gst input", "tds"],
        "trade_rec": ["trade receivable", "sundry debtor"],
    }

    for sheet_name in wb.sheetnames:
        if sheet_name in ("GROSS PROFIT", "bs", "capital", "Fixed Assets C. Yr.", "Fixed Assets P. Yr."):
            continue
        ws = wb[sheet_name]
        current_sec = None
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=250, min_col=2, max_col=5, values_only=False), 1):
            b_val = str(row[0].value or "").strip().lower()
            if not b_val: continue

            for key, keywords in section_keywords.items():
                if any(k in b_val for k in keywords):
                    current_sec = key
                    layout[current_sec] = {"sheet": sheet_name, "start": r_idx + 1, "end": None, "total_row": None, "amount_col": 4, "targets": {}}
                    break
            
            if current_sec and ("total" in b_val or "grand total" in b_val):
                if current_sec in layout:
                    layout[current_sec]["total_row"] = r_idx
                    layout[current_sec]["end"] = r_idx - 1
                current_sec = None

            if current_sec and current_sec in layout:
                if "opening" in b_val: layout[current_sec]["targets"]["opening"] = r_idx
                elif "purchase" in b_val: layout[current_sec]["targets"]["purchase"] = r_idx
                elif "sale" in b_val: layout[current_sec]["targets"]["sale"] = r_idx

    for sec in layout:
        if not layout[sec]["end"]:
            layout[sec]["end"] = layout[sec]["start"] + 25
            layout[sec]["total_row"] = layout[sec]["end"] + 1
    return layout

def find_writable_pl_cell(wb, keyword, priority_sheets=("notes to p&l", "GROSS PROFIT")):
    """SMART DUAL-ROUTE DETECTOR: Finds the first plain-value cell matching a P&L keyword."""
    for sheet_name in priority_sheets:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=80, max_col=6, values_only=False):
            labels = [str(c.value or "").lower() for c in row if c.value]
            label_text = " ".join(labels)
            if keyword.lower() in label_text:
                for col_idx in [4, 5, 3, 2]:
                    cell = ws.cell(row=row[0].row, column=col_idx)
                    if not _is_formula(cell.value) and not isinstance(cell, MergedCell):
                        return sheet_name, cell.row, col_idx
    return None, None, None

def _safe_inject_range(ws, section_info, target_val, label, log, fallback_name=None):
    """RISK 3 HARD GUARD: Injects into section range without ever crossing total_row."""
    if not section_info or target_val == 0: return False
    start, end = section_info["start"], section_info["end"]
    col = section_info["amount_col"]
    targets = section_info.get("targets", {})

    for t_key, t_row in targets.items():
        if t_key in label.lower() and t_row <= end:
            if _safe_write(ws, t_row, col, target_val):
                log.append(f"✓ {ws.title}!{get_column_letter(col)}{t_row} ({label}) = {target_val:,.2f}")
                return True

    for r in range(start, end + 1):
        cell_val = ws.cell(r, 2).value
        if cell_val and _fuzzy_match_name(label, str(cell_val)):
            if _safe_write(ws, r, col, target_val):
                log.append(f"✓ {ws.title}!{get_column_letter(col)}{r} (matched: {label}) = {target_val:,.2f}")
                return True

    for r in range(start, end + 1):
        d_val = ws.cell(r, col).value
        b_val = ws.cell(r, 2).value
        if (d_val is None or d_val == 0) and (b_val is None or str(b_val).strip() == ""):
            if _safe_write(ws, r, col, target_val):
                if fallback_name and not b_val: _safe_write(ws, r, 2, fallback_name)
                log.append(f"✓ {ws.title}!{get_column_letter(col)}{r} (new: {label}) = {target_val:,.2f}")
                return True

    log.append(f"⚠ {ws.title}! Section '{label}' FULL or bounded. Skipped injection.")
    return False

def inject_into_bs(bs_template_path, output_path, aggregated_values, mapping_overrides=None, individual_accounts=None):
    """Main injection function using dynamic scanner and dual-route detection."""
    import shutil
    log = []
    shutil.copy2(bs_template_path, output_path)
    wb = load_workbook(output_path)
    layout = _scan_template_layout(wb)
    injected_count = 0

    # ── 1. P&L INJECTION (Dual-Route + Hard Guard) ─────────────────────
    # RISK 2 FIX: All P&L writes go to 'notes to p&l' or 'GROSS PROFIT' via detector.
    pl_items = [
        ("revenue", "sale"), ("other_income", "other income"),
        ("purchases", "purchase"), ("opening_stock", "opening stock"),
        ("closing_stock", "closing stock"), ("direct_expenses", "direct expense"),
        ("employee_expenses", "employee benefit"), ("finance_cost", "interest"),
        ("depreciation", "depreciation"), ("other_expenses", "other expense")
    ]

    for tb_key, keyword in pl_items:
        val = aggregated_values.get(tb_key, 0)
        if val == 0: continue
        
        sheet, row, col = find_writable_pl_cell(wb, keyword)
        if sheet and row:
            if _safe_write(wb[sheet], row, col, val):
                log.append(f"✓ P&L: {sheet}!{get_column_letter(col)}{row} ({tb_key}) = {val:,.2f}")
                continue
        
        sec = layout.get(tb_key)
        if sec and _safe_inject_range(wb[sec["sheet"]], sec, val, tb_key, log):
            continue
            
        log.append(f"⚠ P&L: Could not place {tb_key} ({val:,.2f}). No writable cell found.")

    # ── 2. BS NOTES INJECTION (Scanner + Hard Guard) ─────────────────────
    bs_items = [
        ("lt_borrowings", "lt_borrowings"), ("st_borrowings", "st_borrowings"),
        ("trade_payables", "trade_payables"), ("other_cl", "other_cl"),
        ("st_provisions", "st_provisions"), ("cash_bank", "cash_bank"),
        ("stla", "stla"), ("trade_rec", "trade_rec")
    ]

    for tb_key, sec_key in bs_items:
        val = aggregated_values.get(tb_key, 0)
        if val == 0: continue
        sec = layout.get(sec_key)
        if sec:
            _safe_inject_range(wb[sec["sheet"]], sec, val, tb_key, log)

    # ── 3. DETAILS SHEET (Creditors/Debtors) ─────────────────────────────
    if "Details" in wb.sheetnames and individual_accounts:
        ws_det = wb["Details"]
        det_layout = _scan_template_layout(wb)
        for tb_key, sec_key in [("trade_payables", "details_creditors"), ("trade_rec", "details_debtors")]:
            items = [a for a in individual_accounts if a.get("bs_head") == tb_key and abs(a.get("net",0))>0]
            sec = det_layout.get(sec_key)
            if sec and items:
                for acct in items:
                    amt = abs(acct["net"])
                    _safe_inject_range(ws_det, sec, amt, acct["name"], log, fallback_name=acct["name"])

    # ── 4. CAPITAL & FIXED ASSETS (Preserved) ─────────────────────────────
    for sk in ["capital", "fixed_assets"]:
        if aggregated_values.get(sk, 0) > 0:
            log.append(f"· {sk} from TB: {aggregated_values[sk]:,.2f} — skipped (fill from ledger/user)")

    # ── 5. COMPUTE & RETURN TALLY ────────────────────────────────────────
    wb.save(output_path)
    wb.close()
    
    rev = aggregated_values.get("revenue", 0)
    oth_inc = aggregated_values.get("other_income", 0)
    pur = aggregated_values.get("purchases", 0)
    dir_exp = aggregated_values.get("direct_expenses", 0)
    emp = aggregated_values.get("employee_expenses", 0)
    oth_exp = aggregated_values.get("other_expenses", 0)
    dep = aggregated_values.get("depreciation", 0)
    op_stk = sum(abs(a["net"]) for a in (individual_accounts or []) if "opening stock" in a.get("name","").lower()) or aggregated_values.get("inventories", 0)
    cl_stk = sum(abs(a["net"]) for a in (individual_accounts or []) if "closing stock" in a.get("name","").lower()) or 0
    
    net_profit = rev + oth_inc + cl_stk - op_stk - pur - dir_exp - emp - oth_exp - dep
    
    total_assets = (aggregated_values.get("fixed_assets", 0) + aggregated_values.get("non_current_investments", 0) +
                    cl_stk + aggregated_values.get("trade_rec", 0) + aggregated_values.get("cash_bank", 0) +
                    aggregated_values.get("stla", 0) + aggregated_values.get("other_current_assets", 0))
                    
    total_liab = (aggregated_values.get("capital", 0) + net_profit +
                  aggregated_values.get("lt_borrowings", 0) + aggregated_values.get("st_borrowings", 0) +
                  aggregated_values.get("trade_payables", 0) + aggregated_values.get("other_cl", 0) +
                  aggregated_values.get("st_provisions", 0))
                  
    diff = abs(total_assets - total_liab)
    log.extend([f"📊 Net Profit: {net_profit:,.2f}", f"📊 Assets: {total_assets:,.2f} | Liab: {total_liab:,.2f}", 
                f"{'✅ TALLIED' if diff < 1 else f'❌ Diff: {diff:,.2f}'}"])

    return {
        "status": "success", "output": output_path, "log": log,
        "injected_count": injected_count, "tally_ok": diff < 1,
        "total_assets": total_assets, "total_liabilities": total_liab, "net_profit": net_profit,
        "aggregated": aggregated_values, "revenue": rev, "other_income": oth_inc,
        "direct_expenses": dir_exp, "opening_stock": op_stk, "closing_stock": cl_stk,
        "purchases": pur, "employee_expenses": emp, "other_expenses": oth_exp,
        "depreciation": dep, "finance_cost": aggregated_values.get("finance_cost", 0)
    }

def _inject_cap_fa(output_path, cap_entries, fa_entries, log):
    from openpyxl import load_workbook
    from openpyxl.cell import MergedCell
    wb = load_workbook(output_path)
    def _is_formula(v): return isinstance(v, str) and v.startswith("=")
    def _safe_write(ws, row, col, value):
        try:
            cell = ws.cell(row, col)
            if isinstance(cell, MergedCell): return False
            if _is_formula(str(cell.value or " ")): return False
            cell.value = round(float(value), 2)
            return True
        except Exception: return False

    if cap_entries:
        cap_sheet = None
        for sn in wb.sheetnames:
            if "capital " in sn.lower(): cap_sheet = sn; break
        if cap_sheet:
            ws = wb[cap_sheet]
            for entry in cap_entries:
                row = entry.get("row")
                if not row: continue
                fields = [("introduced", "Capital Introduced"), ("interest_on_capital", "Interest on Capital"), ("salary", "Salary"), ("withdrawals", "Withdrawals")]
                for field_key, field_label in fields:
                    val = entry.get(field_key, 0)
                    col_idx = entry.get(f"{field_key}_col")
                    if val and col_idx:
                        if _safe_write(ws, row, col_idx, val):
                            log.append(f"✓ {cap_sheet}!{chr(64+col_idx)}{row} ({field_label}) = {float(val):,.2f}")

    if fa_entries:
        fa_sheet = None
        for sn in wb.sheetnames:
            sl = sn.lower()
            if "fixed asset " in sl or "fa  " in sl or sl.startswith("fa") or "ppe " in sl: fa_sheet = sn; break
        if fa_sheet:
            ws = wb[fa_sheet]
            for entry in fa_entries:
                row = entry.get("row")
                if not row: continue
                fields = [("additions_gt180", 3, "Addition  >180d"), ("additions_lt180", 4, "Addition  <180d"), ("sale", 5, "Sale")]
                for field_key, default_col, label in fields:
                    val = entry.get(field_key, 0)
                    col_idx = entry.get(f"{field_key}_col", default_col)
                    if val:
                        if _safe_write(ws, row, col_idx, val):
                            log.append(f"✓ {fa_sheet}!{chr(64+col_idx)}{row} ({label}) = {float(val):,.2f}")
    wb.save(output_path)
    wb.close()

# ═══════════════════════════════════════════════════════════════════════
# MAIN PIPELINE (called from app.py routes)
# ═══════════════════════════════════════════════════════════════════════
def detect_bs_template(file_path):
    wb = load_workbook(file_path, read_only=True, data_only=True)
    result = {"sheets": []}
    for sname in wb.sheetnames:
        ws = wb[sname]
        rows_data = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            rows_data.append(list(row))
            if i  >= 100: break
        result["sheets"].append({"name": sname})
    wb.close()
    if not result["sheets"]: return {"error": "Could not detect Balance Sheet template structure."}
    return result

def analyze_trial_balance(tb_path):
    if tb_path.lower().endswith('.pdf'):
        detection = parse_tb_pdf(tb_path)
        if not detection or not detection.get("accounts"): return {"error": "Could not parse trial balance PDF."}
    else:
        detection = detect_tb_structure(tb_path)
        if "error" in detection: return detection
    classified = classify_accounts(detection["accounts"])

    for acct in classified:
        head = acct.get("bs_head")
        net  = acct.get("net", 0)
        if head == "trade_payables" and net > 0:
            acct["bs_head"] = "advance_to_supplier"; acct["reclassified_from"] = "trade_payables"
        elif head == "trade_rec" and net < 0:
            acct["bs_head"] = "advance_from_customer"; acct["reclassified_from"] = "trade_rec"

    high_conf = [a for a in classified if a["confidence"] == "high"]
    low_conf = [a for a in classified if a["confidence"] == "low"]
    unclassified = [a for a in classified if a["confidence"] == "none"]

    bs_head_options = [{"key": k, "label": v["label"], "side": v["side"]} for k, v in BS_HEADS.items()]
    bs_head_options.append({"key": "advance_to_supplier", "label": "Advance to Supplier (Cr-side debit)", "side": "asset"})
    bs_head_options.append({"key": "advance_from_customer", "label": "Advance from Customer (Dr-side credit)", "side": "liability"})

    manual_review = []
    for a in unclassified:
        manual_review.append({"row": a.get("row"), "name": a.get("name"), "group": a.get("group"), "debit": a.get("debit", 0), "credit": a.get("credit", 0), "net": a.get("net", 0), "dr_cr": "Dr" if a.get("net", 0) >= 0 else "Cr", "bs_head": a.get("bs_head", "unclassified"), "suggested_side": "asset" if a.get("net", 0) >= 0 else "liability"})

    return {
        "status": "success", "detection": {"format_type": detection["format_type"], "sheet_name": detection["sheet_name"], "header_row": detection["header_row"], "data_start_row": detection["data_start_row"], "account_col": detection["account_col"], "debit_col": detection["debit_col"], "credit_col": detection["credit_col"], "net_col": detection["net_col"]},
        "accounts": classified, "manual_review": manual_review, "bs_head_options": bs_head_options,
        "summary": {"total_accounts": len(classified), "high_confidence": len(high_conf), "low_confidence": len(low_conf), "unclassified": len(unclassified), "manual_count": len(unclassified)},
    }

def process_tb_to_bs(tb_path, bs_template_path, output_path, user_mapping=None):
    import os, tempfile
    actual_tb_path = tb_path
    tmp_xlsx = None
    if tb_path.lower().endswith('.pdf'):
        tmp_xlsx = tempfile.mktemp(suffix='.xlsx')
        if not convert_pdf_tb_to_xlsx(tb_path, tmp_xlsx): return {"error": "Could not convert PDF trial balance to Excel."}
        actual_tb_path = tmp_xlsx

    analysis = analyze_trial_balance(tb_path)
    if "error" in analysis:
        if tmp_xlsx and os.path.exists(tmp_xlsx): os.remove(tmp_xlsx)
        return analysis

    accounts = analysis["accounts"]

    if user_mapping:
        name_map = {}
        key_map  = {}
        for raw_key, head in user_mapping.items():
            if not head or head in ("auto", " ", None): continue
            rk = str(raw_key).strip()
            parts = rk.rsplit("_", 1)
            if len(parts) == 2 and parts[1].isdigit(): key_map[rk] = head
            else: name_map[rk.upper()] = head

        for acct in accounts:
            acct_key  = str(acct.get("key", "")).strip()
            acct_name = str(acct.get("name", "")).strip().upper()
            if acct_key in key_map: acct["bs_head"] = key_map[acct_key]; acct["confidence"] = "user"
            elif acct_name in name_map: acct["bs_head"] = name_map[acct_name]; acct["confidence"] = "user"

    aggregated = get_aggregated_values(accounts)

    result = inject_into_bs(bs_template_path, output_path, aggregated, mapping_overrides=None, individual_accounts=accounts)
    result["analysis"] = analysis
    result["aggregated"] = aggregated
    result["classified_accounts"] = accounts
    if tmp_xlsx and os.path.exists(tmp_xlsx): os.remove(tmp_xlsx)
    return result
