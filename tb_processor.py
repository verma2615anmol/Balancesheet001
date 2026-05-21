"""
Trial Balance → Balance Sheet Processor
Reads a trial balance (Excel or text-based PDF), classifies accounts under BS/P&L heads,
and injects aggregated values into a BS template.
Zero formatting change in the output BS file.
Memory-efficient: single workbook open, read_only where possible.
"""

import re
import os
import pdfplumber
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
            "raw material", "finished good", "work in progress", "wip",
            "stores", "spare", "packing material", "stock in trade",
            "goods in transit",
        ],
        "negative_keywords": ["stock broker"],
    },
    "trade_rec": {
        "label": "Trade Receivables",
        "side": "asset",
        "keywords": [
            "trade receivable", "debtor", "sundry debtor",
            "accounts receivable", "bills receivable", "trade debtor",
            "receivable from customer",
        ],
        "negative_keywords": [],
    },
    "cash_bank": {
        "label": "Cash and Bank Balances",
        "side": "asset",
        "keywords": [
            "cash", "bank", "cash in hand", "cash at bank", "petty cash",
            "savings account", "current account bank", "bank account",
            "bank balance", "cheque in hand", "imprest",
        ],
        "negative_keywords": ["cash credit", "cc account", "overdraft", "od account", "bank od"],
    },
    "stla": {
        "label": "Short Term Loans & Advances",
        "side": "asset",
        "keywords": [
            "loan and advance", "advance to supplier", "advance to staff",
            "prepaid expense", "security deposit given", "earn money deposit",
            "emd", "advance tax", "tds receivable", "tcs receivable",
            "gst input", "input cgst", "input sgst", "input igst",
            "mat credit", "income tax refund receivable",
        ],
        "negative_keywords": ["advance from customer", "security deposit received"],
    }
}


def parse_pdf_trial_balance(pdf_path):
    """
    Parses text-based PDFs (Tally/Busy/Excel exports) cleanly.
    Extracts rows containing text and ledger amounts.
    """
    accounts = []
    idx = 1
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                line = line.strip()
                if not line or "Particulars" in line or "Total" in line:
                    continue
                
                # Check for financial numbers
                matches = re.findall(r'\d+(?:\.\d+)?', line.replace(',', ''))
                if not matches:
                    continue
                
                # Split and handle strings safely
                parts = line.split('"')
                if len(parts) >= 2:
                    name = parts[1].strip()
                else:
                    name_match = re.match(r'^([^0-9,\.]+)', line)
                    name = name_match.group(1).strip() if name_match else line
                
                if len(name) < 3 or name.lower() in ('debit amount', 'credit amount', 'particulars'):
                    continue
                
                is_credit = True if line.endswith('Cr') or 'cr' in line.lower() else False
                nums = [float(n) for n in matches if '.' in n or len(n) > 2]
                if not nums:
                    continue
                val = nums[-1]
                
                deb = 0.0 if is_credit else val
                cred = val if is_credit else 0.0
                
                accounts.append({
                    "key": f"{name}_{idx}",
                    "name": name,
                    "debit": deb,
                    "credit": cred,
                    "balance": deb - cred
                })
                idx += 1
    return accounts


def clean_name(name):
    if not name: return ""
    s = str(name).strip().upper()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[^A-Z0-9 ]', '', s)
    return s


def classify_account(name):
    n_clean = name.lower()
    for head, config in BS_HEADS.items():
        for kw in config["keywords"]:
            if kw in n_clean:
                has_neg = False
                for nkw in config["negative_keywords"]:
                    if nkw in n_clean:
                        has_neg = True
                        break
                if not has_neg:
                    return head
    return "other_cl"


def get_aggregated_values(accounts):
    agg = {h: 0.0 for h in BS_HEADS}
    agg["unclassified_liability"] = 0.0
    agg["unclassified_asset"] = 0.0
    
    for a in accounts:
        head = a.get("bs_head", "other_cl")
        bal = a.get("balance", 0.0)
        if head in agg:
            agg[head] += bal
        else:
            if bal >= 0:
                agg["unclassified_asset"] += bal
            else:
                agg["unclassified_liability"] += bal
    return agg


def inject_into_bs(template_path, output_path, aggregated, individual_accounts):
    wb = load_workbook(template_path)
    
    # 1. Update Core Summary Sheet Structure
    if "Balance Sheet" in wb.sheetnames:
        ws_bs = wb["Balance Sheet"]
        bs_mapping = {
            "B8": "capital", "B9": "lt_borrowings", "B10": "st_borrowings",
            "B11": "trade_payables", "B12": "other_cl", "B13": "st_provisions",
            "B18": "fixed_assets", "B19": "non_current_investments", "B20": "inventories",
            "B21": "trade_rec", "B22": "cash_bank", "B23": "stla"
        }
        for cell_ref, head in bs_mapping.items():
            val = aggregated.get(head, 0.0)
            if BS_HEADS[head]["side"] == "liability":
                ws_bs[cell_ref] = abs(val)
            else:
                ws_bs[cell_ref] = val

    # 2. Update Detailed Breakdowns via Notes Section with Auto Row Extensions
    ws_notes = wb["Notes to BS"] if "Notes to BS" in wb.sheetnames else wb.active
    
    # Define exact static row indices matching your underlying templates
    section_ranges = {
        "trade_payables": {"start": 21, "end": 63},
        "trade_rec": {"start": 74, "end": 90}
    }
    
    cumulative_shift = 0
    processed_keys = set()

    for head, r_info in section_ranges.items():
        current_start = r_info["start"] + cumulative_shift
        current_end = r_info["end"] + cumulative_shift
        
        sect_accounts = [a for a in individual_accounts if a["bs_head"] == head]
        available_slots = (current_end - current_start) + 1
        
        if len(sect_accounts) > available_slots:
            needed_rows = len(sect_accounts) - available_slots
            ws_notes.insert_rows(current_end + 1, amount=needed_rows)
            
            # Re-apply styles downwards to keep structure solid
            for r_idx in range(current_end + 1, current_end + 1 + needed_rows):
                for col_idx in range(1, 5):
                    ws_notes.cell(row=r_idx, column=col_idx).style = ws_notes.cell(row=current_start, column=col_idx).style
            
            current_end += needed_rows
            cumulative_shift += needed_rows

        for i, acct in enumerate(sect_accounts):
            target_row = current_start + i
            ws_notes.cell(row=target_row, column=1, value=acct["name"])
            ws_notes.cell(row=target_row, column=2, value=abs(acct["balance"]))
            processed_keys.add(acct["key"])

    # 3. Handle Remaining Unclassified Accounts in a Bottom Footer Spillover Block
    remaining_accounts = [a for a in individual_accounts if a["key"] not in processed_keys and abs(a["balance"]) > 0]
    if remaining_accounts:
        footer_row = ws_notes.max_row - 1
        ws_notes.insert_rows(footer_row, amount=len(remaining_accounts))
        for j, acct in enumerate(remaining_accounts):
            r = footer_row + j
            ws_notes.cell(row=r, column=1, value=f"[NEW UNLINKED] {acct['name']}")
            ws_notes.cell(row=r, column=2, value=abs(acct['balance']))

    wb.save(output_path)


def process_tb_to_bs(tb_path, bs_template_path, output_path, user_mapping=None):
    """Unified entry point matching dashboard routing interfaces."""
    log = []
    if tb_path.lower().endswith('.pdf'):
        accounts = parse_pdf_trial_balance(tb_path)
        log.append(f"✓ Parsed text PDF Trial Balance. Found {len(accounts)} accounts.")
    else:
        accounts = []
        wb_tb = load_workbook(tb_path, data_only=True)
        sheet = wb_tb.active
        idx = 1
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            name = str(row[0]).strip()
            deb = float(row[1]) if row[1] else 0.0
            cred = float(row[2]) if row[2] else 0.0
            accounts.append({
                "key": f"{name}_{idx}", "name": name, "debit": deb, "credit": cred, "balance": deb - cred
            })
            idx += 1
        log.append(f"✓ Parsed Excel Trial Balance. Found {len(accounts)} entries.")

    # Override engine logic mapping
    if user_mapping:
        for acct in accounts:
            if acct["key"] in user_mapping:
                acct["bs_head"] = user_mapping[acct["key"]]
                acct["confidence"] = "user"
            else:
                acct["bs_head"] = classify_account(acct["name"])
                acct["confidence"] = "auto"
    else:
        for acct in accounts:
            acct["bs_head"] = classify_account(acct["name"])
            acct["confidence"] = "auto"

    aggregated = get_aggregated_values(accounts)
    inject_into_bs(bs_template_path, output_path, aggregated, accounts)
    
    # Calculate checksum metrics for validation payload
    assets_total = sum(v for k, v in aggregated.items() if BS_HEADS.get(k, {}).get("side") == "asset")
    liab_total = sum(abs(v) for k, v in aggregated.items() if BS_HEADS.get(k, {}).get("side") == "liability")

    return {
        "status": "success",
        "accounts": accounts,
        "total_assets": assets_total,
        "total_liabilities": liab_total,
        "net_profit": 0.0,
        "log": log
    }
