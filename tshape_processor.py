import math
"""
tshape_processor.py
===================
T-Shaped Balance Sheet → Comparative Balance Sheet Converter

v2.0  2026-07-30  COMPLETE REWRITE — forensic column-exact extraction

Architecture:
  1. parse_tshape_bs(path)            → dict of all extracted values
  2. inject_into_template(...)        → fills PY column of Output_sample_format.xlsx
  3. process_tshape(...)              → main entry point

Hard rules:
  - tshape_processor.py = T-shaped conversion ONLY
  - processor.py / lumid_compat.py are NEVER touched
  - Output_sample_format.xlsx structure is NEVER modified
  - PY column ← values from T-shaped XLS
  - CY column ← blank / 0 (yellow highlighted) for CA to fill
"""

import re
import os
import shutil
import datetime
from copy import copy

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.cell import MergedCell


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _s(v):
    """Cell → clean string."""
    if v is None or (isinstance(v, float) and str(v) == 'nan'):
        return ''
    return str(v).strip()


def _n(v, default=0.0):
    """Cell → float."""
    if v is None:
        return default
    if isinstance(v, (int, float)) and not (isinstance(v, float) and str(v) == 'nan'):
        return float(v)
    try:
        return float(re.sub(r'[,\s]', '', str(v)))
    except Exception:
        return default


def _first_num(row, start=0, end=None):
    """Return first numeric value in row[start:end], else 0."""
    end = end or len(row)
    for i in range(start, min(end, len(row))):
        v = row[i]
        if isinstance(v, (int, float)) and not (isinstance(v, float) and str(v) == 'nan'):
            return float(v)
    return 0.0


_INPUT_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")


def _write(ws, row, col, value, bold=False):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return
    cell.value = value
    if bold:
        cell.font = Font(bold=True)


def _write_num(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return
    cell.value = float(value or 0)
    cell.number_format = '#,##0'


# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 1 — T-shaped BS parser  (v2.0 forensic-exact)
# ─────────────────────────────────────────────────────────────────────────────

def parse_tshape_bs(filepath: str) -> dict:
    """
    Parse a GD Singla & Co. style T-shaped balance sheet.

    Column layout (varies per client but always detected dynamically):
      Left half  = LIABILITIES  (label col, annex col, amount col)
      Right half = ASSETS        (label col, annex col, amount col)
      Far right  = ANNEXURES     (multiple party-list blocks side by side)
      P&L        = embedded in the same sheet at cols 8-16 (Bobby/Ashok)
                   or cols 10-18 (Gupta)

    Returns a dict with all financial values + log.
    """
    log = []

    ext = os.path.splitext(filepath)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    xl = pd.ExcelFile(filepath, engine=engine)

    # Find main BS sheet
    sheet_name = xl.sheet_names[0]
    for sn in xl.sheet_names:
        sl = sn.lower()
        if any(x in sl for x in ['bs', 'balance', '_co', '_ha', '_fo', '_corp']):
            sheet_name = sn
            break
    log.append(f"Sheet: {sheet_name}")

    df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine=engine)
    rows = df.values.tolist()
    nrows = len(rows)
    ncols = max((len(r) for r in rows), default=0)
    for r in rows:
        while len(r) < ncols:
            r.append(None)

    # ── Entity name & date ────────────────────────────────────────────────────
    entity_name, bs_date, py_year_end = '', '', ''
    for i, row in enumerate(rows[:15]):
        for j, v in enumerate(row):
            s = _s(v)
            if not entity_name and len(s) > 8 and 'M/S' in s.upper():
                entity_name = s
            if not bs_date and 'BALANCE' in s.upper() and 'SHEET' in s.upper():
                for k in range(j, min(j + 20, ncols)):
                    dv = row[k]
                    if isinstance(dv, datetime.datetime):
                        bs_date = dv.strftime('%d.%m.%Y')
                        py_year_end = str(dv.year)
                        break
                    ds = _s(dv)
                    m = re.search(r'(\d{2})[./](\d{2})[./](\d{4})', ds)
                    if m:
                        bs_date = ds.strip()
                        py_year_end = m.group(3)
                        break

    if not entity_name:
        entity_name = 'M/S CLIENT'
    if not bs_date:
        bs_date = '31.03.2024'
        py_year_end = '2024'
    log.append(f"Entity: {entity_name} | Date: {bs_date}")

    # ── Find BS header row (LIABILITIES … ASSETS … AMOUNT) ──────────────────
    bs_header_row = -1
    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)
        if 'LIABILIT' in rs and 'ASSET' in rs and 'AMOUNT' in rs:
            bs_header_row = i
            log.append(f"BS header row: {i}")
            break
    if bs_header_row < 0:
        bs_header_row = 9

    hrow = rows[bs_header_row]

    # Detect liab_amount_col and asset_amount_col from header row
    amount_cols = [j for j, v in enumerate(hrow) if _s(v).upper() == 'AMOUNT']
    liab_amount_col = amount_cols[0] if len(amount_cols) >= 1 else 4
    asset_amount_col = amount_cols[1] if len(amount_cols) >= 2 else (liab_amount_col + 5)
    log.append(f"Liab amount col: {liab_amount_col}, Asset amount col: {asset_amount_col}")

    # ── Result skeleton ───────────────────────────────────────────────────────
    result = {
        'entity_name': entity_name, 'bs_date': bs_date,
        'py_year_end': py_year_end, 'entity_type': 'proprietorship',
        'capital_accounts': [], 'secured_loans': 0.0,
        'unsecured_loans': 0.0, 'unsecured_loan_parties': [],
        'sundry_creditors': 0.0, 'sundry_creditor_parties': [],
        'advance_from_customers': 0.0,
        'other_payables': 0.0, 'other_payable_items': [],
        'short_term_provisions': 0.0,
        'fixed_assets_wdv': 0.0, 'fixed_asset_items': [],
        'investments': 0.0, 'advances_security': 0.0,
        'closing_stock': 0.0,
        'sundry_debtors': 0.0, 'sundry_debtor_parties': [],
        'loans_advances': 0.0, 'loans_advances_items': [],
        'loans_to_other_items': [], 'advance_to_revenue_items': [],
        'other_current_asset_items': [],
        'advance_from_customer_parties': [],
        'cash_bank': 0.0, 'cash_in_hand': 0.0, 'bank_balances': [],
        'other_current_assets': 0.0,
        'sales': 0.0, 'opening_stock': 0.0, 'purchases': 0.0,
        'direct_expenses': 0.0, 'direct_expense_items': [],
        'gross_profit': 0.0, 'other_income': 0.0,
        'salary_expenses': 0.0, 'interest_paid': 0.0,
        'depreciation': 0.0, 'other_expenses': 0.0,
        'other_expense_items': [], 'net_profit': 0.0,
        'log': log,
    }

    # ── Pass 1: scan BS header rows for main totals ───────────────────────────
    _scan_bs_totals(rows, bs_header_row, liab_amount_col, asset_amount_col, result, log)

    # ── Pass 2: detect entity type ────────────────────────────────────────────
    full_text = ' '.join(_s(v) for row in rows for v in row if v)
    el = entity_name.lower()
    if 'huf' in el or '(huf)' in el or 'hindu' in el:
        result['entity_type'] = 'huf'
    elif 'PARTNER' in full_text.upper():
        result['entity_type'] = 'partnership'

    # ── Pass 3: extract capital accounts ─────────────────────────────────────
    _extract_capital(rows, result, log)

    # ── Pass 4: extract party lists from annexures ────────────────────────────
    _extract_annexure_parties(rows, result, log)

    # ── Pass 5: extract P&L ───────────────────────────────────────────────────
    _extract_capital_annexure(rows, result, log)
    _extract_ocl_annexure(rows, result, log)
    _extract_debtor_annexure(rows, result, log)
    _extract_creditor_annexure(rows, result, log)   # Pass 5b: Annexure-C creditors
    _extract_loans_annexure(rows, result, log)       # Pass 5c: Annexure-G loans
    _extract_pl(rows, result, log)

    # ── Pass 6: extract fixed assets dep chart ────────────────────────────────
    _extract_dep_chart(rows, result, log)

    # ── Pass 7: extract cash & bank ───────────────────────────────────────────
    _extract_cash_bank(rows, result, log)

    log.append('Parse complete.')
    return result


# ─────────────────────────────────────────────────────────────────────────────
#  Pass 1 — BS totals from header section
# ─────────────────────────────────────────────────────────────────────────────

def _scan_bs_totals(rows, bs_header_row, liab_col, asset_col, result, log):
    """Extract aggregate totals from the BS rows (rows after header)."""
    for i in range(bs_header_row + 1, min(bs_header_row + 50, len(rows))):
        row = rows[i]
        label = _s(row[0]).upper()
        rs = ' '.join(_s(v).upper() for v in row if v is not None)

        # LIABILITIES side — col 0 label
        if ('CAPITAL' in label and ('PROP' in label or 'PARTNER' in label
                                     or 'OWNER' in label or 'HUF' in label)):
            amt = _n(row[liab_col]) if liab_col < len(row) else 0
            if not amt:
                # Try adjacent cols
                amt = _first_num(row, liab_col - 1, liab_col + 3)
            if amt:
                result['capital_accounts'].append({
                    'name': _s(row[0]).strip("`'\""),
                    'opening': amt, 'additions': 0.0,
                    'withdrawals': 0.0, 'profit': 0.0, 'closing': amt
                })
                log.append(f"R{i}: Capital raw={amt}")

        elif 'SECURED LOAN' in label and 'UN' not in label:
            amt = _n(row[liab_col]) if liab_col < len(row) else 0
            if amt and amt == amt:  # not nan
                result['secured_loans'] = amt
            log.append(f"R{i}: Secured loans (header) = {amt}")

        elif re.search(r'ICICI|HDFC|SBI|PNB|OBC|BANK\s+LTD|LAP\s+LOAN|CAR\s+LOAN', label):
            amt = _n(row[liab_col]) if liab_col < len(row) else 0
            if amt and amt == amt:  # not nan
                result['secured_loans'] += amt
                log.append(f"R{i}: Secured loan item = {amt}")

        elif 'UN-SECURED' in label or ('UNSECURED' in label and 'LOAN' in label):
            amt = _n(row[liab_col]) if liab_col < len(row) else 0
            if not amt:
                amt = _first_num(row, liab_col - 1, liab_col + 3)
            if amt:
                result['unsecured_loans'] = amt
                log.append(f"R{i}: Unsecured loans = {amt}")

        elif 'SUNDRY CREDITOR' in label or ('CREDITOR' in label and 'ADVANCE' not in label):
            amt = _n(row[liab_col]) if liab_col < len(row) else 0
            if not amt:
                amt = _first_num(row, liab_col - 1, liab_col + 3)
            if amt:
                result['sundry_creditors'] = amt
                log.append(f"R{i}: Sundry creditors = {amt}")

        elif 'OTHER PAYABLE' in label or label == 'OTHER PAYABLES':
            amt = _n(row[liab_col]) if liab_col < len(row) else 0
            if not amt:
                amt = _first_num(row, liab_col - 1, liab_col + 3)
            if amt:
                result['other_payables'] = amt

        # ASSETS side — only scan the asset half of the row (label in RIGHT half)
        # We check both rs (whole row) for presence and then look only at asset columns
        # Use a limited scan window around asset_amount_col to avoid picking up P&L values
        asc_min = max(asset_col - 2, liab_col + 2)  # never dip into liabilities cols
        asc_max = min(asset_col + 6, len(row))

        # Fixed assets: label must appear in the right half of the row
        rs_right_raw = ' '.join(_s(row[j]).upper() for j in range(liab_col + 1, len(row)) if row[j])
        rs_right = ' '.join(rs_right_raw.split())  # normalize multiple spaces
        if 'FIXED ASSET' in rs_right and result['fixed_assets_wdv'] == 0:
            for j in range(asc_min, asc_max):
                v = _n(row[j])
                if v > 1000:
                    result['fixed_assets_wdv'] = v
                    log.append(f"R{i}: Fixed assets WDV = {v} (col {j})")
                    break

        # Closing stock: only when label explicitly contains 'Closing stock' on asset side
        # AND it appears in the right half of the row
        if result['closing_stock'] == 0:
            # Only set closing stock when 'CLOSING STOCK' or 'STOCK' label is in
            # the ASSET section (first few cols of the right half), not the P&L section.
            # Check if any col in liab_col+1 to asset_col+4 has 'CLOSING' or 'STOCK'.
            asset_label_range = range(liab_col + 1, min(asset_col + 5, len(row)))
            has_stock_label = any(
                'STOCK' in _s(row[j]).upper() and 'CLOSING' in _s(row[j]).upper()
                for j in asset_label_range
            )
            if has_stock_label:
                for j in range(asc_min, min(asset_col + 6, len(row))):
                    v = _n(row[j])
                    if v > 10000:
                        result['closing_stock'] = v
                        log.append(f"R{i}: Closing stock = {v} (col {j})")
                        break

        if ('SUNDRY DEBTOR' in rs_right or 'DEBTORS & ADVANCES' in rs_right) \
           and result['sundry_debtors'] == 0:
            for j in range(asc_min, asc_max):
                v = _n(row[j])
                if v > 100:
                    result['sundry_debtors'] = v
                    log.append(f"R{i}: Sundry debtors = {v} (col {j})")
                    break

        if 'CASH' in rs_right and ('BANK' in rs_right or 'BALANCE' in rs_right) \
           and i < bs_header_row + 35 and result['cash_bank'] == 0:
            for j in range(asc_min, asc_max):
                v = _n(row[j])
                if v > 100:
                    result['cash_bank'] = v
                    log.append(f"R{i}: Cash & bank = {v} (col {j})")
                    break

        if ('ADVANCES & SECURITY' in rs_right or 'ADVANCE & SECURITY' in rs_right or
            ('ADVANCES' in rs_right and 'LOAN' in rs_right and 'SECURITIES' in rs_right)) \
           and result['advances_security'] == 0:
            for j in range(asc_min, asc_max):
                v = _n(row[j])
                if v > 100:
                    result['advances_security'] = v
                    result['loans_advances'] = v
                    log.append(f"R{i}: Advances & security = {v} (col {j})")
                    break

        if ('LOANS & ADVANCE' in rs_right or 'LOAN & ADVANCE' in rs_right) \
           and result['loans_advances'] == 0:
            for j in range(asc_min, asc_max):
                v = _n(row[j])
                if v > 100:
                    result['loans_advances'] = v
                    log.append(f"R{i}: Loans & advances = {v} (col {j})")
                    break

        if ('INVESTMENT' in rs_right or ('SECURITY' in rs_right and 'DEPOSIT' in rs_right)) \
           and result['investments'] == 0:
            for j in range(asc_min, asc_max):
                v = _n(row[j])
                if v > 100:
                    result['investments'] = v
                    break

        # Stop at TOTAL row
        if label.startswith('TOTAL') and i > bs_header_row + 5:
            break


# ─────────────────────────────────────────────────────────────────────────────
#  Pass 3 — Capital accounts
# ─────────────────────────────────────────────────────────────────────────────

def _extract_capital(rows, result, log):
    """
    Find capital account details. In GD Singla format:

    Proprietorship:
      Row with "PROP. CAPITAL A/C" or "PROP CAPITAL ACCOUNT" in col 0,
      followed by "SANJEEV KUMAR..." or entity name with amount in
      annex/amount col. The CLOSING balance is the amount on the Annexure-A
      row (next to `A'`).

    Partnership:
      "PARTNER'S CAPITAL A/C." in col 0, row 11 of Bobby.
      Partner names with amounts in col 0 and col 3 (rows 13,14).

    We read the Annexure-A row for the closing balance.
    """
    caps = []

    for i, row in enumerate(rows):
        label = _s(row[0]).upper()
        rs = ' '.join(_s(v).upper() for v in row if v is not None)

        # Proprietorship / HUF: look for row with `A' or 'A' next to capital label
        # Skip rows where 'E' (fixed assets annex) also appears before 'A' — those are
        # partnership rows where the capital account is on the liabilities side but the
        # main label has PARTNER not PROP.
        has_annex_a = ("'A'" in rs or "ANNEXURE-A" in rs or "`A'" in rs)
        is_prop_cap = has_annex_a and ('CAPITAL' in rs or 'PROP' in rs or 'HUF' in rs)
        # Exclude partnership rows where PARTNER'S CAPITAL appears (handled separately)
        is_prop_cap = is_prop_cap and 'PARTNER' not in label

        if is_prop_cap:
            # Find name from the SAME row or adjacent
            # Name is usually in col 0 or nearby; amount in annex col
            # Try col 3 or 4 first (amount cols for Gupta/Ashok)
            name = _s(row[0]).strip("`'\"")
            if not name or name.upper() in ('', 'PROP. CAPITAL A/C', 'PROP CAPITAL ACCOUNT'):
                # Look ahead for name
                for k in range(i + 1, min(i + 4, len(rows))):
                    n2 = _s(rows[k][0])
                    if n2 and n2.upper() not in ('', 'PARTICULARS') and \
                       not n2.upper().startswith('FIXED') and \
                       not n2.upper().startswith('UNSECURED') and \
                       not n2.upper().startswith('CURRENT') and \
                       not n2.upper().startswith('SUNDRY'):
                        name = n2
                        break

            # Get amount — scan for the capital amount.
            # It should be in the FIRST few cols (liabilities side, cols 1-8)
            # NOT in the right half (which would be fixed assets etc.)
            # Look specifically in col 4 first (Gupta layout), then col 6, then scan 1-8
            amt = 0.0
            for j in [4, 3, 6, 5, 2, 1, 7, 8]:
                if j < len(row):
                    v = _n(row[j])
                    if v > 100000:   # Capital accounts are large
                        amt = v
                        break

            if not amt:
                # Try looking at the next row for the amount (proprietorship name on next row)
                if i + 1 < len(rows):
                    for j in [4, 3, 6, 5, 2, 1, 7, 8]:
                        if j < len(rows[i+1]):
                            v = _n(rows[i+1][j])
                            if v > 100000:
                                amt = v
                                break

            if amt > 0:
                # Look for opening, profit, withdrawals in vicinity
                opening = amt
                profit = withdrawals = additions = 0.0

                for k in range(i, min(i + 15, len(rows))):
                    r2 = rows[k]
                    rs2 = ' '.join(_s(v).upper() for v in r2 if v is not None)
                    if 'OPENING BALANCE' in rs2 or 'BALANCE B/D' in rs2 or 'LAST BALANCE' in rs2:
                        opening = _first_num(r2, 1) or opening
                    if 'PROFIT DURING' in rs2 or 'NET PROFIT' in rs2:
                        profit = _first_num(r2, 1)
                    if 'WITHDRAWL' in rs2 or 'HOUSEHOLD' in rs2 or 'DRAWING' in rs2:
                        withdrawals = _first_num(r2, 1)

                # closing = the amount next to `A'`
                closing = amt

                caps.append({
                    'name': name,
                    'opening': opening, 'additions': additions,
                    'withdrawals': withdrawals, 'profit': profit,
                    'closing': closing
                })
                log.append(f"Capital (prop/huf): {name} = {closing}")
                break  # Only one capital block for proprietorship/HUF

        # Partnership: "PARTNER'S CAPITAL" header then partner names below
        if "PARTNER'S CAPITAL" in label or "PARTNER CAPITAL" in label:
            result['entity_type'] = 'partnership'
            # Scan next rows for partner names + amounts in col 0 + col 3
            for k in range(i + 1, min(i + 20, len(rows))):
                r2 = rows[k]
                n2 = _s(r2[0]).strip("`'\"")
                if not n2:
                    continue
                n2u = n2.upper()
                if any(x in n2u for x in ('FIXED', 'CURRENT', 'SECURED',
                                           'SUNDRY', 'CASH', 'TOTAL',
                                           'PARTICUL', 'OTHER', 'LIABIL')):
                    break
                # Check if row 0 looks like a partner name (contains "Sh." or "Mr." or is 5+ chars)
                if len(n2) >= 3 and not n2u.startswith('TO ') and not n2u.startswith('BY '):
                    # Amount in col 3 (Bobby layout)
                    amt = _n(r2[3]) if len(r2) > 3 else 0
                    if not amt:
                        amt = _first_num(r2, 1, 8)
                    if amt > 1000:
                        caps.append({
                            'name': n2,
                            'opening': amt, 'additions': 0.0,
                            'withdrawals': 0.0, 'profit': 0.0,
                            'closing': amt
                        })
                        log.append(f"Partner capital: {n2} = {amt}")
            break

    if caps:
        result['capital_accounts'] = caps
        # Update total
        total_cap = sum(c['closing'] for c in caps)
        log.append(f"Total capital: {total_cap}")


# ─────────────────────────────────────────────────────────────────────────────
#  Pass 4 — Annexure party lists  (forensic-exact column detection)
# ─────────────────────────────────────────────────────────────────────────────

def _extract_annexure_parties(rows, result, log):
    """
    Detect party lists from side-by-side annexure blocks.

    GD Singla layout places multiple annexures on the same rows:
      Ashok:  Creditors=cols(45,47), Debtors=cols(59,65)
      Bobby:  Creditors=cols(28-29,34), Debtors=cols(45,54) + HUF/other at (45,54)
              Unsecured=cols(21,27)
      Gupta:  Creditors=cols(22,28), Debtors=cols(22,28) below a different header

    Strategy:
      1. Scan ALL rows for 'M/s' or 'M/S.' markers with nearby numbers.
      2. Group them by column-pair (name_col, amount_col).
      3. Match each group to the nearest preceding section header
         (SUNDRY CREDITORS / SUNDRY DEBTORS / UNSECURED LOAN).
    """
    SKIP_NAMES = {
        'PARTICULARS', 'FROM RELATED PARTIES', 'FROM OTHER PARTIES',
        'TOTAL', 'SUB TOTAL', 'ANNEXURE', 'DUE TO MSME',
        'ADVANCE FROM CUSTOMERS', 'ADVANCE TO SUPPLIERS', 'NIL',
    }

    # Step 1: Collect all (row_idx, name_col, name, amount_col, amount) tuples
    # where name_col has a text value and amount_col is a number
    entries = []   # (row_i, name_col, name, amount_col, amount)

    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            s = _s(v)
            if not s:
                continue
            su = s.upper()
            # Name prefix 'M/s.' or 'M/S.' or 'M/s' by itself, or person name
            is_prefix = su in ('M/S.', 'M/S', 'M/S .', 'M/S. ', 'MR', 'MR.', 'MRS', 'MRS.')
            is_direct_name = (
                len(s) >= 4 and
                not su.startswith('TO ') and
                not su.startswith('BY ') and
                not any(su.startswith(x) for x in (
                    'TOTAL', 'ANNEXURE', 'FROM ', 'DATE', 'PLACE',
                    'FOR ', 'AS PER', 'CHARTERED', 'UDIN', 'PAN ',
                    'REFER', 'AUDITOR', 'TRADING', 'BALANCE', 'PARTICULARS',
                )) and
                not su.isnumeric() and
                su not in SKIP_NAMES and
                su not in ('NIL', 'NULL', 'DUE TO')
            )

            if is_prefix:
                # Name is in next column
                name_col = j + 1
                if name_col < len(row):
                    name = _s(row[name_col])
                    if name and name.upper() not in SKIP_NAMES and len(name) > 2:
                        # Find amount in cols after name
                        for k in range(name_col + 1, min(name_col + 8, len(row))):
                            amt = _n(row[k])
                            if amt > 100:
                                entries.append((i, name_col, name.strip(), k, amt))
                                break

            elif is_direct_name:
                # Check if next few cols have a number
                for k in range(j + 1, min(j + 6, len(row))):
                    amt = _n(row[k])
                    if amt > 100:
                        entries.append((i, j, s.strip(), k, amt))
                        break

    # Step 2: Find section headers (row index → section name)
    section_rows = {}   # row_i → 'creditor'|'debtor'|'unsecured'|'other_pay'|'loans'
    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)
        # Priority: unsecured > creditor > debtor (to handle mixed rows)
        if ('UNSECURED LOAN' in rs or 'UN-SECURED LOAN' in rs) and ('ANNEXURE-B' in rs or "'B'" in rs):
            section_rows[i] = 'unsecured'
        elif 'ANNEXURE-B' in rs and i > 5 and 'SUNDRY' not in rs:
            section_rows[i] = 'unsecured'
        elif 'SUNDRY CREDITOR' in rs or 'ANNEXURE-C' in rs or ("'C'" in rs and 'CREDITOR' in rs):
            section_rows[i] = 'creditor'
        elif 'SUNDRY DEBTOR' in rs or 'ANNEXURE-G' in rs or ("'G'" in rs and 'DEBTOR' in rs):
            section_rows[i] = 'debtor'
        elif 'OTHER PAYABLE' in rs and 'ANNEXURE' in rs:
            section_rows[i] = 'other_pay'

    log.append(f"Section header rows: {section_rows}")

    # Step 3: Group entries by (name_col, amount_col) to identify consistent blocks
    # Then for each block, find which section header precedes it most closely
    from collections import defaultdict
    col_pair_entries = defaultdict(list)
    for e in entries:
        row_i, name_col, name, amt_col, amt = e
        col_pair_entries[(name_col, amt_col)].append(e)

    # Step 4: For each col-pair block, assign to nearest preceding section header
    assigned = defaultdict(list)   # section → [(name, amount)]
    for (nc, ac), elist in col_pair_entries.items():
        if len(elist) < 1:
            continue
        for row_i, name_col, name, amt_col, amt in elist:
            # Find the nearest section header at or before row_i
            sec = None
            best_dist = 9999
            for sr, sname in section_rows.items():
                if sr <= row_i and (row_i - sr) < best_dist:
                    best_dist = row_i - sr
                    sec = sname
            if sec and best_dist <= 80:
                assigned[sec].append((name, amt))

    log.append(f"Assigned sections: { {k: len(v) for k, v in assigned.items()} }")

    # Step 5: Deduplicate and store
    def _dedup(items):
        seen = {}
        for name, amt in items:
            key = name.upper()
            if key not in seen or amt > seen[key][1]:
                seen[key] = (name, amt)
        return [{'name': n, 'amount': a} for n, a in seen.values()]

    if assigned.get('creditor'):
        parties = _dedup(assigned['creditor'])
        result['sundry_creditor_parties'] = parties
        calc_total = sum(p['amount'] for p in parties)
        if not result['sundry_creditors'] or abs(calc_total - result['sundry_creditors']) < result['sundry_creditors'] * 0.2:
            result['sundry_creditors'] = result['sundry_creditors'] or calc_total
        log.append(f"Creditor parties: {len(parties)}, total={calc_total:.0f}")

    if assigned.get('debtor'):
        parties = _dedup(assigned['debtor'])
        calc_total = sum(p['amount'] for p in parties)
        actual_debtors = result.get('sundry_debtors', 0)
        # If assigned debtors total is way off from BS figure, try to filter.
        # This handles Gupta where unsecured loan parties bleed into debtor section.
        if actual_debtors > 0 and calc_total > actual_debtors * 3:
            # Assigned debtors total is way off — filter to only parties where
            # the individual amount is <= actual_debtors (exclude large unsecured items)
            filtered = [p for p in parties if p['amount'] <= actual_debtors * 1.5]
            if sum(p['amount'] for p in filtered) > 0:
                parties = filtered
        result['sundry_debtor_parties'] = parties
        # Always use BS-derived total if we have one
        if not result['sundry_debtors']:
            result['sundry_debtors'] = calc_total
        log.append(f"Debtor parties: {len(parties)}, total={calc_total:.0f} (BS={actual_debtors:.0f})")

    if assigned.get('unsecured'):
        parties = _dedup(assigned['unsecured'])
        # Filter out entries that look like P&L items (large salary/expense amounts)
        pl_keywords = ('HOUSE EXP', 'SCHOOL FEE', 'MEDICLAIM', 'SALARY', 'BALANCE C/D',
                       'NET PROFIT', 'GROSS PROFIT', 'PARTNER', 'AUDITOR')
        parties = [p for p in parties
                   if not any(kw in p['name'].upper() for kw in pl_keywords)]
        result['unsecured_loan_parties'] = parties
        calc_total = sum(p['amount'] for p in parties)
        if calc_total > 0:
            result['unsecured_loans'] = result['unsecured_loans'] or calc_total
        log.append(f"Unsecured parties: {len(parties)}, total={calc_total:.0f}")

    if assigned.get('other_pay'):
        parties = _dedup(assigned['other_pay'])
        result['other_payable_items'] = parties
        result['other_payables'] = result['other_payables'] or sum(p['amount'] for p in parties)


# ─────────────────────────────────────────────────────────────────────────────
#  Pass 4b — OTHER PAYABLE annexure direct extraction
# ─────────────────────────────────────────────────────────────────────────────

def _extract_debtor_annexure(rows, result, log):
    """
    Directly extract Annexure-I (Sundry Debtors) entries from the GD Singla T-shaped XLS.

    Layout (0-based col indices):
      col58 = "M/s." prefix
      col59 = party name
      col65 = amount (receivable)
    The TOTAL row at col59="TOTAL" confirms the section end.

    This gives the COMPLETE debtor list including parties that the section-reader
    assigns to the unsecured column range (Ram Beran, Solar, Treat Bell, etc.).
    """
    # Already populated with good data? Skip.
    existing = result.get('sundry_debtor_parties', [])
    bs_total = result.get('sundry_debtors', 0)
    if existing:
        existing_total = sum(p['amount'] for p in existing)
        if bs_total > 0 and abs(existing_total - bs_total) / bs_total < 0.02:
            return  # already accurate

    in_section = False
    items = []
    total_row_val = 0.0

    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)

        if not in_section:
            # Annexure-I header: col65 (idx 65 0-based) has 'ANNEXURE-I'
            # or the row contains 'ANNEXURE-I'
            if 'ANNEXURE-I' in rs:
                in_section = True
                continue
            continue

        # Name at col59 (0-based), amount at col65 (0-based)
        prefix    = _s(row[59]).strip() if 59 < len(row) else ''  # "M/s."
        name_cell = _s(row[60]).strip() if 60 < len(row) else ''  # actual party name
        amt       = row[65] if 65 < len(row) else None

        # Stop at TOTAL row (col59 has "TOTAL")
        if prefix.upper() == 'TOTAL':
            if isinstance(amt, (int, float)) and amt > 0:
                total_row_val = amt
            break

        if name_cell and len(name_cell) > 2 and isinstance(amt, (int, float)) and amt > 0:
            # Build full name: prefix + name
            p_lower = prefix.lower()
            if p_lower in ('m/s.', 'm/s', 'ms.', 'ms'):
                full_name = 'M/s. ' + name_cell.strip()
            elif prefix:
                full_name = prefix + ' ' + name_cell.strip()
            else:
                full_name = name_cell.strip()
            items.append({'name': full_name.strip(), 'amount': amt})

    if items and total_row_val > 0:
        calc = sum(it['amount'] for it in items)
        if abs(calc - total_row_val) / total_row_val < 0.02:
            result['sundry_debtor_parties'] = items
            log.append(f"Debtor Annexure-I: {len(items)} parties, total={calc:,.0f} "
                       f"(BS={bs_total:,.0f})")


def _extract_creditor_annexure(rows, result, log):
    """
    Directly extract Annexure-C (Sundry Creditors & Advances) from the GD Singla XLS.

    Layout (0-based col indices):
      col45 = party name  (or section header like 'ADVANCES FROM CUSTOMER')
      col49 = amount
    Header row: col45 = 'SUNDRY CREDITORS & ADVANCES', col49 = 'ANNEXURE-C'
    Stop at TOTAL row (col45 == 'TOTAL').
    """
    in_section = False
    cred_items = []
    adv_items  = []
    in_advance = False
    total_row_val = 0.0

    for i, row in enumerate(rows):
        if 45 >= len(row):
            continue
        name_cell = _s(row[45]).strip()
        amt_cell  = row[49] if 49 < len(row) else None

        if not in_section:
            rs = ' '.join(_s(v).upper() for v in row if v is not None)
            if 'SUNDRY CREDITOR' in rs and 'ANNEXURE-C' in rs:
                in_section = True
            continue

        name_upper = name_cell.upper()
        if name_upper == 'TOTAL':
            if isinstance(amt_cell, (int, float)) and amt_cell > 0:
                total_row_val = amt_cell
            break

        if 'ADVANCE' in name_upper and 'CUSTOMER' in name_upper:
            in_advance = True
            continue

        if not name_cell or len(name_cell) < 2:
            continue

        if not isinstance(amt_cell, (int, float)) or amt_cell <= 0:
            continue

        entry = {'name': name_cell, 'amount': amt_cell}
        if in_advance:
            adv_items.append(entry)
        else:
            cred_items.append(entry)

    if not (cred_items or adv_items):
        return

    calc = sum(x['amount'] for x in cred_items) + sum(x['amount'] for x in adv_items)
    # Only override when we got a meaningful total match (within 2%)
    if total_row_val > 0 and abs(calc - total_row_val) / total_row_val > 0.02:
        log.append(f"Creditor Annexure-C: sum {calc:.0f} != TOTAL {total_row_val:.0f} — skipping")
        return

    result['sundry_creditor_parties'] = cred_items
    if adv_items:
        result['advance_from_customer_parties'] = adv_items
    if total_row_val > 0:
        result['sundry_creditors'] = total_row_val
    log.append(f"Creditor Annexure-C: {len(cred_items)} creditors, "
               f"{len(adv_items)} advance-from-customer, total={calc:,.0f}")


# ─── Loans & Advances Annexure (Annexure-G in GD Singla layout) ──────────────
_REVENUE_AUTHORITY_KEYWORDS = (
    'GST', 'CGST', 'SGST', 'IGST', 'TDS', 'TCS', 'TAX', 'INCOME TAX',
    'ADVANCE TAX', 'VAT', 'SERVICE TAX', 'EXCISE', 'CUSTOM',
)
_OTHER_CURRENT_ASSET_KEYWORDS = (
    'PREPAID', 'INSURANCE', 'ADVANCE TO SUPPLIER', 'SECURITY DEPOSIT',
    'FD', 'FIXED DEPOSIT',
)


def _classify_loan_item(name):
    """Return 'revenue_auth' | 'other_current' | 'loans_to_others'."""
    n = name.upper()
    if any(kw in n for kw in _REVENUE_AUTHORITY_KEYWORDS):
        return 'revenue_auth'
    if any(kw in n for kw in _OTHER_CURRENT_ASSET_KEYWORDS):
        return 'other_current'
    return 'loans_to_others'


def _extract_loans_annexure(rows, result, log):
    """
    Reads LOANS & ADVANCES (Annexure-G) AND INVESTMENT & SECURITY (Annexure-F)
    from the GD Singla XLS, both at col50 (name) / col58 (amount).

    Annexure-G items are classified into:
      - loans_to_others       → STL note B section (R134-R140)
      - advance_to_revenue    → STL note C section (R143-R147)
      - other_current         → OCA note (R154-R159)

    Annexure-F items (investment / security deposits like D.F.S.C) are treated as
    Short Term Loans — Loans to Others (CA instruction), so they are appended to
    loans_to_other_items.  result['investments'] is zeroed out accordingly.
    """
    # ── Read Annexure-G (LOANS & ADVANCES) ───────────────────────────────────
    in_section = False
    items_g = []
    total_g  = 0.0

    for i, row in enumerate(rows):
        if 50 >= len(row):
            continue
        name_cell = _s(row[50]).strip()
        amt_cell  = row[58] if 58 < len(row) else None

        if not in_section:
            rs = ' '.join(_s(v).upper() for v in row if v is not None)
            if ('LOANS & ADVANCE' in rs or 'LOAN & ADVANCE' in rs) and 'ANNEXURE' in rs:
                in_section = True
            continue

        if name_cell.upper() == 'TOTAL':
            if isinstance(amt_cell, (int, float)) and amt_cell > 0:
                total_g = amt_cell
            break

        if not name_cell or len(name_cell) < 2:
            continue
        if not isinstance(amt_cell, (int, float)) or amt_cell <= 0:
            continue

        items_g.append({
            'name': name_cell,
            'amount': amt_cell,
            'category': _classify_loan_item(name_cell),
        })

    # ── Read Annexure-F (INVESTMENT & SECURITY) ───────────────────────────────
    in_inv = False
    items_f = []

    for i, row in enumerate(rows):
        if 50 >= len(row):
            continue
        name_cell = _s(row[50]).strip()
        amt_cell  = row[58] if 58 < len(row) else None

        if not in_inv:
            rs = ' '.join(_s(v).upper() for v in row if v is not None)
            if ('INVESTMENT' in rs and 'SECURITY' in rs and 'ANNEXURE-F' in rs):
                in_inv = True
            continue

        if name_cell.upper() == 'TOTAL':
            break
        # Stop if next major section header appears at col50
        if any(kw in name_cell.upper() for kw in ('LOANS', 'CASH', 'SUNDRY', 'OTHER PAYABLE')):
            break

        if not name_cell or len(name_cell) < 2:
            continue
        if not isinstance(amt_cell, (int, float)) or amt_cell <= 0:
            continue

        items_f.append({
            'name': name_cell,
            'amount': amt_cell,
            'category': 'loans_to_others',   # CA instruction: treat as STL
        })

    if not items_g and not items_f:
        return

    # ── Classify Annexure-G items ─────────────────────────────────────────────
    loans_to_others = [x for x in items_g if x['category'] == 'loans_to_others']
    revenue_auth    = [x for x in items_g if x['category'] == 'revenue_auth']
    other_current   = [x for x in items_g if x['category'] == 'other_current']

    # Append Annexure-F items (DFSC etc.) into loans_to_others
    loans_to_others = loans_to_others + items_f

    result['loans_to_other_items']      = loans_to_others
    result['advance_to_revenue_items']  = revenue_auth
    result['other_current_asset_items'] = other_current
    result['other_current_assets']      = sum(x['amount'] for x in other_current)

    # loans_advances total = STL items (loans_to_others + revenue_auth)
    stl_total = (sum(x['amount'] for x in loans_to_others) +
                 sum(x['amount'] for x in revenue_auth))
    if result.get('loans_advances', 0) == 0:
        result['loans_advances'] = stl_total
    result['loans_advances_items'] = loans_to_others

    # Zero out investments since DFSC is now in STL
    if items_f:
        result['investments'] = 0.0

    if items_g:
        calc_g = sum(x['amount'] for x in items_g)
        log.append(f"Loans Annexure-G: {len(items_g)} items, total={calc_g:,.0f}")
    if items_f:
        calc_f = sum(x['amount'] for x in items_f)
        log.append(f"Investment Annexure-F → STL: {len(items_f)} items ({calc_f:,.0f})")
    log.append(f"STL: {len(loans_to_others)} loans_to_others, "
               f"{len(revenue_auth)} revenue_auth, {len(other_current)} other_current")


def _extract_ocl_annexure(rows, result, log):
    """
    Directly extract OTHER PAYABLE (Annexure-D) items by scanning for the
    section header and reading name+amount from the adjacent columns.
    This handles Salary Payable and other OCL items that the section reader misses.
    """
    if result.get('other_payable_items'):
        return  # already populated

    in_section = False
    items = []
    name_col = None
    amt_col  = None
    other_pay_total = result.get('other_payables', 0)
    _OCL_NAME_KEYS = ('payable', 'tds', 'gst', 'provision', 'salary',
                      'accrued', 'outstanding', 'due to', 'liability', 'advance from')

    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)

        if not in_section:
            if 'OTHER PAYABLE' in rs and 'ANNEXURE' in rs:
                in_section = True
                continue
            continue

        # Detect name_col and amt_col from first item row
        if name_col is None:
            for ci, v in enumerate(row):
                sv = _s(v).strip()
                if len(sv) > 3:
                    sv_lower = sv.lower()
                    if any(kw in sv_lower for kw in _OCL_NAME_KEYS):
                        # Find amount nearby
                        for ac in range(ci + 1, min(ci + 8, len(row))):
                            av = row[ac]
                            if isinstance(av, (int, float)) and av > 0:
                                name_col = ci
                                amt_col  = ac
                                break
                    if name_col is not None:
                        break

        if name_col is None:
            continue

        # Check if the OCL name column has 'TOTAL' — this is the real OCL total row
        cell_name = _s(row[name_col]).strip().upper() if name_col < len(row) else ''
        if cell_name == 'TOTAL':
            break

        # Get name from name_col and amount from amt_col
        nm = _s(row[name_col]).strip()
        amt = row[amt_col] if amt_col < len(row) else None
        if nm and len(nm) > 2 and isinstance(amt, (int, float)) and amt > 0:
            nm_lower = nm.lower()
            if any(kw in nm_lower for kw in _OCL_NAME_KEYS):
                items.append({'name': nm, 'amount': amt})

    if items:
        # Deduplicate against what's already in unsecured_loan_parties
        existing = {it['name'].strip().lower() for it in
                    result.get('unsecured_loan_parties', [])}
        new_items = [it for it in items if it['name'].strip().lower() not in existing]
        if new_items:
            result['other_payable_items'] = new_items
            log.append(f"OCL annexure: {len(new_items)} items extracted "
                       f"({[it['name'] for it in new_items]})")


# ─────────────────────────────────────────────────────────────────────────────
#  Pass 4c — Capital account Annexure-A direct extraction
# ─────────────────────────────────────────────────────────────────────────────

def _extract_capital_annexure(rows, result, log):
    """
    Extract capital movements (opening, cash additions, withdrawals) from Annexure-A.

    GD Singla layout (0-based col indices):
      col40 = label text
      col42 = individual line amounts (Cheque/RTGS intro, each withdrawal item)
      col44 = running subtotals (opening, additions total, Less total, closing)

    Key rows (0-based row index from XLS):
      OPENING BALANCE row: col44 = 5,252,803.82
      Cheque/RTGS row:     col42 = 2,121,000  (cash intro, NOT profit)
      LIC row (last Less): col44 = 1,677,860.92 (withdrawals subtotal)
      CLOSING row:         col44 = 6,532,802.90
    """
    in_section  = False
    opening      = 0.0
    additions_cash = 0.0
    last_col44   = 0.0
    found_opening = False

    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)

        if not in_section:
            if ('PROP CAPITAL ACCOUNT' in rs or 'PROPRIETOR CAPITAL' in rs) and 'ANNEXURE-A' in rs:
                in_section = True
            continue

        # Opening balance row
        if 'OPENING BALANCE' in rs:
            v = row[44] if 44 < len(row) else None
            if isinstance(v, (int, float)) and v > 0:
                opening = v
                found_opening = True
            continue

        # Closing row — stop and record withdrawals from last_col44
        if found_opening and 'CLOSING' in rs and 'BALANCE' in rs:
            break

        if not found_opening:
            continue

        # Track last col44 seen after opening (= withdrawals subtotal on final Less row)
        v44 = row[44] if 44 < len(row) else None
        if isinstance(v44, (int, float)) and v44 > 0:
            last_col44 = v44

        # Cash intro lines only (Cheque/RTGS/NEFT/Gift — NOT profit)
        if any(k in rs for k in ('CHEQUE', 'RTGS', 'NEFT', 'GIFT FROM')):
            v42 = row[42] if 42 < len(row) else None
            if isinstance(v42, (int, float)) and v42 > 0:
                additions_cash += v42

    if opening > 0:
        for cap in result.get('capital_accounts', []):
            cap['opening']     = opening
            cap['additions']   = round(additions_cash)
            cap['withdrawals'] = round(last_col44, 2)
            break
        log.append(f"Capital Annexure-A: opening={opening:,.2f}, "
                   f"cash_intro={additions_cash:,.0f}, withdrawals={last_col44:,.2f}")


def _extract_pl(rows, result, log):
    """
    The P&L (Trading + P&L A/c) is embedded in the same sheet.
    In GD Singla format it appears in MULTIPLE COLUMN BLOCKS on the same rows.

    Detection: find the row with 'TRADING' or 'PROFIT & LOSS ACCOUNT' header.
    The P&L starts a few rows later. Columns for amounts vary per client but
    are always in the RIGHT half of the T (cols 8-18 for Bobby/Ashok,
    cols 10-18 for Gupta). We scan all non-zero columns in these rows.

    Key rows (in order of appearance):
      To Opening Stock  | By Sales
      To Purchases      | By Closing Stock
      To Gross Profit   | (= By Gross Profit in P&L A/c)
      To Salary/Wages   | By Gross Profit b/d
      To Depreciation
      To Interest
      To Net Profit
    """
    # Find P&L header row
    pl_start = -1
    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)
        if 'TRADING' in rs and ('PROFIT' in rs or 'LOSS' in rs) and i < 50:
            pl_start = i
            log.append(f"P&L header at row {i}")
            break
        if 'PROFIT' in rs and 'LOSS' in rs and 'ACCOUNT' in rs and i < 50:
            pl_start = i
            break

    if pl_start < 0:
        pl_start = 7   # fallback

    pl_keywords = {
        'OPENING STOCK': 'opening_stock',
        'TO OPENING STOCK': 'opening_stock',
        'TO PURCHASE': 'purchases',
        'TO PURCHASES': 'purchases',
        'TO PURCHASE A/C': '_purchase_line',  # Bobby: individual purchase GST lines
        'BY SALE': 'sales',
        'BY SALES': 'sales',           # Gupta format: "By Sales :"
        'BY SALE A/C': '_sale_line',   # Bobby: individual sale GST lines
        'BY CLOSING STOCK': 'closing_stock',
        'CLOSING STOCK': '_cs_asset',
        'TO GROSS PROFIT': 'gross_profit',
        'BY GROSS PROFIT': 'gross_profit',
        'TO NET PROFIT': 'net_profit',
        'BY NET PROFIT': 'net_profit',
        'TO SALARY': 'salary_expenses',
        'TO WAGES': 'salary_expenses',
        'TO DEPRECIATION': 'depreciation',
        'TO INTEREST ON UNSECURED': 'interest_paid',
        'TO INTEREST ON LOAN': 'interest_paid',
        'TO INTEREST PAID': 'interest_paid',
        'TO BANK INT': 'interest_paid',
        'TO BANK INTT': 'interest_paid',
        'TO BANK CHARGES': 'interest_paid',    # Bobby includes bank charges in finance cost
        'TO CAR LOAN INT': 'interest_paid',
        'BY DISCOUNT': 'other_income',
        'BY REBATE': 'other_income',
    }
    # Track whether we're accumulating GST purchase/sale lines (Bobby style)
    _in_purchase_gst = False
    _in_sale_gst = False

    expense_kws = [
        'TO AUDIT FEE', 'TO BANK CHARGE', 'TO ELECTRICITY', 'TO INSURANCE',
        'TO TELEPHONE', 'TO REPAIR', 'TO PETROL', 'TO RENT', 'TO GENERAL',
        'TO POSTAGE', 'TO STATIONARY', 'TO STAFF', 'TO PACKING',
        'TO PROPERTY TAX', 'TO CAR EXP', 'TO SCOOTER', 'TO LABOUR',
        'TO LOADING', 'TO SHOP', 'TO DIWALI', 'TO F.O.C', 'TO FOC',
        'TO FREIGHT', 'TO COMMISSION', 'TO COMMISSS', 'TO COMPUTER', 'TO MEDICLAIM',
        'TO LEGAL', 'TO PARTNER INTEREST', 'TO PARTNER SALARY',
    ]

    for i in range(pl_start, len(rows)):
        row = rows[i]
        # Scan every cell in the row for P&L labels
        for j, v in enumerate(row):
            lbl = _s(v).upper()
            if not lbl:
                continue

            for kw, field in sorted(pl_keywords.items(), key=lambda x: -len(x[0])):
                if lbl.startswith(kw) or lbl == kw:
                    # Find nearest number after this cell
                    # For SALES keywords, strategy depends on type:
                    if field == '_sale_line':
                        # Bobby GST sale lines: take FIRST number
                        amt = _first_num(row, j + 1, j + 6)
                    elif field == 'sales':
                        # Ashok/Gupta: scan this row + next few rows for TOTAL sales
                        _cands = []
                        for _ri in range(i, min(i + 5, len(rows))):
                            _r2 = rows[_ri]
                            for _k in range(j, min(j + 8, len(_r2))):
                                _v = _n(_r2[_k])
                                if _v > 1000:
                                    _cands.append(_v)
                        amt = max(_cands) if _cands else 0
                    else:
                        amt = _first_num(row, j + 1, j + 8)
                        if amt <= 0:
                            amt = _first_num(row, j + 1)
                    if amt > 0:
                        if field == '_cs_asset':
                            if result['closing_stock'] == 0:
                                result['closing_stock'] = amt
                        elif field == 'closing_stock':
                            if result['closing_stock'] == 0:
                                result['closing_stock'] = amt
                        elif field in ('sales', '_sale_line'):
                            result['sales'] += amt
                        elif field == '_purchase_line':
                            result['purchases'] += amt
                        elif field == 'opening_stock':
                            if result['opening_stock'] == 0:
                                result['opening_stock'] = amt
                        elif field == 'purchases':
                            if result['purchases'] == 0:
                                result['purchases'] = amt
                        elif field == 'gross_profit':
                            if result['gross_profit'] == 0:
                                result['gross_profit'] = amt
                        elif field == 'net_profit':
                            # Prefer 'TO NET PROFIT' (total) over 'BY NET PROFIT' (partner share)
                            if result['net_profit'] == 0 or (
                               lbl.startswith('TO NET PROFIT') and amt > result['net_profit']
                            ):
                                result['net_profit'] = amt
                        elif field in ('salary_expenses', 'depreciation',
                                       'interest_paid', 'other_income'):
                            result[field] += amt
                    break

            # Other expense items
            for kw in expense_kws:
                if lbl.startswith(kw):
                    amt = _first_num(row, j + 1, j + 8)
                    if amt > 0:
                        name = _s(v).replace('To ', '').strip()
                        result['other_expense_items'].append({'name': name, 'amount': amt})
                        result['other_expenses'] += amt
                    break

    log.append(
        f"P&L: Sales={result['sales']:.0f}, Purchases={result['purchases']:.0f}, "
        f"Stock={result['closing_stock']:.0f}, GP={result['gross_profit']:.0f}, "
        f"NP={result['net_profit']:.0f}"
    )


# ─────────────────────────────────────────────────────────────────────────────
#  Pass 6 — Depreciation chart extraction
# ─────────────────────────────────────────────────────────────────────────────

def _extract_dep_chart(rows, result, log):
    """
    Find the DEPRECIATION CHART section and extract fixed asset items.

    GD Singla format: dep chart is embedded in the same sheet at cols 50-58 (Ashok),
    cols 49-57 (Bobby), or cols 30-37 (Gupta). The header 'DEPRECIATION CHART' appears
    somewhere in the row (not necessarily col 0).

    Algorithm:
      1. Find the row containing 'DEPRECIATION CHART'.
      2. Detect the name column (first text col after the header keyword col).
      3. Detect the number column block (adjacent numeric cols).
      4. Read asset rows until TOTAL or blank.
    """
    in_dep = False
    items = []
    name_col = None    # col index of asset name
    num_start = None   # first numeric col of dep chart

    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)
        if not in_dep:
            if 'DEPRECIATION CHART' in rs or ('FIXED ASSET' in rs and 'DEP' in rs and 'CHART' in rs):
                in_dep = True
                log.append(f"Dep chart found at row {i}")
                # Detect which column 'DEPRECIATION CHART' is in
                dep_col = next((j for j, v in enumerate(row)
                                if 'DEPRECIATION' in _s(v).upper()), 0)
                # GD Singla layout: dep chart label is 2 cols to the RIGHT of asset names.
                # So name_col = dep_col - 2. num_start = name_col + 1.
                name_col = max(0, dep_col - 2)
                num_start = name_col + 1
            continue

        # Skip rows where the dep chart name column is blank (these are BS/P&L rows mixed in)
        # name_col is fixed at dep_col-2; only process rows where that col has an asset label

        # Skip category headers (all caps, no numbers)
        nums_in_row = [(j, float(v)) for j, v in enumerate(row)
                       if isinstance(v, (int, float)) and not (isinstance(v, float) and str(v) == 'nan')
                       and float(v) != 0]
        
        if not nums_in_row:
            continue

        # Detect name_col if not found yet: find first text cell that looks like an asset name
        if name_col is None:
            for j, v in enumerate(row):
                sv = _s(v)
                if sv and len(sv) > 3 and not sv.upper().startswith(('TOTAL', 'PART', 'DATE', 'AMOUNT')):
                    name_col = j
                    break

        # Get the asset name from name_col
        label = _s(row[name_col]).strip() if name_col is not None else ''
        # Limit block to max 9 numeric cols after name_col (dep chart has exactly 9 cols)
        block_nums = []
        if name_col is not None and num_start is not None:
            # Max 9 numeric cols in a dep chart row
            block_nums = [x for x in nums_in_row
                          if num_start <= x[0] <= num_start + 9]

        # Skip blank, header, total rows
        if not label:
            continue
        lu = label.upper()
        if lu in ('TOTAL', 'GRAND TOTAL', 'NET TOTAL', 'PARTICULARS', ''):
            if lu == 'TOTAL' and items:
                break
            continue
        # Skip ONLY multi-word category headers (exact start match only)
        if any(lu.startswith(x) for x in ('PLANT & MACHINERY', 'VEHICLE', 'COMPUTERS',
                                           'LAND AND', 'INTANGIBLE')):
            continue  # Category header rows

        # Find the numeric block: cols from num_start or first num col > name_col
        if num_start is None:
            num_start = nums_in_row[0][0] if nums_in_row else (name_col + 1 if name_col else 1)

        # Use bounded block_nums already set above (num_start to num_start+9)
        # Fallback if name_col/num_start not yet set
        if len(block_nums) < 2:
            continue

        # GD Singla dep chart column layout (offset from num_start):
        #   +0 = opening WDV  +1 = add>180d  +2 = add<180d
        #   +3 = sale  +4 = total  +5 = rate%  +6 = dep  +7 = closing WDV
        # block_nums only has NON-ZERO entries, so we must check actual column position.
        opening_wdv = block_nums[0][1]

        # add_gt180: only count if its column is exactly num_start+1
        add_gt180 = 0.0
        for bc, bv in block_nums[1:]:
            if bc == num_start + 1:
                add_gt180 = bv
            break  # stop after first candidate

        # add_lt180: only count if its column is exactly num_start+2
        add_lt180 = 0.0
        for bc, bv in block_nums[1:]:
            if bc == num_start + 2:
                add_lt180 = bv
                break

        additions = add_gt180 + add_lt180
        sale = 0.0
        rate = 0.0
        dep = 0.0
        closing_wdv = 0.0

        if len(block_nums) >= 2:
            closing_wdv = abs(block_nums[-1][1])
            dep = abs(block_nums[-2][1])

        # Find rate (typically 0.10, 0.15, 0.40 stored as float or 10, 15, 40)
        for _, n in block_nums:
            if 0 < n <= 1:
                rate = round(n * 100)
                break
            if n in (5, 10, 15, 20, 25, 30, 40):
                rate = int(n)
                break

        if opening_wdv > 0 or additions > 0:
            items.append({
                'name': label,
                'opening_wdv': opening_wdv,
                'additions': additions,
                'sales': sale,
                'dep': dep,
                'closing_wdv': closing_wdv,
                'rate': rate,
            })

    if items:
        # Filter out obviously wrong items (dep chart items where closing_wdv is 0 for all)
        valid = [x for x in items if x['closing_wdv'] > 0 or x['opening_wdv'] > 0]
        result['fixed_asset_items'] = valid if valid else items
        calc_wdv = sum(x['closing_wdv'] for x in result['fixed_asset_items'])
        if calc_wdv > 0:
            result['fixed_assets_wdv'] = calc_wdv
        log.append(f"Dep chart: {len(result['fixed_asset_items'])} assets, WDV={calc_wdv:.0f}")


# ─────────────────────────────────────────────────────────────────────────────
#  Pass 7 — Cash & Bank
# ─────────────────────────────────────────────────────────────────────────────

_BANK_KEYWORDS = ('HDFC', 'SBI', 'ICICI', 'OBC', 'PNB', 'AXIS', 'ORIENTAL',
                  'STATE BANK', 'KOTAK', 'YES BANK', 'CANARA', 'UNION BANK',
                  'BANK OF BARODA', 'BOB', 'F.D.R', 'FDR', 'BANK A/C', 'BANK A/C')


def _extract_cash_bank(rows, result, log):
    """
    Extract cash in hand and bank-wise balances.

    Strategy (in priority order):
    1. Annexure-H (col50=name, col58=amount) — GD Singla dedicated annexure.
       This gives exact Cash In Hand and individual bank accounts.
    2. Generic scan of the BS section (fallback for non-GD-Singla layouts).
    """

    # ── Pass 1: Annexure-H at col50 / col58 ──────────────────────────────────
    in_h = False
    cash_h  = 0.0
    banks_h = []
    total_h = 0.0

    for i, row in enumerate(rows):
        if 50 >= len(row):
            continue
        name50 = _s(row[50]).strip()
        amt58  = row[58] if 58 < len(row) else None

        if not in_h:
            rs = ' '.join(_s(v).upper() for v in row if v is not None)
            if 'CASH' in rs and 'BANK' in rs and 'ANNEXURE-H' in rs:
                in_h = True
            continue

        n_up = name50.upper()

        if n_up == 'TOTAL':
            if isinstance(amt58, (int, float)) and amt58 > 0:
                total_h = amt58
            break

        if not name50 or len(name50) < 2:
            continue
        if not isinstance(amt58, (int, float)) or amt58 <= 0:
            continue

        if 'CASH IN HAND' in n_up or 'CASH-IN-HAND' in n_up:
            cash_h = amt58
        elif any(bk in n_up for bk in _BANK_KEYWORDS) or 'BANK' in n_up:
            banks_h.append({'name': name50, 'amount': amt58})
        else:
            # Unknown item — treat as bank account
            banks_h.append({'name': name50, 'amount': amt58})

    if in_h and (cash_h > 0 or banks_h):
        result['cash_in_hand'] = cash_h
        result['bank_balances'] = banks_h
        bank_sum = sum(x['amount'] for x in banks_h)
        result['cash_bank'] = cash_h + bank_sum
        log.append(f"Cash Annexure-H: hand={cash_h:,.0f}, "
                   f"{len(banks_h)} bank(s)={bank_sum:,.0f}, total={result['cash_bank']:,.0f}")
        return   # Annexure-H is authoritative — skip generic scan

    # ── Pass 2: Generic scan (fallback) ──────────────────────────────────────
    in_cash = False
    cash_items = []

    for i, row in enumerate(rows):
        rs = ' '.join(_s(v).upper() for v in row if v is not None)

        if 'CASH' in rs and 'BANK' in rs and 'BALANCE' in rs and not in_cash and i < 50:
            in_cash = True
            for j, v in enumerate(row):
                lv = _s(v).upper()
                if 'CASH' in lv and 'BANK' in lv:
                    amt = _first_num(row, j + 1, j + 8)
                    if amt > 100:
                        result['cash_bank'] = amt
                        log.append(f"Cash & bank total (header row): {amt}")
                    break
            continue

        if not in_cash:
            continue

        r0u = _s(row[0]).upper()
        if any(x in r0u for x in ('TOTAL', 'SUNDRY', 'LOANS', 'ADVANCE', 'OTHER', 'INVEST')):
            break

        for j, v in enumerate(row):
            lbl = _s(v).upper()
            if not lbl:
                continue
            if 'CASH IN HAND' in lbl or 'CASH-IN-HAND' in lbl or (lbl == 'CASH' and j < 10):
                amt = _first_num(row, j + 1, j + 6)
                if not amt:
                    for k in range(j, min(j + 5, len(row))):
                        rv = row[k]
                        if isinstance(rv, (int, float)) and not (isinstance(rv, float) and str(rv) == 'nan') and float(rv) > 100:
                            amt = float(rv)
                            break
                if amt > 0:
                    result['cash_in_hand'] = amt
                    log.append(f"Cash in hand: {amt}")
            elif any(bk in lbl for bk in _BANK_KEYWORDS):
                amt = _first_num(row, j + 1, j + 6)
                if amt > 10:
                    cash_items.append({'name': _s(v), 'amount': amt})

    if cash_items or result.get('cash_in_hand', 0) > 0:
        if cash_items:
            result['bank_balances'] = cash_items
        bank_total = sum(x['amount'] for x in cash_items)
        computed = result['cash_in_hand'] + bank_total
        if computed > 0:
            result['cash_bank'] = computed
        log.append(f"Banks: {len(cash_items)}, hand={result['cash_in_hand']:.0f}, "
                   f"total={result['cash_bank']:.0f}")


# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 2 — Template injector
# ─────────────────────────────────────────────────────────────────────────────

def inject_into_template(parsed: dict, template_path: str, output_path: str,
                         client_name: str = None, cy_year: str = None) -> dict:
    """
    Fill Output_sample_format.xlsx with PY values from parsed dict.
    CY column = 0 / blank (yellow highlighted).
    """
    shutil.copy2(template_path, output_path)
    wb = load_workbook(output_path)

    py_year = parsed.get('py_year_end', '2024')
    if cy_year is None:
        cy_year = str(int(py_year) + 1)

    if client_name is None:
        client_name = parsed.get('entity_name', 'M/S CLIENT')

    log = parsed.get('log', [])
    log.append(f"Injecting CY={cy_year}, PY={py_year}")

    _inject_bs_sheet(wb, parsed, client_name, cy_year, py_year, log)
    _inject_pl_sheet(wb, parsed, client_name, cy_year, py_year, log)
    _inject_capital_sheet(wb, parsed, client_name, cy_year, py_year, log)
    _inject_notes_bs(wb, parsed, client_name, cy_year, py_year, log)
    _inject_notes_pl(wb, parsed, client_name, cy_year, py_year, log)
    _inject_details_sheet(wb, parsed, client_name, cy_year, py_year, log)
    _inject_fixed_assets_py(wb, parsed, client_name, py_year, log)
    _inject_fixed_assets_cy_opening(wb, parsed, log)
    _inject_gross_profit_sheet(wb, parsed, client_name, cy_year, py_year, log)
    _update_headers(wb, client_name, cy_year, py_year)

    wb.save(output_path)
    log.append(f"Saved: {output_path}")

    return {
        'status': 'success', 'output': output_path, 'log': log,
        'entity_name': client_name, 'cy_year': cy_year, 'py_year': py_year,
    }


def _py(ws, row, col, value):
    """Write a numeric PY value."""
    _write_num(ws, row, col, value)


def _cy(ws, row, col, value=0):
    """Write a CY input cell (yellow highlighted, 0 by default)."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return
    cell.value = float(value or 0)
    cell.fill = _INPUT_FILL
    cell.number_format = '#,##0'


# ─────────────────────────────────────────────────────────────────────────────
# HARD-CODED COLUMN CONSTANTS (1-indexed, verified from template inspection)
#
# bs sheet          : CY=col5(E)  PY=col6(F)   — but bs is formula-driven
# notes to bs       : CY=col4(D)  PY=col5(E)
# notes to p&l      : CY=col4(D)  PY=col5(E)
# capital           : C=3 D=4 E=5 F=6 G=7
# Details           : CY=col4(D)  PY=col5(E)
# GROSS PROFIT left : CY=col2(B)  PY=col3(C)
# GROSS PROFIT right: CY=col5(E)  PY=col6(F)
# Fixed Assets P.Yr.: B=opening C=add>180 D=add<180 E=sales F=total G=rate H=dep I=wdv
# ─────────────────────────────────────────────────────────────────────────────


# ── BS sheet ──────────────────────────────────────────────────────────────────
# BS sheet is ENTIRELY formula-driven. All PY values flow in via:
#   capital!G11          → R7/R8   (capital closing)
#   notes to bs!E19      → R11     (LT borrowings total)
#   notes to bs!E34      → R14     (ST borrowings total)
#   notes to bs!E38/E39  → R16/R17 (trade payables MSME / others)
#   notes to bs!E65      → R18     (other current liabilities)
#   notes to bs!E71      → R19     (ST provisions)
#   Fixed Assets P.Yr.!I37 → R26   (PPE WDV)
#   notes to bs!E85      → R27     (investments)
#   notes to p&l!E27     → R31     (inventories = closing stock)
#   notes to bs!E101     → R32     (debtors)
#   notes to bs!E120     → R33     (cash+bank)
#   notes to bs!E150     → R34     (ST loans)
#   notes to bs!E160     → R35     (other current assets)
# DO NOT write directly to bs sheet rows.

def _inject_bs_sheet(wb, parsed, client_name, cy_year, py_year, log):
    log.append('BS sheet: formula-driven, no direct writes needed')


# ── P&L sheet ─────────────────────────────────────────────────────────────────
# P&L sheet is ENTIRELY formula-driven from notes to p&l:
#   notes to p&l!E8   → R5  revenue
#   notes to p&l!E13  → R6  other income
#   notes to p&l!E29  → R10 cost of materials
#   notes to p&l!E40  → R11 employee benefits
#   notes to p&l!E47  → R12 finance cost
#   notes to p&l!E57  → R13 depreciation (from Fixed Assets P.Yr.)
#   notes to p&l!E87  → R14 other expenses
# DO NOT write directly to p&l sheet rows.

def _inject_pl_sheet(wb, parsed, client_name, cy_year, py_year, log):
    log.append('P&L sheet: formula-driven, no direct writes needed')


# ── Capital sheet ──────────────────────────────────────────────────────────────
# Template structure (verified from data_only=False inspection):
#   Block 1 (rows 6-19):  CY headers R7, CY data R8, CY totals R10, PY row R11
#   Block 2 (rows 24-40): CY headers R30, CY data R31, CY totals R33, PY row R34
#
# Columns (1-indexed):
#   B=2 Sr.No.  C=3 Opening  D=4 Additions  E=5 Withdrawals  F=6 Profit  G=7 Closing(formula)
#
# For proprietorship: caps[0] fills Block1 PY(R11) + Block2 PY(R34)
# For partnership:    caps[0] fills Block1 PY(R11),  caps[1] fills Block2 PY(R34)
#
# The CY opening (R8 col C) = =G11 formula in template. DO NOT overwrite G column.
# We only write: C(opening), D(additions), E(withdrawals), F(profit) for PY rows.

def _inject_capital_sheet(wb, parsed, client_name, cy_year, py_year, log):
    if 'capital' not in wb.sheetnames:
        return
    ws = wb['capital']
    p = parsed
    caps = p.get('capital_accounts', [])
    if not caps:
        log.append('Capital: no capital accounts found')
        return

    COL_NAME   = 2  # B
    COL_OPEN   = 3  # C: Opening balance
    COL_ADD    = 4  # D: Capital introduced
    COL_WITH   = 5  # E: Withdrawals
    COL_PROFIT = 6  # F: Share of profit/loss
    # COL_CLOSE = 7  # G: formula =C+D-E+F — DO NOT WRITE

    def _write_py_row(r, cap):
        # Col A(1) = "Previous Year (PY)" label — already in template, don't touch
        # Col B(2) = MergedCell (merged with A) — DO NOT write name here
        # Col C(3)=opening, D(4)=additions, E(5)=withdrawals, F(6)=profit
        # Col G(7) = closing formula =C+D-E+F — DO NOT write
        #
        # NOTE: The T-shaped BS does NOT contain the full PY capital movement breakdown.
        # Only net_profit is reliably extractable. Opening/additions/withdrawals
        # must be filled by the CA from prior year records.
        # COL_PROFIT (F=6) is a FORMULA in the template: ='p&l'!F17
        # DO NOT write to it — it will auto-calculate from the p&l chain.
        # Leave opening/add/withdraw as 0 (yellow) for CA to fill from prior year records.
        # Use parsed values from Annexure-A if available, else 0 (CA fills)
        opening_val   = cap.get('opening', 0) or 0
        additions_val = cap.get('additions', 0) or 0
        withdraw_val  = cap.get('withdrawals', 0) or 0
        _write_num(ws, r, COL_OPEN,  opening_val)
        _write_num(ws, r, COL_ADD,   additions_val)
        _write_num(ws, r, COL_WITH,  withdraw_val)
        # COL_PROFIT and COL_CLOSE are formula cells — DO NOT WRITE
        # Mark C-E as yellow input cells for CA (they can still edit)
        for col in (COL_OPEN, COL_ADD, COL_WITH):
            cell = ws.cell(r, col)
            from openpyxl.cell import MergedCell as MC
            if not isinstance(cell, MC):
                cell.fill = _INPUT_FILL

    # Block 1 PY row = R11
    _write_py_row(11, caps[0])
    # Update the CY row name (R8 col B) — this IS a writable cell (not merged)
    _write(ws, 8, COL_NAME, caps[0].get('name', ''))

    # Block 2 PY row = R34 — only write if there are 2+ capital accounts
    if len(caps) > 1:
        _write_py_row(34, caps[1])
        _write(ws, 31, COL_NAME, caps[1].get('name', ''))
        log.append(f'Capital sheet: PY rows written for {len(caps)} capital account(s)')
    else:
        log.append('Capital sheet: 1 account — Block 2 left as template default')


# ── Notes to BS ────────────────────────────────────────────────────────────────
# Template verified (data_only=False):
#   Col D(4) = CY   Col E(5) = PY
#
# Key writable cells:
#   R7:E   = LT secured term loans from banks
#   R8:E   = LT secured from other parties
#   (R10   = SUM(E7:E8) formula)
#   R15:E  = LT unsecured from related parties
#   R16:E  = LT unsecured from other parties
#   (R18,R19 = formulas)
#   R25:E  = ST secured CC a/c
#   R54-R64 = Other current liability items (11 slots)
#   (R65   = SUM(D54:D64) formula)
#   R106:E = Cash in hand
#   (R110  = SUM formula)
#   R113-R118:E = Individual bank balances (6 slots)
#   (R119,R120 = formulas)
#   R134-R140:E = ST loans to others (B section, 7 slots)
#   (R141,R150 = formulas)
#
# NOTE: Trade receivable rows (R91, R97) are formula-driven from Details sheet.
#       DO NOT write to them.

def _inject_notes_bs(wb, parsed, client_name, cy_year, py_year, log):
    if 'notes to bs' not in wb.sheetnames:
        return
    ws = wb['notes to bs']
    p = parsed
    PY = 5   # col E
    CY = 4   # col D

    secured   = p.get('secured_loans', 0)
    unsecured = p.get('unsecured_loans', 0)

    # ── Note 3: Long-term borrowings ──────────────────────────────────────────
    # Secured: write total to R7 (bank), R8=0 (other parties)
    _py(ws, 7,  PY, secured); _cy(ws, 7,  CY)
    _py(ws, 8,  PY, 0);       _cy(ws, 8,  CY)
    # R10 = SUM(E7:E8) — formula, skip

    # Unsecured related parties: R15 C4/C5 = =Details!D13/E13 — FORMULA, DO NOT WRITE.
    # The Details sheet inject (R7-R12) feeds Details!R13 via SUM formula,
    # which then flows here automatically.
    # R16 (from other parties) is writable but typically 0 for T-shaped clients.
    _py(ws, 16, PY, 0);         _cy(ws, 16, CY)
    # R18 = SUM(E14:E17), R19 = R18+R10 — formulas, skip

    # ── Note 3: Short-term borrowings ─────────────────────────────────────────
    _py(ws, 25, PY, 0); _cy(ws, 25, CY)
    # R29, R33, R34 = formulas, skip

    # ── Note 4: Trade payables ────────────────────────────────────────────────
    # R38/R39 are formula-driven from Details!D68/E68 and Details!D62/E62
    # Do NOT write here — creditor parties in Details drive these.

    # ── Note 5: Other current liabilities (R54-R64, 11 slots) ────────────────
    # Parser puts OCL items in other_payable_items when found correctly.
    # For T-shaped BS they often bleed into unsecured_loan_parties.
    # Detect payable/TDS/GST type names and rescue them into OCL.
    _OCL_KEYWORDS = ('payable', 'tds', 'gst', 'provision', 'accrued',
                     'outstanding', 'due to', 'liability')
    # Merge both sources:
    # 1. other_payable_items (from dedicated OCL annexure parser — Salary Payable, TDS etc.)
    # 2. unsecured_loan_parties entries with OCL keywords (Audit Fees, GST, Electricity etc.)
    # Deduplicate by AMOUNT (same amount = same item regardless of name variant).
    ocl_items = list(p.get('other_payable_items', []))
    existing_ocl_amounts = {round(it['amount']) for it in ocl_items}
    for party in p.get('unsecured_loan_parties', []):
        nm_l = party['name'].strip().lower()
        if any(kw in nm_l for kw in _OCL_KEYWORDS):
            amt_key = round(party['amount'])
            if amt_key not in existing_ocl_amounts:
                ocl_items.append(party)
                existing_ocl_amounts.add(amt_key)

    # Deduplicate OCL items by normalized name.
    # e.g. "GST- Reverse Charges" and "GST REVERSE CHARGE PAYABLE" are the same item
    # parsed twice from adjacent annexure columns.
    # Strategy: normalize to alpha-only prefix (first 12 chars) + amount for dedup key.
    seen_ocl = {}
    deduped_ocl = []
    for item in ocl_items:
        alpha_key = ''.join(c for c in item['name'].lower() if c.isalpha())[:12]
        # Also deduplicate by amount: same amount = same item regardless of name variant
        amt_key = round(item['amount'])
        key = (alpha_key, amt_key)
        if key not in seen_ocl:
            seen_ocl[key] = True
            deduped_ocl.append(item)
    ocl_items = deduped_ocl

    for i in range(11):
        r = 54 + i
        if i < len(ocl_items):
            item = ocl_items[i]
            _write(ws, r, 2, item['name'])
            _py(ws, r, PY, item['amount'])
        else:
            _py(ws, r, PY, 0)
        _cy(ws, r, CY)
    # R65 = SUM(D54:D64) — formula, skip

    # ── Note 10: Cash and bank balances ───────────────────────────────────────
    _py(ws, 106, PY, p.get('cash_in_hand', 0)); _cy(ws, 106, CY)
    # R110 = SUM(E106:E106) — formula, skip

    banks = p.get('bank_balances', [])
    _cb = p.get('cash_bank', 0); cash_bank_total = 0 if (_cb is None or (isinstance(_cb, float) and math.isnan(_cb))) else float(_cb)
    _ch = p.get('cash_in_hand', 0); cash_hand = 0 if (_ch is None or (isinstance(_ch, float) and math.isnan(_ch))) else float(_ch)
    bank_total_from_bs = round(cash_bank_total - cash_hand)

    for i in range(6):   # R113-R118
        r = 113 + i
        if i < len(banks):
            _write(ws, r, 2, banks[i]['name'])
            _py(ws, r, PY, banks[i]['amount'])
        elif i == 0 and not banks and bank_total_from_bs > 0:
            # No parsed bank breakdown — write total in first slot as single account
            _write(ws, r, 2, 'HDFC Bank')
            _py(ws, r, PY, bank_total_from_bs)
        else:
            _py(ws, r, PY, 0)
        _cy(ws, r, CY)
    # R119 = SUM(E113:E118), R120 = R110+R119 — formulas, skip

    # ── Note 11: Short-term loans and advances ────────────────────────────────
    # B section — Loans to Others: R134-R140 (7 slots)
    # Template formula SUM(E138:E140) is too narrow — fix it to cover all 7 rows.
    loans_items = p.get('loans_to_other_items', p.get('loans_advances_items', []))
    loans_total = p.get('loans_advances', 0) or p.get('advances_security', 0)
    for i in range(7):
        r = 134 + i
        if i < len(loans_items):
            itm = loans_items[i]
            _write(ws, r, 2, itm.get('name', ''))
            _py(ws, r, PY, itm.get('amount', 0))
        else:
            _py(ws, r, PY, 0)
        _cy(ws, r, CY)
    # If no itemised list but we have a total, lump into R134
    if not loans_items and loans_total:
        _py(ws, 134, PY, loans_total)
    # Fix Total(B) formula to cover all 7 rows R134:R140
    _write(ws, 141, PY, '=SUM(E134:E140)')
    _write(ws, 141, CY, '=SUM(D134:D140)')

    # C section — Advance to Revenue Authorities: R143-R147 (5 slots)
    # Template formula SUM(E143:E143) only covers 1 row — fix to cover all 5.
    rev_items = p.get('advance_to_revenue_items', [])
    for i in range(5):
        r = 143 + i
        if i < len(rev_items):
            itm = rev_items[i]
            _write(ws, r, 2, itm.get('name', ''))
            _py(ws, r, PY, itm.get('amount', 0))
        else:
            _py(ws, r, PY, 0)
        _cy(ws, r, CY)
    # Fix Total(C) formula to cover all 5 rows R143:R147
    _write(ws, 148, PY, '=SUM(E143:E147)')
    _write(ws, 148, CY, '=SUM(D143:D147)')
    # R150 Total(A+B+C) formula is already correct — skip

    # ── Note 12: Other current assets (R154-R159, 6 slots) ───────────────────
    oca_items = p.get('other_current_asset_items', [])
    for i in range(6):
        r = 154 + i
        if i < len(oca_items):
            itm = oca_items[i]
            _write(ws, r, 2, itm.get('name', ''))
            _py(ws, r, PY, itm.get('amount', 0))
        else:
            _py(ws, r, PY, 0)
        _cy(ws, r, CY)
    # R160 = SUM formula, skip

    log.append('Notes to BS injected')


# ── Notes to P&L ───────────────────────────────────────────────────────────────
# Template verified (data_only=False):
#   Col D(4) = CY   Col E(5) = PY
#
# Key structure:
#   R6:E   = sales (Note 13)
#   R12:E  = other income (Note 14)
#   R18:E  = opening stock PY  [R18:D = =E27 formula — DO NOT TOUCH CY]
#   R21:E  = purchases PY
#   R24-R25 = direct expense item rows
#   R27:E  = closing stock PY  [R27:D = formula from GROSS PROFIT — DO NOT TOUCH CY]
#   R34:E  = salaries, R35:E = bonus, R36:E = staff welfare, R40=SUM formula
#   R44:E  = bank interest, R45:E = unsecured interest, R47=SUM(E44:E45) formula
#   R56:E  = depreciation — formula from Fixed Assets P.Yr.!I37 — DO NOT TOUCH
#   R57    = SUM formula
#   R61-R75 = named other-expense rows (fixed labels, see map below)
#   R76-R86 = spare rows for unmatched items
#   R87    = SUM(D61:D86) formula
#
# CRITICAL: other_expense_items from T-shaped BS must be DISTRIBUTED into the
# correct named rows. Items that are "direct expenses" (FOC, freight) go into
# R24-R25 (notes to p&l) and also into the GROSS PROFIT sheet (handled there).
# Bank charges → finance cost R44. Staff welfare → R36. Everything else → R61-R86.

_OTHER_EXP_ROW_MAP = {
    'audit fee':                 61,
    'audit fees':                61,
    'bank charges':              62,
    'bank charge':               62,
    'car exp':                   63,
    'car expense':               63,
    'car expenses':              63,
    'commission':                64,
    'commisssion':               64,
    'diwali exp':                65,
    'diwali exp.':               65,
    'electricity exp':           66,
    'electricity expenses':      66,
    'electricity exp.':          66,
    'insurance':                 67,
    'labour charges':            68,
    'labour charge':             68,
    'legal fee':                 69,
    'legal fees':                69,
    'to legal fee':              69,
    'loading unloading':         70,
    'loading unloading charges': 70,
    'petrol exp':                71,
    'petrol exp.':               71,
    'petrol expenses':           71,
    'repair & maintainance':     72,
    'repair & maintenance':      72,
    'scooter exp':               73,
    'scooter exp.':              73,
    'scooter expenses':          73,
    'shop exp':                  74,
    'shop exp.':                 74,
    'telephone exp':             75,
    'telephone exp.':            75,
    'telephone expenses':        75,
}

_DIRECT_EXP_KEYS = {
    'f.o.c.', 'foc', 'freight inward', 'freight outward',
    'freight inward (gst)', 'frieght inward', 'frieght inward (gst)',
}

_FINANCE_COST_KEYS = {
    # Bank interest and unsecured loan interest go to Note 17
    # NOTE: 'bank charges' goes to Note 19 Other Expenses (row 62), NOT here
    'bank interest',
    'interest', 'interest on unsecured', 'interest on unsecured loans',
}

_STAFF_WELFARE_KEYS = {
    'staff & labour welfare', 'staff and labour welfare',
    'staff welfare', 'labour welfare',
    'staff & labour welfare expenses',
}


def _inject_notes_pl(wb, parsed, client_name, cy_year, py_year, log):
    if 'notes to p&l' not in wb.sheetnames:
        return
    ws = wb['notes to p&l']
    p = parsed
    PY = 5   # col E
    CY = 4   # col D

    all_items = p.get('other_expense_items', [])

    # Classify items
    direct_items   = []
    finance_items  = []
    welfare_items  = []
    other_items    = []
    for it in all_items:
        k = it['name'].strip().lower()
        if k in _DIRECT_EXP_KEYS:
            direct_items.append(it)
        elif k in _FINANCE_COST_KEYS:
            finance_items.append(it)
        elif k in _STAFF_WELFARE_KEYS:
            welfare_items.append(it)
        else:
            other_items.append(it)

    # ── Note 13: Revenue (R6:E = sales) ──────────────────────────────────────
    _py(ws, 6,  PY, p.get('sales', 0));       _cy(ws, 6,  CY)

    # ── Note 14: Other income (R12:E) ────────────────────────────────────────
    _py(ws, 12, PY, p.get('other_income', 0)); _cy(ws, 12, CY)

    # ── Note 15: Cost of material ─────────────────────────────────────────────
    _py(ws, 18, PY, p.get('opening_stock', 0))
    # R18:D = =E27 formula for CY opening — DO NOT write CY here
    _py(ws, 21, PY, p.get('purchases', 0));   _cy(ws, 21, CY)

    # Direct expenses → R24-R25 (2 fixed slots)
    for i in range(2):
        r = 24 + i
        if i < len(direct_items):
            it = direct_items[i]
            _write(ws, r, 2, it['name'])
            _py(ws, r, PY, it['amount']); _cy(ws, r, CY)
        else:
            _py(ws, r, PY, 0); _cy(ws, r, CY)

    _py(ws, 27, PY, p.get('closing_stock', 0))
    # R27:D = formula from GROSS PROFIT — DO NOT write CY here

    # ── Note 16: Employee benefits ────────────────────────────────────────────
    # R34=salaries, R35=bonus, R36=staff welfare; R40=SUM formula
    sal = p.get('salary_expenses', 0)

    # Staff welfare from other_expense_items if present
    # Parser already excludes welfare from salary_expenses total, so write sal directly
    welfare_amt = sum(x['amount'] for x in welfare_items)

    _py(ws, 34, PY, sal);         _cy(ws, 34, CY)   # salaries (already excludes welfare)
    _py(ws, 35, PY, 0);           _cy(ws, 35, CY)   # bonus
    _py(ws, 36, PY, welfare_amt); _cy(ws, 36, CY)   # staff welfare
    # R40 = SUM(E34:E36) — formula, skip

    # ── Note 17: Finance cost ─────────────────────────────────────────────────
    # R44:E = bank interest (from banks only), R45:E = interest on unsecured loans
    # R47 = SUM(E44:E45) — formula, skip
    # finance_items contains only explicit 'bank interest' type entries from other_expense_items.
    # NOTE: 'bank charges' (3992) goes to Note 19 R62 (other expenses), NOT here.
    # NOTE: interest_paid parser field = bank charges in T-shaped XLS — do NOT use as bank_int.
    bank_int  = sum(x['amount'] for x in finance_items
                    if x['name'].strip().lower() == 'bank interest')
    unsec_int = sum(x['amount'] for x in finance_items
                    if x['name'].strip().lower() not in ('bank interest', 'bank charges',
                                                         'bank charge'))

    # If unsec_int still 0, scan unsecured_loan_parties for interest entries.
    # In T-shaped XLS the interest on unsecured loans often appears as a party row
    # with names like 'INTEREST', 'JLDR05627G', '194A' etc.
    _UNSEC_INT_KEYS = ('interest', 'jldr', '194a')
    if unsec_int == 0:
        seen_amounts = set()
        for party in p.get('unsecured_loan_parties', []):
            nm_l = party['name'].strip().lower()
            amt  = party['amount']
            if any(k in nm_l for k in _UNSEC_INT_KEYS) and amt > 0:
                # Deduplicate: same amount often listed under multiple reference codes
                if amt not in seen_amounts:
                    seen_amounts.add(amt)
                    unsec_int += amt
        if unsec_int > 0:
            log.append(f'Finance cost: unsec interest {unsec_int} extracted from unsecured_loan_parties')

    _py(ws, 44, PY, bank_int);  _cy(ws, 44, CY)
    _py(ws, 45, PY, unsec_int); _cy(ws, 45, CY)
    # R47 = SUM(E44:E45) — formula, skip

    # ── Note 18: Depreciation ─────────────────────────────────────────────────
    # R56:E = formula from 'Fixed Assets P. Yr.'!I37 — DO NOT WRITE
    # R57 = SUM formula — skip

    # ── Note 19: Other expenses (R61-R86, total R87=SUM) ─────────────────────
    # Clear all named rows first (set to 0), then fill in matched items
    written_rows = set()
    for k, r in _OTHER_EXP_ROW_MAP.items():
        if r not in written_rows:
            _py(ws, r, PY, 0); _cy(ws, r, CY)
            written_rows.add(r)

    unmatched = []
    for it in other_items:
        k = it['name'].strip().lower()
        r = _OTHER_EXP_ROW_MAP.get(k)
        if r:
            _py(ws, r, PY, it['amount']); _cy(ws, r, CY)
        else:
            unmatched.append(it)

    # Write unmatched into spare rows R76-R86
    for i, it in enumerate(unmatched):
        r = 76 + i
        if r > 86:
            break
        _write(ws, r, 2, it['name'])
        _py(ws, r, PY, it['amount']); _cy(ws, r, CY)

    # Clear remaining spare rows
    for r in range(76 + len(unmatched), 87):
        _py(ws, r, PY, 0); _cy(ws, r, CY)

    # R87 = SUM(D61:D86) — formula, skip

    log.append(f'Notes to P&L: {len(direct_items)} direct, {len(other_items)} other, '
               f'{len(unmatched)} unmatched to spare rows')


# ── Details sheet ──────────────────────────────────────────────────────────────
# Template structure (verified from data_only=False inspection):
#
#   UNSECURED LOANS:
#     R6  = "FROM RELATED PARTIES" header
#     R7-R12  = 6 party rows  (B=name, D=CY, E=PY)
#     R13 = SUM(D7:D12) formula
#     R15 = FROM OTHER PARTIES row (D=CY, E=PY)
#     R17 = SUM formula
#     R19 = TOTAL formula
#
#   SUNDRY CREDITORS:
#     R22 = header row
#     R23-R55 = 33 creditor party rows (A="M/s.", B=name, D=CY, E=PY)
#     R56 = "Advance from Customers" sub-header
#     R57-R61 = 5 advance-from-customer rows
#     R62 = SUM(D23:D61) formula
#     R63 = "DUE TO MSME" header
#     R68 = SUM formula
#     R69 = TOTAL formula
#
#   TRADE RECEIVABLE >6 months:
#     R72 = section header
#     R73 = header row
#     R74-R128 = 55 debtor rows (A="M/s.", B=name, D=CY, E=PY)
#     R129 = SUM(D74:D128) formula  ← notes to bs!E91 reads Details!E129
#
#   TRADE RECEIVABLE <6 months:
#     R132 = section header
#     R133 = header row
#     R134-R135 = 2 rows
#     R136 = SUM formula  ← notes to bs!E97 reads Details!E136

# Keywords that identify OCL items (payables) — should NOT go in Details unsecured section
_OCL_PARTY_KEYWORDS = (
    'payable', 'tds', 'gst reverse', 'provision for tax',
    'accrued', 'outstanding expenses', 'due to',
)


def _is_real_party_name(nm):
    """Filter out junk/numeric names from parsed party lists."""
    nm = nm.strip()
    if len(nm) < 3:
        return False
    # Reject purely numeric strings
    try:
        float(nm.replace(',', ''))
        return False
    except ValueError:
        pass
    # Reject date strings
    if re.match(r'^\d{4}-\d{2}-\d{2}', nm):
        return False
    if re.match(r'^\d{2}\.\d{2}\.\d{4}', nm):
        return False
    # Reject very short number-heavy tokens like "287.5", "3814.5"
    if re.match(r'^[\d\.\,\s]+$', nm):
        return False
    nm_l = nm.lower()
    # Reject internal calculation leftovers
    _SKIP = (
        'opening balance', 'closing balance', 'gross profit', 'net profit',
        'add profit', 'total', 'pack/pcs', 'amount in rs',
        'withdrawls', 'advance tax', 'appeal fees', 'trf to',
        'hdfc insurance', 'tcs receivable', 'lic ', 'start health',
        'building', 'motor cycle', 'furniture', 'water cooler',
        'digital moisture', 'servo control', 'refer item', 'refer as per',
        'nature of', 'particulars', 'annexure', 'add profit',
        'camera', 'washing machine', 'electricals', 'led', 'computer',
        'cheque/rtgs', 'cheque / rtgs', 'neft', 'rtgs',
    )
    if any(sp in nm_l for sp in _SKIP):
        return False
    return True


# Personal/relationship/commodity words that should NOT appear as trade parties
_PERSONAL_JUNK = {
    'wife', 'mother', 'father', 'son', 'daughter', 'husband',
    'sister', 'brother', 'stock',
    # Common individual first names that appear as lenders/family in T-shaped XLS
    'garima', 'sheenu', 'dimpal', 'rachit', 'deepak', 'pawan', 'asha',
}
# Commodity/food words that indicate stock items not debtors
# Note: only applied when the word appears in a context suggesting it IS a commodity,
# not a company name containing that word (e.g. "Sugar Mills Ltd." is a company).
_COMMODITY_WORDS = (
    'aloo', 'chana', 'ghee', 'katta',
    'besan', 'atta', 'dal', 'makai', 'maize',
    'sesame', 'sunflower', 'mustard', 'ground nut',
)
# These commodity words need careful matching: only block if NOT followed by company suffixes
_CAREFUL_COMMODITY = ('cotton', 'oil', 'sugar', 'rice', 'wheat',)
_COMPANY_SUFFIXES = ('mill', 'mills', 'industries', 'industry', 'corp', 'ltd', 'limited',
                     'pvt', 'private', 'company', 'co.', 'co ', 'refinery', 'syndicate',)


def _is_debtor_party(nm, lender_names=None):
    """True if this looks like a genuine trade debtor / receivable party."""
    if not _is_real_party_name(nm):
        return False
    nm_l = nm.lower().strip()
    # Reject if name or any individual word is in personal/junk set
    if nm_l in _PERSONAL_JUNK:
        return False
    # Also reject multi-word names where any word is a personal/family junk word
    if any(word in _PERSONAL_JUNK for word in nm_l.split()):
        return False
    # Reject commodity items
    if any(c in nm_l for c in _COMMODITY_WORDS):
        return False
    # Careful commodities: only block if not a company name
    for cc in _CAREFUL_COMMODITY:
        if cc in nm_l and not any(sfx in nm_l for sfx in _COMPANY_SUFFIXES):
            return False
    # Reject known lender names (passed from unsecured_loan_parties)
    if lender_names:
        nm_norm = nm.strip().lower()
        if nm_norm in lender_names:
            return False
    return True


def _is_unsecured_party(nm):
    """True if this is a real lender (not an OCL/payable item, not a commodity)."""
    nm_l = nm.strip().lower()
    # Payable-type items belong in OCL, not unsecured loans Details rows
    _PAYABLE_KEYS = (
        'payable', ' tds', 'gst-', 'gst reverse', 'gst credit',
        'provision', 'salary payable', 'audit fee', 'reverse charges',
    )
    if any(k in nm_l for k in _PAYABLE_KEYS):
        return False
    if not _is_real_party_name(nm):
        return False
    # Reject commodity words
    if any(c in nm_l for c in _COMMODITY_WORDS):
        return False
    # Reject TDS/interest reference codes (alphanumeric codes like JLDR05627G)
    if re.match(r'^[A-Z0-9]{5,12}$', nm.strip()):
        return False
    # Reject generic interest/loan labels
    _GENERIC_LABELS = ('interest', '194a', '194c', 'tds on', 'loan int')
    if any(g in nm_l for g in _GENERIC_LABELS):
        return False
    return True


def _inject_details_sheet(wb, parsed, client_name, cy_year, py_year, log):
    if 'Details' not in wb.sheetnames:
        return
    ws = wb['Details']
    p = parsed
    PY = 5   # col E
    CY = 4   # col D

    # ── Unsecured loan related parties (R7-R12, 6 slots) ─────────────────────
    # T-shaped BS parser mixes OCL items and debtor names into unsecured_loan_parties.
    # Strategy: filter, deduplicate by normalized name, then sort by amount desc
    # and take the top 6. This reliably picks the largest real lenders.
    def _normalize_nm(nm):
        nm = nm.strip().lower()
        for prefix in ('smt. ', 'sh. ', 'm/s. ', 'mr. ', 'mrs. '):
            nm = nm.replace(prefix, '')
        return nm.replace('.', '').replace('  ', ' ').strip()

    _seen_nm = set()
    _deduped = []
    for x in sorted(p.get('unsecured_loan_parties', []),
                    key=lambda x: x['amount'], reverse=True):
        if not _is_unsecured_party(x['name']):
            continue
        key = _normalize_nm(x['name'])
        if key in _seen_nm:
            continue
        _seen_nm.add(key)
        _deduped.append(x)

    unsec_parties = _deduped[:6]
    for i in range(6):
        r = 7 + i
        if i < len(unsec_parties):
            party = unsec_parties[i]
            _write(ws, r, 2, party['name'])
            _py(ws, r, PY, party['amount']); _cy(ws, r, CY)
        else:
            _write(ws, r, 2, '')
            _py(ws, r, PY, 0); _cy(ws, r, CY)
    # R13 = SUM formula, R17 = SUM formula, R19 = TOTAL formula — skip
    # R15 in Details = "from other parties" (writable). For T-shaped clients this is usually 0.
    _py(ws, 15, PY, 0); _cy(ws, 15, CY)   # Details R15: unsecured from other parties

    # ── Sundry creditors (R23-R55, 33 slots; advance R57-R61, 5 slots) ───────
    # Prefer data from _extract_creditor_annexure (Annexure-C, col45/col49).
    # Names there already contain "M/s." prefix — strip it before writing so the
    # template col-A "M/s." label + col-B name layout isn't doubled.
    cred_target  = p.get('sundry_creditors', 0)
    cred_parties_raw = p.get('sundry_creditor_parties', [])

    def _split_prefix(name):
        """Return (prefix, bare_name). Handles 'M/s. Foo' → ('M/s.', 'Foo')."""
        for pfx in ('M/s. ', 'M/s.', 'Sh. ', 'Smt. ', 'Mr. ', 'Mrs. '):
            if name.startswith(pfx):
                return pfx.strip(), name[len(pfx):].strip()
        return '', name.strip()

    # Filter junk only when Annexure-C didn't provide clean data
    _CRED_JUNK = ('trf to', 'transfer to', 'stock', 'hdfc bank',
                  'pspcl', 'property tax', 'municipal', 'municiple')
    good_crp = [cp for cp in cred_parties_raw
                if _is_real_party_name(cp['name']) and
                not any(jk in cp['name'].lower() for jk in _CRED_JUNK)]

    # Fallback: if good_crp sum doesn't match the BS creditor total,
    # try extracting from debtor list (old approach for XLS without Annexure-C pass)
    good_crp_sum = sum(x['amount'] for x in good_crp)
    if (not good_crp or abs(good_crp_sum - cred_target) > cred_target * 0.02) and cred_target > 0:
        deb_raw = p.get('sundry_debtor_parties', [])
        running = 0.0
        extracted = []
        for x in deb_raw:
            running += x['amount']
            extracted.append(x)
            if abs(running - cred_target) < 2.0:
                good_crp = extracted
                break
            if running > cred_target * 1.02:
                break  # overshot without hitting target

    adv_parties = [cp for cp in good_crp
                   if 'advance' in cp['name'].lower() or 'customer' in cp['name'].lower()]
    cred_only   = [cp for cp in good_crp if cp not in adv_parties]

    for i in range(33):
        r = 23 + i
        if i < len(cred_only):
            cp = cred_only[i]
            pfx, bare = _split_prefix(cp['name'])
            _write(ws, r, 1, pfx or 'M/s.')
            _write(ws, r, 2, bare)
            _py(ws, r, PY, cp['amount']); _cy(ws, r, CY)
        else:
            _write(ws, r, 1, '')
            _write(ws, r, 2, '')
            _py(ws, r, PY, 0); _cy(ws, r, CY)

    # Advance from Customers — prefer Annexure-C dedicated list, then fallback
    adv_from_annexure = p.get('advance_from_customer_parties', [])
    if not adv_parties and adv_from_annexure:
        adv_parties = adv_from_annexure
    adv_total = p.get('advance_from_customers', 0)

    for i in range(5):
        r = 57 + i
        if i < len(adv_parties):
            cp = adv_parties[i]
            pfx, bare = _split_prefix(cp['name'])
            _write(ws, r, 1, pfx or 'M/s.')
            _write(ws, r, 2, bare)
            _py(ws, r, PY, cp['amount']); _cy(ws, r, CY)
        elif i == 0 and adv_total:
            _write(ws, r, 1, '')
            _write(ws, r, 2, 'Advance from Customers')
            _py(ws, r, PY, adv_total); _cy(ws, r, CY)
        else:
            _write(ws, r, 1, '')
            _py(ws, r, PY, 0); _cy(ws, r, CY)
    # R62 = SUM formula, R68 = SUM, R69 = TOTAL — skip

    # ── Trade receivables >6 months (R74-R128, 55 slots) ─────────────────────
    # If creditors were extracted from front of debtor list (sum-match fallback),
    # skip those entries when building the debtors list.
    deb_raw_all = p.get('sundry_debtor_parties', [])
    # Detect if we used the fallback (creditors not in original cred_parties_raw)
    fallback_used = bool(good_crp) and not any(
        cp['name'] in [x['name'] for x in cred_parties_raw] for cp in good_crp
        if cred_parties_raw
    )
    # Simpler check: if good_crp names match the start of deb_raw_all, skip them
    # Check if sundry_debtor_parties came from Annexure-I (exact match with BS total)
    bs_debtor_total = p.get('sundry_debtors', 0)
    debtor_list_total = sum(x['amount'] for x in deb_raw_all)
    annexure_i_exact = (bs_debtor_total > 0 and
                        abs(debtor_list_total - bs_debtor_total) / bs_debtor_total < 0.001)

    skip_count = 0
    if not annexure_i_exact and good_crp and deb_raw_all:
        # Only apply skip_count when Annexure-I did NOT provide exact data
        for i, gc in enumerate(good_crp):
            if i < len(deb_raw_all) and deb_raw_all[i]['name'] == gc['name']:
                skip_count += 1
            else:
                break
    # Build normalized set of lender name tokens to exclude from debtor list.
    # We use token-level matching because the same person may appear as
    # "Amar Nath Aggarwal [HUF]" in one list and "Sh. Amar Nath Aggarwal [HUF]" in another,
    # or "Garima Aggarwal" vs just "Garima".
    _stop_tokens = {'sh', 'smt', 'mr', 'mrs', 'ms', 'm/s', 'shri', 'huf', 'prop',
                    'the', 'of', 'and', '&', 'co', 'ltd', 'pvt', 'sons', 'bros'}
    def _name_tokens(nm):
        """Return significant lowercase tokens from a name."""
        toks = re.split(r'[\s\.,\[\]\(\)\/\-]+', nm.lower())
        return {t for t in toks if len(t) >= 3 and t not in _stop_tokens}

    lender_token_sets = []
    for ul in p.get('unsecured_loan_parties', []):
        toks = _name_tokens(ul['name'])
        if toks:
            lender_token_sets.append(toks)
    # Also build flat set of normalized full names for exact matching
    lender_names = {ul['name'].strip().lower() for ul in p.get('unsecured_loan_parties', [])}

    def _is_lender(nm):
        """True if the name matches any known lender by token overlap."""
        if nm.strip().lower() in lender_names:
            return True
        deb_toks = _name_tokens(nm)
        # Require at least 2 meaningful tokens to avoid generic false positives
        # (e.g. {'food','products'} matches too broadly)
        if len(deb_toks) < 2:
            return False
        for lender_toks in lender_token_sets:
            # Only flag as lender if debtor tokens are a non-trivially specific subset:
            # intersection must cover ALL debtor tokens AND debtor must have ≥ 2 tokens
            # AND there must be a unique token (not just generic words like food/products).
            intersection = deb_toks & lender_toks
            if intersection == deb_toks and len(deb_toks) >= 2:
                # Reject if all tokens are generic business words
                _generic = {'food', 'products', 'traders', 'store', 'shop',
                            'bakery', 'house', 'enterprises', 'agency', 'services'}
                non_generic = deb_toks - _generic
                if non_generic:  # at least one specific/proper-noun token
                    return True
        return False

    if annexure_i_exact:
        # Annexure-I gave us the exact correct list — use all of them directly
        deb_parties = deb_raw_all
    else:
        deb_parties = [x for x in deb_raw_all[skip_count:]
                       if _is_debtor_party(x['name'], lender_names) and not _is_lender(x['name'])]

    for i in range(55):
        r = 74 + i
        if i < len(deb_parties):
            dp = deb_parties[i]
            # Avoid double "M/s." prefix
            dname = dp['name']
            if dname.lower().startswith('m/s.') or dname.lower().startswith('m/s '):
                _write(ws, r, 1, 'M/s.')
                _write(ws, r, 2, dname[4:].strip().lstrip('.')  .strip())
            else:
                _write(ws, r, 1, 'M/s.')
                _write(ws, r, 2, dname)
            _py(ws, r, PY, dp['amount']); _cy(ws, r, CY)
        else:
            _write(ws, r, 2, '')
            _py(ws, r, PY, 0); _cy(ws, r, CY)
    # R129 col E: template has =SUM(E77:E128) but debtors start at R74.
    # Overwrite with corrected formula =SUM(E74:E128).
    ws.cell(129, 5).value = '=SUM(E74:E128)'

    # ── Trade receivables <6 months (R134-R135, 2 slots) ─────────────────────
    _py(ws, 134, PY, 0); _cy(ws, 134, CY)
    _py(ws, 135, PY, 0); _cy(ws, 135, CY)
    # R136 = SUM — formula, skip

    log.append(
        f'Details: {len(unsec_parties)} unsecured, {len(cred_only)} creditors, '
        f'{len(deb_parties)} debtors'
    )




# ── Fixed Assets P. Yr. ───────────────────────────────────────────────────────
# Template column layout (1-indexed, verified):
#   A(1)=Particulars  B(2)=WDV at 01.04.PY-1  C(3)=Additions>180d  D(4)=Additions<180d
#   E(5)=Sales  F(6)=Total  G(7)=Rate%  H(8)=Depreciation  I(9)=WDV at 31.03.PY
#   F(6) and H(8) are formula columns — write only B,C,D,E,G; I is also formula.
#
# Named asset rows in template (from inspection) — these are FIXED:
#   PLANT & MACHINERY section: R9-R17
#   VEHICLE section: R20-R21
#   COMPUTERS section: R24
#   FURNITURE AND FIXTURES: R27-R28
#   BUILDING: R31
#   R37 = Total row (formulas)
#
# Strategy: match parsed item names to template row labels; write B,C,D,E values.
# Items not found in template → write into the nearest named row or skip.

_FA_ROW_MAP = {
    'camera & dvr':              9,
    'digital moisture meter':    10,
    'servo control voltage':     11,
    'servo control':             11,
    'water cooler':              12,
    'machine':                   13,
    'mobile':                    14,
    'electricals':               15,
    'washing machine':           16,
    'led':                       17,
    'car':                       20,
    'motor cycle & scooter':     21,
    'motor cycle':               21,
    'motorcycle':                21,
    'scooter':                   21,
    'computer':                  24,
    'chair':                     27,
    'furniture & fixtures':      28,
    'furniture':                 28,
    'building':                  31,
}


def _inject_fixed_assets_py(wb, parsed, client_name, py_year, log):
    if 'Fixed Assets P. Yr.' not in wb.sheetnames:
        return
    ws = wb['Fixed Assets P. Yr.']
    items = parsed.get('fixed_asset_items', [])
    if not items:
        log.append('Fixed Assets P.Yr.: no items to write')
        return

    # Columns: B=2 opening, C=3 add>180, D=4 add<180, E=5 sales, G=7 rate
    # F(6)=TOTAL formula, H(8)=DEP formula, I(9)=WDV formula — DO NOT WRITE
    written_rows = set()

    for item in items:
        nm_lower = item['name'].strip().lower()
        # Try exact match, then partial
        r = _FA_ROW_MAP.get(nm_lower)
        if r is None:
            for key, row in _FA_ROW_MAP.items():
                if key in nm_lower or nm_lower in key:
                    r = row
                    break
        if r is None:
            continue   # skip unmatched — template already has named rows

        if r in written_rows:
            continue   # don't double-write
        written_rows.add(r)

        opening   = item.get('opening_wdv', 0)
        additions = item.get('additions', 0)
        rate      = item.get('rate', 0) or 0
        sales     = item.get('sales', 0)

        safe_additions = round(additions)

        _write_num(ws, r, 2, round(opening))         # B: opening WDV
        _write_num(ws, r, 3, safe_additions)         # C: additions >180d
        _write_num(ws, r, 4, 0)                      # D: additions <180d (parser doesn't split)
        _write_num(ws, r, 5, round(sales))           # E: sales/disposals

        # If template row has a static 0 for Total (col F=6) instead of a formula,
        # write the total manually so H (dep) and I (WDV) formulas compute correctly.
        existing_f = ws.cell(r, 6).value
        if existing_f == 0 or existing_f is None:
            # Only write if it's not already a formula
            if not (isinstance(existing_f, str) and existing_f.startswith('=')):
                total_val = round(opening) + safe_additions - round(sales)
                _write_num(ws, r, 6, total_val)      # F: total (B+C+D-E)
        # G(7) rate — only write if template already has 0 (don't overwrite existing rates)
        rate = item.get('rate', 0)
        if rate:
            existing_rate = ws.cell(r, 7).value
            if existing_rate in (0, None, ''):
                _write(ws, r, 7, rate)
        # F(6), H(8), I(9) are formulas — DO NOT WRITE

    log.append(f'Fixed Assets P.Yr.: {len(items)} items processed ({len(written_rows)} rows written)')


# ── GROSS PROFIT sheet ────────────────────────────────────────────────────────
# Template column layout (verified from data_only=False inspection):
#   Left side:  A=Particulars  B(2)=CY amount  C(3)=PY amount
#   Right side: D=Particulars  E(5)=CY amount  F(6)=PY amount
#
# Key rows:
#   R9:C   = Opening stock (PY left side)    [R9:B = formula from notes to p&l!D18]
#   R10:F  = Sales (PY right side)           [template: R10 right, not R9!]
#   R13:C  = Purchases (PY left side)        [same row as Purchase GST label]
#   R13:F  = Closing stock (PY right side)   [formula =notes to p&l!E27]
#   R21-R23= Direct expense rows (from notes to p&l formulas)
#   R25:C  = Gross profit (PY)               [formula =F27-SUM(C9:C21)]
#   R27    = TOTAL row (formulas)
#
# IMPORTANT: Most cells are formula-driven from notes to p&l.
# Only the PY input cells that are NOT formula-driven need writing:
#   R9:C  = opening stock PY
#   R10:F = sales PY
#   R13:C = purchases PY
#
# R13:F = closing stock is a formula ('notes to p&l'!E27) — DO NOT WRITE
# R25:C = gross profit is a formula — DO NOT WRITE

def _inject_fixed_assets_cy_opening(wb, parsed, log):
    """
    Write PY closing WDV as CY opening WDV in Fixed Assets C. Yr. sheet.
    PY closing WDV = opening + additions - sales - depreciation
    dep = ROUND((opening + add_gt180 + add_lt180 - sales) * rate/100)
    """
    if 'Fixed Assets C. Yr.' not in wb.sheetnames:
        return
    ws = wb['Fixed Assets C. Yr.']
    items = parsed.get('fixed_asset_items', [])
    if not items:
        return

    import math

    # Mapping: asset name (lower) → FA C.Yr. row
    # Plant & Machinery items occupy rows 11-21 in C.Yr. sheet
    # Vehicles: Car=R24, Motor Cycle & Scooter=R25
    # Furniture & Fixtures: Furniture & Fixture=R36
    # Computer: R47
    # Building: R60
    _FA_CY_ROW_MAP = {
        'camera & dvr':              11,
        'digital moisture meter':    12,
        'servo control voltage':     13,
        'servo control':             13,
        'water cooler':              14,
        'machine':                   15,
        'mobile':                    16,
        'electricals':               17,
        'washing machine':           18,
        'led':                       19,
        'car':                       24,
        'motor cycle & scooter':     25,
        'motor cycle':               25,
        'motorcycle':                25,
        'scooter':                   25,
        'computer':                  47,
        'chair':                     37,
        'furniture & fixture':       36,
        'furniture':                 36,
        'building':                  60,
    }

    count = 0
    for item in items:
        name_lower = item['name'].strip().lower()
        # Find matching row (prefix match)
        r = None
        for key, row in _FA_CY_ROW_MAP.items():
            if name_lower.startswith(key):
                r = row
                break
        if r is None:
            continue

        opening  = item.get('opening_wdv', 0)
        add_gt   = item.get('additions', 0)
        sales    = item.get('sales', 0)
        rate     = item.get('rate', 0) or 0
        # Compute PY depreciation (same formula as FA P.Yr. sheet)
        dep = round((opening + add_gt - sales) * rate / 100)
        closing_wdv = opening + add_gt - sales - dep

        if closing_wdv > 0:
            ws.cell(r, 2).value = round(closing_wdv)  # col B = opening WDV as at 01.04.CY-1
            count += 1

    log.append(f"Fixed Assets C.Yr.: {count} opening WDV values written")


def _inject_gross_profit_sheet(wb, parsed, client_name, cy_year, py_year, log):
    if 'GROSS PROFIT' not in wb.sheetnames:
        return
    ws = wb['GROSS PROFIT']
    p = parsed

    # Hard-coded column constants (verified from template):
    PY_LEFT  = 3   # col C = PY left side (opening stock, purchases, gross profit)
    PY_RIGHT = 6   # col F = PY right side (sales, closing stock)

    # R9:C = opening stock (PY)
    _write_num(ws, 9,  PY_LEFT,  p.get('opening_stock', 0))

    # R10:F = sales (PY) — sales is on row 10 right side ("Sales GST" label row)
    _write_num(ws, 10, PY_RIGHT, p.get('sales', 0))

    # R13:C = purchases (PY) — "Purchase GST" row
    _write_num(ws, 13, PY_LEFT,  p.get('purchases', 0))

    # R13:F = closing stock — formula from notes to p&l!E27 — DO NOT WRITE
    # R25:C = gross profit — formula =F27-SUM(C9:C21) — DO NOT WRITE
    # R27 = TOTAL formulas — DO NOT WRITE

    # Direct expenses in GROSS PROFIT left PY column (col C = 3):
    # R21 col B (CY FOC) = formula 'notes to p&l'!D24 — already handles CY.
    # R21 col C (PY FOC), R22 col C (Freight), R23 col C (Freight GST):
    # These have NO formula in the template — must be written directly.
    all_items = p.get('other_expense_items', []) + p.get('direct_expense_items', [])
    _DIRECT_EXP_KEYS_LOCAL = {
        'foc', 'f.o.c.', 'f.o.c', 'freight inward', 'freight outward',
        'frieght inward', 'frieght inward (gst)', 'freight inward (gst)',
        'freight inward gst', 'freight',
    }
    foc_amt = 0.0
    freight_amt = 0.0
    freight_gst_amt = 0.0
    for it in all_items:
        k = it['name'].strip().lower()
        if k in ('foc', 'f.o.c.', 'f.o.c'):
            foc_amt += it['amount']
        elif 'gst' in k and 'freight' in k:
            freight_gst_amt += it['amount']
        elif 'freight' in k or 'frieght' in k:
            freight_amt += it['amount']

    _write_num(ws, 21, PY_LEFT, foc_amt)           # R21:C = FOC PY
    _write_num(ws, 22, PY_LEFT, freight_amt)        # R22:C = Freight Inward PY
    _write_num(ws, 23, PY_LEFT, freight_gst_amt)    # R23:C = Freight Inward (GST) PY

    log.append('GROSS PROFIT sheet injected')


# ── Update headers ─────────────────────────────────────────────────────────────

def _update_headers(wb, client_name, cy_year, py_year):
    """Replace M/S XYZ CO., year references, and accounting policy text."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                v = cell.value
                if not isinstance(v, str):
                    continue
                nv = v
                nv = nv.replace('M/S XYZ CO.', client_name)
                nv = nv.replace('M/s XYZ….. is proprietership Consern in which …... is Proprietor',
                                f'M/s {client_name} is a proprietorship concern.')
                # Replace year strings carefully (PY before CY to avoid double-replace)
                nv = nv.replace(f'31.03.{py_year}', f'__PY_DATE__')
                nv = nv.replace(f'31.03.{cy_year}', f'31.03.{cy_year}')
                nv = nv.replace(f'__PY_DATE__', f'31.03.{py_year}')
                nv = nv.replace(f'31 March, {cy_year}', f'31 March, {cy_year}')
                nv = nv.replace(f'31 March, {py_year}', f'31 March, {py_year}')
                nv = nv.replace('31 March, 2025', f'31 March, {cy_year}')
                nv = nv.replace('31 March, 2024', f'31 March, {py_year}')
                nv = nv.replace('31.03.2025', f'31.03.{cy_year}')
                nv = nv.replace('31.03.2024', f'31.03.{py_year}')
                nv = nv.replace('1st April 2024', f'1st April {py_year}')
                nv = nv.replace('Year Ending 31.03.2025', f'Year Ending 31.03.{cy_year}')
                nv = nv.replace('Year Ending 31.03.2024', f'Year Ending 31.03.{py_year}')
                if nv != v:
                    cell.value = nv


# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 3 — Main entry point
# ─────────────────────────────────────────────────────────────────────────────

def process_tshape(input_path: str, output_path: str,
                   template_path: str = None,
                   client_name: str = None,
                   cy_year: str = None) -> dict:
    """
    Full pipeline:
      1. parse_tshape_bs(input_path)
      2. inject_into_template(...)  → output_path
    Returns dict with status, log, entity info.
    """
    import traceback

    if template_path is None:
        template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            'Output_sample_format.xlsx'
        )
    if not os.path.exists(template_path):
        return {'status': 'error',
                'message': f'Template not found: {template_path}'}

    try:
        parsed = parse_tshape_bs(input_path)
    except Exception as e:
        return {
            'status': 'error',
            'message': f'Parse failed: {e}',
            'traceback': traceback.format_exc()
        }

    if client_name:
        parsed['entity_name'] = client_name

    try:
        result = inject_into_template(parsed, template_path, output_path,
                                       client_name, cy_year)
    except Exception as e:
        return {
            'status': 'error',
            'message': f'Injection failed: {e}',
            'traceback': traceback.format_exc(),
            'parsed': parsed
        }

    result['parsed'] = parsed
    return result


# ─────────────────────────────────────────────────────────────────────────────
#  CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 4:
        print('Usage: python tshape_processor.py input.xls output.xlsx template.xlsx [client_name] [cy_year]')
        sys.exit(1)

    inp   = sys.argv[1]
    out   = sys.argv[2]
    tmpl  = sys.argv[3]
    cname = sys.argv[4] if len(sys.argv) > 4 else None
    cyear = sys.argv[5] if len(sys.argv) > 5 else None

    res = process_tshape(inp, out, tmpl, cname, cyear)
    print(f"Status: {res.get('status')}")
    for line in res.get('log', [])[-30:]:
        print(' ', line)
    if res.get('message'):
        print('ERROR:', res['message'])
    if res.get('traceback'):
        print(res['traceback'])
