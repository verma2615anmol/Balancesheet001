"""
Balance Sheet Year-Shift Processor  (fixed v3 — auto-detect)
Shifts CY→PY, clears CY constants, updates all date references.
Fully preserves formatting using openpyxl.

Key changes vs v2:
  1. AUTO-DETECTS CY and PY columns by scanning the first 15 rows of every
     sheet for date headers containing the closing year ("31.03.2025",
     "31st March, 2025", etc.) and the previous year.  Falls back to the
     hardcoded SHEET_COL_MAP only as a secondary lookup (case-insensitive).
  2. Case-insensitive sheet-name matching for the hardcoded map.
  3. Every sheet with detected CY/PY columns gets the full treatment:
     copy CY→PY, clear CY constants, update dates, fix PY header.
  4. Sheets with NO detected columns still get date-text updates.
"""

from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter


# ─── Sheet → (CY_col, PY_col) FALLBACK mapping (case-insensitive lookup) ────
SHEET_COL_MAP = {
    "bs":           [("E", "F")],
    "p&l":          [("E", "F")],
    "notes to bs":  [("D", "E")],
    "notes to p&l": [("D", "E")],
    "details":      [("D", "E")],
    "gross profit": [("B", "C"), ("F", "G")],
}

TEXT_ONLY_SHEETS_LOWER = {s.strip().lower() for s in [
    "notes to accounts", "Fixed Assets C. Yr.", "Fixed Assets P. Yr.",
    "FA2022", "Tax audit ", "Tax Audit report", "PPE",
]}

CAPITAL_CY_ROW   = 8
CAPITAL_PY_ROW   = 11
CAPITAL_DATA_COLS = ["C", "D", "E"]

# How many header rows to scan for date strings
HEADER_SCAN_ROWS = 15


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _is_formula(val):
    return isinstance(val, str) and val.strip().startswith("=")


def _is_numeric(val):
    """Return True only for actual numeric values (int/float)."""
    return isinstance(val, (int, float))


def _date_replacements(cy: str, ny: str) -> list:
    """
    Build (old, new) replacement pairs.
    Uses a two-pass placeholder approach so 2024→2025→2026 chaining can't happen.
    cy = closing year string (e.g. "2025")
    ny = new year string      (e.g. "2026")
    """
    prev_open = str(int(cy) - 1)   # e.g. "2024" (opening date of CY)
    PH = "__NEWCY__"

    step_a = [
        (f"31.03.{cy}",                           f"31.03.{PH}"),
        (f"31 March, {cy}",                        f"31 March, {PH}"),
        (f"31 March {cy}",                         f"31 March {PH}"),
        (f"31st March, {cy}",                      f"31st March, {PH}"),
        (f"31st March {cy}",                       f"31st March {PH}"),
        (f"31ST MARCH ,{cy}",                      f"31ST MARCH ,{PH}"),
        (f"31ST MARCH, {cy}",                      f"31ST MARCH, {PH}"),
        (f"31ST MARCH {cy}",                       f"31ST MARCH {PH}"),
        (f"31 MARCH, {cy}",                        f"31 MARCH, {PH}"),
        (f"31 MARCH {cy}",                         f"31 MARCH {PH}"),
        (f"year ended 31 March, {cy}",             f"year ended 31 March, {PH}"),
        (f"year ended, 31st March, {cy}",          f"year ended, 31st March, {PH}"),
        (f"year end 31 March, {cy}",               f"year end 31 March, {PH}"),
        (f"year ending 31.03.{cy}",                f"year ending 31.03.{PH}"),
        (f"YEAR ENDING 31.03.{cy}",                f"YEAR ENDING 31.03.{PH}"),
        (f"YEAR ENDING 31ST MARCH ,{cy}",          f"YEAR ENDING 31ST MARCH ,{PH}"),
        (f"YEAR ENDING 31ST MARCH, {cy}",          f"YEAR ENDING 31ST MARCH, {PH}"),
        (f"YEAR ENDING 31ST MARCH {cy}",           f"YEAR ENDING 31ST MARCH {PH}"),
        (f"as at 31 March, {cy}",                  f"as at 31 March, {PH}"),
        (f"AS AT 31ST MARCH {cy}",                 f"AS AT 31ST MARCH {PH}"),
        (f"AS AT 31ST MARCH, {cy}",                f"AS AT 31ST MARCH, {PH}"),
        (f"AS AT 31 MARCH, {cy}",                  f"AS AT 31 MARCH, {PH}"),
        (f"for the year ended, 31st March, {cy}",  f"for the year ended, 31st March, {PH}"),
        (f"for the year ended 31 March, {cy}",     f"for the year ended 31 March, {PH}"),
        (f"FOR THE YEAR ENDED 31ST MARCH, {cy}",   f"FOR THE YEAR ENDED 31ST MARCH, {PH}"),
        (f"FOR THE YEAR ENDED 31ST MARCH {cy}",    f"FOR THE YEAR ENDED 31ST MARCH {PH}"),
    ]

    step_b = [
        (f"1st April {prev_open}",  f"1st April {cy}"),
        (f"1 April {prev_open}",    f"1 April {cy}"),
        (f"01.04.{prev_open}",      f"01.04.{cy}"),
        (f"31.03.{prev_open}",      f"31.03.{cy}"),
        (f"31st March {prev_open}", f"31st March {cy}"),
        (f"31st March, {prev_open}",f"31st March, {cy}"),
        (f"31 March, {prev_open}",  f"31 March, {cy}"),
        (f"31 March {prev_open}",   f"31 March {cy}"),
        (f"31ST MARCH {prev_open}", f"31ST MARCH {cy}"),
        (f"31ST MARCH, {prev_open}",f"31ST MARCH, {cy}"),
        (f"AS AT 31ST MARCH {prev_open}",  f"AS AT 31ST MARCH {cy}"),
        (f"AS AT 31ST MARCH, {prev_open}", f"AS AT 31ST MARCH, {cy}"),
    ]

    step_c = [
        (f"31.03.{PH}",                           f"31.03.{ny}"),
        (f"31 March, {PH}",                        f"31 March, {ny}"),
        (f"31 March {PH}",                         f"31 March {ny}"),
        (f"31st March, {PH}",                      f"31st March, {ny}"),
        (f"31st March {PH}",                       f"31st March {ny}"),
        (f"31ST MARCH ,{PH}",                      f"31ST MARCH ,{ny}"),
        (f"31ST MARCH, {PH}",                      f"31ST MARCH, {ny}"),
        (f"31ST MARCH {PH}",                       f"31ST MARCH {ny}"),
        (f"31 MARCH, {PH}",                        f"31 MARCH, {ny}"),
        (f"31 MARCH {PH}",                         f"31 MARCH {ny}"),
        (f"year ended 31 March, {PH}",             f"year ended 31 March, {ny}"),
        (f"year ended, 31st March, {PH}",          f"year ended, 31st March, {ny}"),
        (f"year end 31 March, {PH}",               f"year end 31 March, {ny}"),
        (f"year ending 31.03.{PH}",                f"year ending 31.03.{ny}"),
        (f"YEAR ENDING 31.03.{PH}",                f"YEAR ENDING 31.03.{ny}"),
        (f"YEAR ENDING 31ST MARCH ,{PH}",          f"YEAR ENDING 31ST MARCH ,{ny}"),
        (f"YEAR ENDING 31ST MARCH, {PH}",          f"YEAR ENDING 31ST MARCH, {ny}"),
        (f"YEAR ENDING 31ST MARCH {PH}",           f"YEAR ENDING 31ST MARCH {ny}"),
        (f"as at 31 March, {PH}",                  f"as at 31 March, {ny}"),
        (f"AS AT 31ST MARCH {PH}",                 f"AS AT 31ST MARCH {ny}"),
        (f"AS AT 31ST MARCH, {PH}",                f"AS AT 31ST MARCH, {ny}"),
        (f"AS AT 31 MARCH, {PH}",                  f"AS AT 31 MARCH, {ny}"),
        (f"for the year ended, 31st March, {PH}",  f"for the year ended, 31st March, {ny}"),
        (f"for the year ended 31 March, {PH}",     f"for the year ended 31 March, {ny}"),
        (f"FOR THE YEAR ENDED 31ST MARCH, {PH}",   f"FOR THE YEAR ENDED 31ST MARCH, {ny}"),
        (f"FOR THE YEAR ENDED 31ST MARCH {PH}",    f"FOR THE YEAR ENDED 31ST MARCH {ny}"),
    ]

    # Fiscal year range patterns: "2024-25", "2024-2025", "2023-24", etc.
    prev_open_short = prev_open[2:]   # "24"
    cy_short        = cy[2:]          # "25"
    ny_short        = ny[2:]          # "26"
    PH_RANGE        = "__FYRNG__"
    prev_prev       = str(int(prev_open) - 1)
    prev_prev_short = prev_prev[2:]

    range_steps = [
        # Step A: CY fiscal range → placeholder   e.g. "2024-25" → placeholder
        (f"{prev_open}-{cy_short}",  PH_RANGE),
        (f"{prev_open}-{cy}",        f"{PH_RANGE}L"),
        # Step B: PY fiscal range → new PY range  e.g. "2023-24" → "2024-25"
        (f"{prev_prev}-{prev_open_short}", f"{prev_open}-{cy_short}"),
        (f"{prev_prev}-{prev_open}",       f"{prev_open}-{cy}"),
        # Step C: placeholder → new CY range      e.g. placeholder → "2025-26"
        (f"{PH_RANGE}L", f"{cy}-{ny}"),
        (PH_RANGE,       f"{cy}-{ny_short}"),
    ]

    return step_a + step_b + step_c + range_steps


def _replace_text(val, pairs):
    if not isinstance(val, str):
        return val
    for old, new in pairs:
        val = val.replace(old, new)
    return val


# ─── Auto-detection of CY/PY columns ────────────────────────────────────────

def _cell_contains_year_date(val, year_str):
    """Check if a cell's string value contains a date reference for the given year."""
    if not isinstance(val, str):
        return False
    flat = val.replace("\n", " ").replace("\r", " ")
    if year_str not in flat:
        return False
    # Must appear in a date/financial context — not just any random number
    markers = ["31.03.", "31 03", "March", "MARCH", "march",
               "year end", "Year end", "YEAR END",
               "as at", "As at", "AS AT"]
    # Also accept bare "31.03.YYYY" which is a standalone date header
    if f"31.03.{year_str}" in flat:
        return True
    return any(m in flat for m in markers)


def _detect_cy_py_columns(ws, closing_year):
    """
    Scan the first HEADER_SCAN_ROWS rows of a worksheet for date headers.
    Returns a list of (cy_col_letter, py_col_letter) tuples found.

    A column header containing the closing_year date = CY column.
    A column header containing (closing_year - 1) date = PY column.
    Pairs them by proximity: CY immediately left of PY (most common layout).
    """
    cy_str = str(closing_year)
    py_str = str(closing_year - 1)

    cy_cols = set()
    py_cols = set()

    max_col = min(ws.max_column or 1, 50)

    for row_idx in range(1, HEADER_SCAN_ROWS + 1):
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(cell, MergedCell):
                continue
            val = cell.value
            if not isinstance(val, str):
                continue
            if _is_formula(val):
                continue
            if _cell_contains_year_date(val, cy_str):
                cy_cols.add(col_idx)
            if _cell_contains_year_date(val, py_str):
                py_cols.add(col_idx)

    if not cy_cols:
        return []

    # Pair CY with nearest PY column to its right
    pairs = []
    used_py = set()

    for cy_c in sorted(cy_cols):
        best_py = None
        for offset in [1, 2, 3]:
            candidate = cy_c + offset
            if candidate in py_cols and candidate not in used_py:
                best_py = candidate
                break
        if best_py:
            pairs.append((get_column_letter(cy_c), get_column_letter(best_py)))
            used_py.add(best_py)

    return pairs


# ─── Core operations ─────────────────────────────────────────────────────────

def _copy_cy_to_py(ws_formula, ws_data, cy_letter, py_letter):
    """
    Copy calculated CY values → PY column.
    Only copies NUMERIC values — never strings.
    """
    cy_col = column_index_from_string(cy_letter)
    py_col = column_index_from_string(py_letter)

    for row_idx in range(1, ws_formula.max_row + 1):
        cy_data_cell = ws_data.cell(row=row_idx, column=cy_col)
        py_cell      = ws_formula.cell(row=row_idx, column=py_col)

        if isinstance(py_cell, MergedCell):
            continue

        cv = cy_data_cell.value

        if cv is None:
            continue
        if not _is_numeric(cv):
            continue
        if py_cell.value is None:
            continue

        py_cell.value = cv


def _clear_cy_constants(ws, cy_letter):
    """Delete hardcoded numeric constants in CY column; leave formulas and text."""
    cy_col = column_index_from_string(cy_letter)
    for row_idx in range(1, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=cy_col)
        if isinstance(cell, MergedCell):
            continue
        val = cell.value
        if val is None:
            continue
        if _is_formula(val):
            continue
        if isinstance(val, str):
            continue
        cell.value = None


def _update_text_in_sheet(ws, pairs):
    """Replace date text in all non-formula string cells."""
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and not _is_formula(cell.value):
                new_val = _replace_text(cell.value, pairs)
                if new_val != cell.value:
                    cell.value = new_val


def _fix_py_header_in_column(ws, py_letter, closing_year, new_year):
    """
    After date replacement, the PY column header might incorrectly show new_year.
    Scan the first rows and revert it to closing_year.
    """
    py_col = column_index_from_string(py_letter)
    ny_str = str(new_year)
    cy_str = str(closing_year)

    for row_idx in range(1, HEADER_SCAN_ROWS + 1):
        cell = ws.cell(row=row_idx, column=py_col)
        if isinstance(cell, MergedCell):
            continue
        val = cell.value
        if not isinstance(val, str):
            continue
        if _is_formula(val):
            continue
        if ny_str in val and any(marker in val for marker in
                                  ("31.03.", "March", "MARCH", "march",
                                   "year ending", "Year ending", "YEAR ENDING",
                                   "as at", "As at", "AS AT")):
            cell.value = val.replace(ny_str, cy_str)


# ─── Main entry point ────────────────────────────────────────────────────────

def process(input_path: str, output_path: str, closing_year: int, new_year: int) -> dict:
    """
    closing_year : year whose data is currently in the CY column (e.g. 2025)
    new_year     : year we are preparing for (e.g. 2026)
    Returns dict with processing log.
    """
    pairs = _date_replacements(str(closing_year), str(new_year))

    wb      = load_workbook(input_path)
    wb_vals = load_workbook(input_path, data_only=True)

    log = []
    processed_sheets = set()

    # Build case-insensitive lookup from the hardcoded map
    fallback_map = {}
    for key, col_pairs in SHEET_COL_MAP.items():
        fallback_map[key.strip().lower()] = col_pairs

    # ── Process every sheet ───────────────────────────────────────────────────
    for sheet_name in wb.sheetnames:
        ws      = wb[sheet_name]
        ws_data = wb_vals[sheet_name]
        sn_lower = sheet_name.strip().lower()

        # Skip text-only sheets (only date update, no CY/PY shift)
        if sn_lower in TEXT_ONLY_SHEETS_LOWER:
            _update_text_in_sheet(ws, pairs)
            log.append(f"✓ {sheet_name}: dates updated (text-only)")
            processed_sheets.add(sheet_name)
            continue

        # Try auto-detection first
        detected_pairs = _detect_cy_py_columns(ws, closing_year)

        # Fallback to hardcoded map if auto-detection found nothing
        if not detected_pairs and sn_lower in fallback_map:
            detected_pairs = fallback_map[sn_lower]

        if detected_pairs:
            for cy_letter, py_letter in detected_pairs:
                _copy_cy_to_py(ws, ws_data, cy_letter, py_letter)
                _clear_cy_constants(ws, cy_letter)

            _update_text_in_sheet(ws, pairs)

            for cy_letter, py_letter in detected_pairs:
                _fix_py_header_in_column(ws, py_letter, closing_year, new_year)

            cols_desc = ", ".join(f"{cy}→{py}" for cy, py in detected_pairs)
            log.append(f"✓ {sheet_name}: CY→PY copied ({cols_desc}), CY cleared, dates updated")
            processed_sheets.add(sheet_name)

    # ── Capital sheet (row-based, not column-based) ───────────────────────────
    capital_sheet = None
    for sn in wb.sheetnames:
        if sn.strip().lower() == "capital":
            capital_sheet = sn
            break

    if capital_sheet and capital_sheet not in processed_sheets:
        ws      = wb[capital_sheet]
        ws_data = wb_vals[capital_sheet]

        for col_letter in CAPITAL_DATA_COLS:
            col          = column_index_from_string(col_letter)
            cy_cell_data = ws_data.cell(row=CAPITAL_CY_ROW, column=col)
            py_cell      = ws.cell(row=CAPITAL_PY_ROW, column=col)
            if isinstance(py_cell, MergedCell):
                continue
            if cy_cell_data.value is not None and _is_numeric(cy_cell_data.value):
                py_cell.value = cy_cell_data.value

        for col_letter in CAPITAL_DATA_COLS:
            col  = column_index_from_string(col_letter)
            cell = ws.cell(row=CAPITAL_CY_ROW, column=col)
            if isinstance(cell, MergedCell):
                continue
            if not _is_formula(cell.value):
                cell.value = None

        _update_text_in_sheet(ws, pairs)
        log.append(f"✓ {capital_sheet}: CY row→PY row copied, CY cleared, dates updated")
        processed_sheets.add(capital_sheet)

    # ── Remaining sheets: date-text update only ───────────────────────────────
    for sheet_name in wb.sheetnames:
        if sheet_name not in processed_sheets:
            _update_text_in_sheet(wb[sheet_name], pairs)
            log.append(f"✓ {sheet_name}: dates updated")

    wb.save(output_path)
    return {"status": "success", "log": log, "output": output_path}


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 5:
        print("Usage: python processor.py input.xlsx output.xlsx closing_year new_year")
        sys.exit(1)
    result = process(sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]))
    for line in result["log"]:
        print(line)
    print(f"\nSaved → {result['output']}")
