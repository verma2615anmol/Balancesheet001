"""
Balance Sheet Year-Shift Processor  (fixed v2)
Shifts CY→PY, clears CY constants, updates all date references.
Fully preserves formatting using openpyxl.

Key fixes vs v1:
  1. _copy_cy_to_py now SKIPS string values — text headers/labels must never
     be copied from the CY column to the PY column, because that overwrites
     the PY date header (e.g. "As at 31.03.2024") with the CY header text
     ("As at 31.03.2025"), causing both columns to show the new year.
  2. PY header correction is now done by AUTO-DETECTION (scan for the PY year
     string in the PY column header rows) instead of hardcoded cell addresses.
     This makes the tool work with any workbook regardless of which row the
     headers sit on.
"""

from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.utils import column_index_from_string


# ─── Sheet → (CY_col, PY_col) mapping ───────────────────────────────────────
SHEET_COL_MAP = {
    "bs":           [("E", "F")],
    "p&l":          [("E", "F")],
    "notes to bs":  [("D", "E")],
    "notes to p&l": [("D", "E")],
    "Details":      [("D", "E")],
    "GROSS PROFIT": [("B", "C"), ("F", "G")],
}

TEXT_ONLY_SHEETS = [
    "notes to accounts", "Fixed Assets C. Yr.", "Fixed Assets P. Yr.",
    "FA2022", "Tax audit ", "Tax Audit report", "PPE",
]

CAPITAL_CY_ROW   = 8
CAPITAL_PY_ROW   = 11
CAPITAL_DATA_COLS = ["C", "D", "E"]


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

    # Pass A: mark all CY year occurrences with a placeholder
    # Pass B: update old PY dates (prev_open / cy-1) → cy  (they become new PY labels)
    # Pass C: replace placeholder → ny

    step_a = [
        (f"31.03.{cy}",                           f"31.03.{PH}"),
        (f"31 March, {cy}",                        f"31 March, {PH}"),
        (f"31 March {cy}",                         f"31 March {PH}"),
        (f"31st March, {cy}",                      f"31st March, {PH}"),
        (f"31st March {cy}",                       f"31st March {PH}"),
        (f"31ST MARCH ,{cy}",                      f"31ST MARCH ,{PH}"),
        (f"31ST MARCH, {cy}",                      f"31ST MARCH, {PH}"),
        (f"31 MARCH, {cy}",                        f"31 MARCH, {PH}"),
        (f"year ended 31 March, {cy}",             f"year ended 31 March, {PH}"),
        (f"year ended, 31st March, {cy}",          f"year ended, 31st March, {PH}"),
        (f"year end 31 March, {cy}",               f"year end 31 March, {PH}"),
        (f"year ending 31.03.{cy}",                f"year ending 31.03.{PH}"),
        (f"YEAR ENDING 31ST MARCH ,{cy}",          f"YEAR ENDING 31ST MARCH ,{PH}"),
        (f"as at 31 March, {cy}",                  f"as at 31 March, {PH}"),
        (f"for the year ended, 31st March, {cy}",  f"for the year ended, 31st March, {PH}"),
        (f"for the year ended 31 March, {cy}",     f"for the year ended 31 March, {PH}"),
    ]

    step_b = [
        # Old opening dates (April of prev_open year) → April of cy
        (f"1st April {prev_open}",  f"1st April {cy}"),
        (f"1 April {prev_open}",    f"1 April {cy}"),
        (f"01.04.{prev_open}",      f"01.04.{cy}"),
        # Old PY closing dates → cy  (these become new PY labels after copy)
        (f"31.03.{prev_open}",      f"31.03.{cy}"),
        (f"31st March {prev_open}", f"31st March {cy}"),
        (f"31st March, {prev_open}",f"31st March, {cy}"),
        (f"31 March, {prev_open}",  f"31 March, {cy}"),
    ]

    step_c = [
        (f"31.03.{PH}",                           f"31.03.{ny}"),
        (f"31 March, {PH}",                        f"31 March, {ny}"),
        (f"31 March {PH}",                         f"31 March {ny}"),
        (f"31st March, {PH}",                      f"31st March, {ny}"),
        (f"31st March {PH}",                       f"31st March {ny}"),
        (f"31ST MARCH ,{PH}",                      f"31ST MARCH ,{ny}"),
        (f"31ST MARCH, {PH}",                      f"31ST MARCH, {ny}"),
        (f"31 MARCH, {PH}",                        f"31 MARCH, {ny}"),
        (f"year ended 31 March, {PH}",             f"year ended 31 March, {ny}"),
        (f"year ended, 31st March, {PH}",          f"year ended, 31st March, {ny}"),
        (f"year end 31 March, {PH}",               f"year end 31 March, {ny}"),
        (f"year ending 31.03.{PH}",                f"year ending 31.03.{ny}"),
        (f"YEAR ENDING 31ST MARCH ,{PH}",          f"YEAR ENDING 31ST MARCH ,{ny}"),
        (f"as at 31 March, {PH}",                  f"as at 31 March, {ny}"),
        (f"for the year ended, 31st March, {PH}",  f"for the year ended, 31st March, {ny}"),
        (f"for the year ended 31 March, {PH}",     f"for the year ended 31 March, {ny}"),
    ]

    return step_a + step_b + step_c


def _replace_text(val, pairs):
    if not isinstance(val, str):
        return val
    for old, new in pairs:
        val = val.replace(old, new)
    return val


def _copy_cy_to_py(ws_formula, ws_data, cy_letter, py_letter):
    """
    Copy calculated CY values → PY column.

    CRITICAL FIX: Only copy NUMERIC values. String values (text labels,
    date headers like "As at 31.03.2025") must NEVER be copied to the PY
    column — doing so overwrites the PY date header with the CY date text,
    causing both columns to display the new year.
    """
    cy_col = column_index_from_string(cy_letter)
    py_col = column_index_from_string(py_letter)

    for row_idx in range(1, ws_formula.max_row + 1):
        cy_data_cell = ws_data.cell(row=row_idx, column=cy_col)
        py_cell      = ws_formula.cell(row=row_idx, column=py_col)

        if isinstance(py_cell, MergedCell):
            continue

        cv = cy_data_cell.value

        # ── FIX: skip None and all string values ──────────────────────────────
        # Only numeric values represent financial figures that need to be
        # carried forward as the new PY comparison column.
        # Strings are either: date headers, row labels, or dash placeholders —
        # none of which should overwrite the PY column's own text.
        if cv is None:
            continue
        if not _is_numeric(cv):
            continue

        # Only copy if the PY cell already had something (don't fill blanks)
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
            continue   # text labels are part of structure — keep them
        # It's a numeric constant → clear it
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
    After date replacement, the PY column's header cell will now say
    "As at 31.03.{new_year}" (because _date_replacements updated it).
    Find it by scanning the first 15 rows of the PY column and restore it
    to "As at 31.03.{closing_year}" (the year that just became PY).

    Auto-detects the header row — no hardcoded row numbers needed.
    """
    py_col   = column_index_from_string(py_letter)
    ny_str   = str(new_year)
    cy_str   = str(closing_year)

    for row_idx in range(1, 16):
        cell = ws.cell(row=row_idx, column=py_col)
        if isinstance(cell, MergedCell):
            continue
        val = cell.value
        if not isinstance(val, str):
            continue
        if _is_formula(val):
            continue
        # If this header now incorrectly shows the new year, fix it back to closing year
        if ny_str in val and any(marker in val for marker in
                                  ("31.03.", "March", "MARCH", "year ending", "Year ending")):
            cell.value = val.replace(ny_str, cy_str)


def _fix_gross_profit_py_headers(ws, closing_year, new_year):
    """GROSS PROFIT has two PY header cells: C10 and G10."""
    _fix_py_header_in_column(ws, "C", closing_year, new_year)
    _fix_py_header_in_column(ws, "G", closing_year, new_year)


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

    # ── 1. Sheets with CY/PY column mapping ──────────────────────────────────
    for sheet_name, col_pairs in SHEET_COL_MAP.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws      = wb[sheet_name]
        ws_data = wb_vals[sheet_name]

        for cy_letter, py_letter in col_pairs:
            _copy_cy_to_py(ws, ws_data, cy_letter, py_letter)
            _clear_cy_constants(ws, cy_letter)

        # Update all date text (this will also update PY header — fixed next)
        _update_text_in_sheet(ws, pairs)

        # Fix PY column header(s) that got incorrectly updated to new_year
        if sheet_name == "GROSS PROFIT":
            _fix_gross_profit_py_headers(ws, closing_year, new_year)
        else:
            for cy_letter, py_letter in col_pairs:
                _fix_py_header_in_column(ws, py_letter, closing_year, new_year)

        log.append(f"✓ {sheet_name}: CY→PY copied, CY constants cleared, dates updated")

    # ── 2. Capital sheet (row-based, not column-based) ────────────────────────
    if "capital" in wb.sheetnames:
        ws      = wb["capital"]
        ws_data = wb_vals["capital"]

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
        log.append("✓ capital: CY row→PY row copied, CY constants cleared, dates updated")

    # ── 3. Text-only sheets ───────────────────────────────────────────────────
    for sheet_name in TEXT_ONLY_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        _update_text_in_sheet(wb[sheet_name], pairs)
        log.append(f"✓ {sheet_name}: dates updated")

    # ── 4. Any remaining sheets ───────────────────────────────────────────────
    handled = (set(SHEET_COL_MAP.keys()) | set(TEXT_ONLY_SHEETS) | {"capital"})
    for sheet_name in wb.sheetnames:
        if sheet_name not in handled:
            _update_text_in_sheet(wb[sheet_name], pairs)
            log.append(f"✓ {sheet_name}: dates updated (catch-all)")

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
