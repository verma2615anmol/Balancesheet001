"""
Balance Sheet Year-Shift Processor — Universal Edition
Supports .xlsx and .xls files. Handles datetime cells safely.
"""

import re
import sys
import os
import subprocess
import tempfile
import datetime
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.utils import get_column_letter, column_index_from_string


def _is_formula(val):
    return isinstance(val, str) and val.strip().startswith("=")


def _replace_text(val, pairs):
    if not isinstance(val, str):
        return val
    for old, new in pairs:
        val = val.replace(old, new)
    return val


def _date_pairs(closing_year: int, new_year: int) -> list:
    cy, ny  = str(closing_year), str(new_year)
    py_open = str(closing_year - 1)
    ny_open = str(closing_year)
    return [
        (f"31.03.{cy}",                           f"31.03.{ny}"),
        (f"31 March, {cy}",                        f"31 March, {ny}"),
        (f"31 March {cy}",                         f"31 March {ny}"),
        (f"31st March, {cy}",                      f"31st March, {ny}"),
        (f"31st March {cy}",                       f"31st March {ny}"),
        (f"31ST MARCH ,{cy}",                      f"31ST MARCH ,{ny}"),
        (f"31ST MARCH, {cy}",                      f"31ST MARCH, {ny}"),
        (f"31 MARCH, {cy}",                        f"31 MARCH, {ny}"),
        (f"year ended 31 March, {cy}",             f"year ended 31 March, {ny}"),
        (f"year ended, 31st March, {cy}",          f"year ended, 31st March, {ny}"),
        (f"year ending 31.03.{cy}",                f"year ending 31.03.{ny}"),
        (f"YEAR ENDING 31ST MARCH ,{cy}",          f"YEAR ENDING 31ST MARCH ,{ny}"),
        (f"YEAR ENDING 31ST MARCH, {cy}",          f"YEAR ENDING 31ST MARCH, {ny}"),
        (f"as at 31 March, {cy}",                  f"as at 31 March, {ny}"),
        (f"As at 31 March, {cy}",                  f"As at 31 March, {ny}"),
        (f"for the year ended, 31st March, {cy}",  f"for the year ended, 31st March, {ny}"),
        (f"for the year ended 31 March, {cy}",     f"for the year ended 31 March, {ny}"),
        (f"For the year ended 31 March, {cy}",     f"For the year ended 31 March, {ny}"),
        (f"1st April {py_open}",                   f"1st April {ny_open}"),
        (f"1 April {py_open}",                     f"1 April {ny_open}"),
        (f"01.04.{py_open}",                       f"01.04.{ny_open}"),
        (f"1ST APRIL {py_open}",                   f"1ST APRIL {ny_open}"),
        (f"1ST APRIL, {py_open}",                  f"1ST APRIL, {ny_open}"),
    ]


DATE_KEYWORDS = [
    '31.03.', '31 march', '31st march', '31 MARCH', '31ST MARCH',
    'year ending', 'year ended', 'as at',
]


def _col_has_values(ws_vals, col_idx, min_count=3):
    if col_idx <= 1:
        return False
    count = 0
    for row in ws_vals.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.column == col_idx and isinstance(cell.value, (int, float)):
                count += 1
                if count >= min_count:
                    return True
    return False


def _find_col_pairs(ws, ws_vals, closing_year):
    cy_str = str(closing_year)
    py_str = str(closing_year - 1)
    cy_cols, py_cols = set(), set()

    for row in ws.iter_rows(max_row=15):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            val = str(cell.value) if cell.value else ""
            if not any(kw.lower() in val.lower() for kw in DATE_KEYWORDS):
                continue
            if cy_str in val:
                cy_cols.add(cell.column)
            if py_str in val:
                py_cols.add(cell.column)

    cy_data = sorted(c for c in cy_cols if _col_has_values(ws_vals, c))
    py_data = sorted(c for c in py_cols if _col_has_values(ws_vals, c))

    pairs = []
    used = set()
    for cy in cy_data:
        candidates = [p for p in py_data if p > cy and p not in used]
        if candidates:
            best = min(candidates)
            pairs.append((get_column_letter(cy), get_column_letter(best)))
            used.add(best)
    return pairs


def _snapshot_py_formulas(ws, py_letter):
    py_col = column_index_from_string(py_letter)
    formulas = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.column == py_col and _is_formula(cell.value):
                formulas[cell.row] = cell.value
    return formulas


def _copy_cy_to_py(ws, ws_vals, cy_letter, py_letter):
    cy_col = column_index_from_string(cy_letter)
    py_col = column_index_from_string(py_letter)
    for row_idx in range(1, ws.max_row + 1):
        cy_val  = ws_vals.cell(row=row_idx, column=cy_col).value
        py_cell = ws.cell(row=row_idx, column=py_col)
        if isinstance(py_cell, MergedCell):
            continue
        if cy_val is None:
            continue
        if isinstance(cy_val, str) and cy_val.strip() == "":
            continue
        # Skip datetime objects - they are not financial data
        if isinstance(cy_val, (datetime.datetime, datetime.date)):
            continue
        py_cell.value = cy_val


def _restore_py_formulas(ws, py_letter, formulas):
    py_col = column_index_from_string(py_letter)
    for row_idx, formula in formulas.items():
        cell = ws.cell(row=row_idx, column=py_col)
        if not isinstance(cell, MergedCell):
            cell.value = formula


def _clear_cy_constants(ws, cy_letter):
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
        # Never clear datetime cells - they are asset purchase dates etc.
        if isinstance(val, (datetime.datetime, datetime.date)):
            continue
        if isinstance(val, str):
            if val.strip() == "":
                continue
            try:
                float(val.replace(",", "").strip())
            except ValueError:
                continue
        cell.value = None


def _update_text(ws, pairs):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and not _is_formula(cell.value):
                new_val = _replace_text(cell.value, pairs)
                if new_val != cell.value:
                    cell.value = new_val


def _fix_py_headers(ws, cy_letter, py_letter, closing_year, new_year):
    cy_col = column_index_from_string(cy_letter)
    py_col = column_index_from_string(py_letter)
    ny_str = str(new_year)
    cy_str = str(closing_year)

    cy_header_texts = {}
    for row_idx in range(1, 31):
        cell = ws.cell(row=row_idx, column=cy_col)
        if isinstance(cell, MergedCell):
            continue
        val = str(cell.value) if cell.value else ""
        if ny_str in val and any(kw.lower() in val.lower() for kw in DATE_KEYWORDS):
            cy_header_texts[row_idx] = val

    for row_idx in range(1, 31):
        py_cell = ws.cell(row=row_idx, column=py_col)
        if isinstance(py_cell, MergedCell):
            continue
        val = py_cell.value
        if val is None:
            continue
        val_str = str(val)
        if not _is_formula(val) and ny_str in val_str:
            if any(kw.lower() in val_str.lower() for kw in DATE_KEYWORDS):
                py_cell.value = val_str.replace(ny_str, cy_str)
                continue
        if _is_formula(val):
            for check_row in [row_idx, row_idx-1, row_idx+1, row_idx-2, row_idx+2]:
                if check_row in cy_header_texts:
                    py_cell.value = cy_header_texts[check_row].replace(ny_str, cy_str)
                    break

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.column != py_col:
                continue
            val = cell.value
            if val is None or _is_formula(val):
                continue
            val_str = str(val)
            if ny_str in val_str and any(kw.lower() in val_str.lower() for kw in DATE_KEYWORDS):
                cell.value = val_str.replace(ny_str, cy_str)


def _convert_xls_to_xlsx(xls_path: str) -> str:
    """Convert .xls to .xlsx using LibreOffice. Returns path to converted file."""
    out_dir = tempfile.mkdtemp()
    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "xlsx",
             "--outdir", out_dir, xls_path],
            capture_output=True, timeout=60
        )
        base = os.path.splitext(os.path.basename(xls_path))[0]
        xlsx_path = os.path.join(out_dir, base + ".xlsx")
        if os.path.exists(xlsx_path):
            return xlsx_path
    except Exception:
        pass
    raise ValueError("Could not convert .xls file. Please open in Excel and Save As .xlsx first.")


def process(input_path: str, output_path: str, closing_year: int, new_year: int) -> dict:
    """
    Universal balance sheet year-shift.
    Supports both .xlsx and .xls input files.
    """
    # Handle .xls conversion
    work_path = input_path
    converted = False
    if input_path.lower().endswith(".xls"):
        work_path = _convert_xls_to_xlsx(input_path)
        converted = True

    pairs   = _date_pairs(closing_year, new_year)
    wb      = load_workbook(work_path)
    wb_vals = load_workbook(work_path, data_only=True)
    log     = []

    for sheet_name in wb.sheetnames:
        ws      = wb[sheet_name]
        ws_vals = wb_vals[sheet_name]
        col_pairs = _find_col_pairs(ws, ws_vals, closing_year)

        if col_pairs:
            for cy_letter, py_letter in col_pairs:
                py_formulas = _snapshot_py_formulas(ws, py_letter)
                _copy_cy_to_py(ws, ws_vals, cy_letter, py_letter)
                _restore_py_formulas(ws, py_letter, py_formulas)
                _clear_cy_constants(ws, cy_letter)

            _update_text(ws, pairs)

            for cy_letter, py_letter in col_pairs:
                _fix_py_headers(ws, cy_letter, py_letter, closing_year, new_year)

            desc = ", ".join(f"{c}→{p}" for c, p in col_pairs)
            log.append(f"✓ {sheet_name}: [{desc}] shifted, formulas kept, dates updated")
        else:
            _update_text(ws, pairs)
            log.append(f"  {sheet_name}: date text updated")

    wb.save(output_path)

    # Cleanup converted temp file
    if converted:
        try:
            os.remove(work_path)
            os.rmdir(os.path.dirname(work_path))
        except Exception:
            pass

    return {"status": "success", "log": log, "output": output_path}


if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("Usage: python processor.py input.xlsx output.xlsx closing_year new_year")
        sys.exit(1)
    result = process(sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]))
    for line in result["log"]:
        print(line)
    print(f"\nSaved → {result['output']}")
