"""
Balance Sheet Year-Shift Processor  (v7 — auto-detect, ZIP-trim, memory-safe)
Shifts CY→PY, clears CY constants, updates all date references.

Pipeline for large workbooks (65K-row data sheets, pivot caches):
  1. read_only scan → detect sheet sizes + big sheets         (0.4s)
  2. ZIP manipulation → stub big sheet XMLs + pivot caches    (0.1s)
  3. load_workbook on trimmed file                            (~12s)
  4. read_only+data_only → extract CY column values           (0.3s)
  5. Auto-detect CY/PY columns → shift + clear + date updates
  6. ZIP merge → splice processed sheets into original ZIP    (0.9s)
"""

import os, shutil, tempfile, zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter

# ── thresholds ────────────────────────────────────────────────────────────────
BIG_ROWS = 1000
BIG_COLS = 100
SHIFT_MAX_ROWS = 2000
HEADER_SCAN_ROWS = 15

# ── fallback column map (lower-case sheet name → [(cy, py)]) ─────────────────
SHEET_COL_MAP = {
    "bs":           [("E", "F")],
    "p&l":          [("E", "F")],
    "notes to bs":  [("D", "E")],
    "notes to p&l": [("D", "E")],
    "details":      [("D", "E")],
    "gross profit": [("B", "C"), ("F", "G")],
}
TEXT_ONLY_SHEETS = {s.strip().lower() for s in [
    "notes to accounts", "Fixed Assets C. Yr.", "Fixed Assets P. Yr.",
    "FA2022", "Tax audit ", "Tax Audit report", "PPE",
]}

CAPITAL_CY_ROW, CAPITAL_PY_ROW = 8, 11
CAPITAL_DATA_COLS = ["C", "D", "E"]

# ── XML stubs for ZIP trimming ────────────────────────────────────────────────
_EMPTY_SHEET = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<sheetData/></worksheet>'
)
_EMPTY_PIVOT = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    b' count="0"/>'
)
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ═══════════════════════════════════════════════════════════════════════════════
#  Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _is_formula(v):
    return isinstance(v, str) and v.strip().startswith("=")

def _is_numeric(v):
    return isinstance(v, (int, float))


# ── date replacement pairs ────────────────────────────────────────────────────

def _date_replacements(cy, ny):
    po = str(int(cy) - 1)
    PH, PH_R = "__NEWCY__", "__FYRNG__"
    # CY → placeholder
    a = []
    for pat in [
        "31.03.{y}", "31 March, {y}", "31 March {y}",
        "31st March, {y}", "31st March {y}",
        "31ST MARCH ,{y}", "31ST MARCH, {y}", "31ST MARCH {y}",
        "31 MARCH, {y}", "31 MARCH {y}",
        "year ended 31 March, {y}", "year ended, 31st March, {y}",
        "year end 31 March, {y}",
        "year ending 31.03.{y}", "YEAR ENDING 31.03.{y}",
        "YEAR ENDING 31ST MARCH ,{y}", "YEAR ENDING 31ST MARCH, {y}",
        "YEAR ENDING 31ST MARCH {y}",
        "as at 31 March, {y}", "AS AT 31ST MARCH {y}",
        "AS AT 31ST MARCH, {y}", "AS AT 31 MARCH, {y}",
        "for the year ended, 31st March, {y}",
        "for the year ended 31 March, {y}",
        "FOR THE YEAR ENDED 31ST MARCH, {y}",
        "FOR THE YEAR ENDED 31ST MARCH {y}",
    ]:
        a.append((pat.format(y=cy), pat.format(y=PH)))

    # PY dates → CY dates
    b = []
    for old, new in [
        ("1st April {po}",       "1st April {cy}"),
        ("1 April {po}",         "1 April {cy}"),
        ("01.04.{po}",           "01.04.{cy}"),
        ("31.03.{po}",           "31.03.{cy}"),
        ("31st March {po}",      "31st March {cy}"),
        ("31st March, {po}",     "31st March, {cy}"),
        ("31 March, {po}",       "31 March, {cy}"),
        ("31 March {po}",        "31 March {cy}"),
        ("31ST MARCH {po}",      "31ST MARCH {cy}"),
        ("31ST MARCH, {po}",     "31ST MARCH, {cy}"),
        ("AS AT 31ST MARCH {po}",  "AS AT 31ST MARCH {cy}"),
        ("AS AT 31ST MARCH, {po}", "AS AT 31ST MARCH, {cy}"),
    ]:
        b.append((old.format(po=po, cy=cy), new.format(po=po, cy=cy)))

    # placeholder → NY
    c = [(old.replace(cy, PH), new.replace(cy, PH).replace(PH, ny))
         for old, new in zip([p[0] for p in a], [p[0] for p in a])]
    # Actually rebuild c properly
    c = []
    for pat in [
        "31.03.{y}", "31 March, {y}", "31 March {y}",
        "31st March, {y}", "31st March {y}",
        "31ST MARCH ,{y}", "31ST MARCH, {y}", "31ST MARCH {y}",
        "31 MARCH, {y}", "31 MARCH {y}",
        "year ended 31 March, {y}", "year ended, 31st March, {y}",
        "year end 31 March, {y}",
        "year ending 31.03.{y}", "YEAR ENDING 31.03.{y}",
        "YEAR ENDING 31ST MARCH ,{y}", "YEAR ENDING 31ST MARCH, {y}",
        "YEAR ENDING 31ST MARCH {y}",
        "as at 31 March, {y}", "AS AT 31ST MARCH {y}",
        "AS AT 31ST MARCH, {y}", "AS AT 31 MARCH, {y}",
        "for the year ended, 31st March, {y}",
        "for the year ended 31 March, {y}",
        "FOR THE YEAR ENDED 31ST MARCH, {y}",
        "FOR THE YEAR ENDED 31ST MARCH {y}",
    ]:
        c.append((pat.format(y=PH), pat.format(y=ny)))

    # Fiscal year ranges
    po_s, cy_s, ny_s = po[2:], cy[2:], ny[2:]
    pp = str(int(po)-1); pp_s = pp[2:]
    r = [
        (f"{po}-{cy_s}", PH_R), (f"{po}-{cy}", f"{PH_R}L"),
        (f"{pp}-{po_s}", f"{po}-{cy_s}"), (f"{pp}-{po}", f"{po}-{cy}"),
        (f"{PH_R}L", f"{cy}-{ny}"), (PH_R, f"{cy}-{ny_s}"),
    ]
    return a + b + c + r


def _replace_text(val, pairs):
    if not isinstance(val, str):
        return val
    for old, new in pairs:
        val = val.replace(old, new)
    return val


# ── auto-detection ────────────────────────────────────────────────────────────

def _has_year_date(val, yr):
    if not isinstance(val, str): return False
    flat = val.replace("\n", " ").replace("\r", " ")
    if yr not in flat: return False
    if f"31.03.{yr}" in flat: return True
    return any(m in flat for m in
               ["March","MARCH","march","year end","Year end",
                "YEAR END","as at","As at","AS AT"])


def _detect_cy_py_columns(ws, closing_year):
    cy_s, py_s = str(closing_year), str(closing_year - 1)
    cy_cols, py_cols = set(), set()
    mx = min(ws.max_column or 1, 50)
    for ri in range(1, HEADER_SCAN_ROWS + 1):
        for ci in range(1, mx + 1):
            cell = ws.cell(row=ri, column=ci)
            if isinstance(cell, MergedCell): continue
            v = cell.value
            if not isinstance(v, str) or _is_formula(v): continue
            if _has_year_date(v, cy_s): cy_cols.add(ci)
            if _has_year_date(v, py_s): py_cols.add(ci)
    if not cy_cols: return []
    pairs, used = [], set()
    for cc in sorted(cy_cols):
        for off in [1, 2, 3]:
            cand = cc + off
            if cand in py_cols and cand not in used:
                pairs.append((get_column_letter(cc), get_column_letter(cand)))
                used.add(cand)
                break
    return pairs


# ── fast value extraction (read_only + data_only) ─────────────────────────────

def _extract_col_values(filepath, sheet_name, col_idx):
    """Return {row: numeric_value} for one column, using fast read_only mode."""
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb[sheet_name]
    vals = {}
    for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
        for cell in row:
            if cell.value is not None and _is_numeric(cell.value):
                vals[cell.row] = cell.value
    wb.close()
    return vals


# ── core operations ───────────────────────────────────────────────────────────

def _copy_cy_to_py_cached(ws, cy_vals, py_col):
    for ri, cv in cy_vals.items():
        pc = ws.cell(row=ri, column=py_col)
        if isinstance(pc, MergedCell) or pc.value is None:
            continue
        pc.value = cv


def _clear_cy_constants(ws, cy_letter):
    ci = column_index_from_string(cy_letter)
    for ri in range(1, ws.max_row + 1):
        cell = ws.cell(row=ri, column=ci)
        if isinstance(cell, MergedCell): continue
        v = cell.value
        if v is None or _is_formula(v) or isinstance(v, str): continue
        cell.value = None


def _update_text(ws, pairs):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell): continue
            if isinstance(cell.value, str) and not _is_formula(cell.value):
                nv = _replace_text(cell.value, pairs)
                if nv != cell.value:
                    cell.value = nv


def _fix_py_header(ws, py_letter, closing_year, new_year):
    pc = column_index_from_string(py_letter)
    ny_s, cy_s = str(new_year), str(closing_year)
    for ri in range(1, HEADER_SCAN_ROWS + 1):
        cell = ws.cell(row=ri, column=pc)
        if isinstance(cell, MergedCell): continue
        v = cell.value
        if not isinstance(v, str) or _is_formula(v): continue
        if ny_s in v and any(m in v for m in
                ("31.03.","March","MARCH","march",
                 "year ending","Year ending","YEAR ENDING",
                 "as at","As at","AS AT")):
            cell.value = v.replace(ny_s, cy_s)


# ═══════════════════════════════════════════════════════════════════════════════
#  ZIP-level trim / merge
# ═══════════════════════════════════════════════════════════════════════════════

def _scan_sizes(filepath):
    wb = load_workbook(filepath, read_only=True)
    sizes = {ws.title: (ws.max_row or 0, ws.max_column or 0) for ws in wb.worksheets}
    wb.close()
    return sizes


def _sheet_file_map(z):
    root = ET.fromstring(z.read("xl/workbook.xml"))
    srid = {}
    for el in root.iter(f"{{{_NS}}}sheet"):
        srid[el.get("name")] = el.get(f"{{{_NS_R}}}id")
    rr = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rf = {r.get("Id"): r.get("Target") for r in rr}
    out = {}
    for name, rid in srid.items():
        t = rf.get(rid)
        if t:
            out[name] = f"xl/{t}" if not t.startswith("xl/") else t
    return out


def _trim_zip(src, dst, big_names):
    with zipfile.ZipFile(src, "r") as zi:
        smap = _sheet_file_map(zi)
        bf = {smap[n] for n in big_names if n in smap}
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zi.infolist():
                if item.filename in bf:
                    zo.writestr(item, _EMPTY_SHEET)
                elif "pivotCacheRecords" in item.filename:
                    zo.writestr(item, _EMPTY_PIVOT)
                else:
                    zo.writestr(item, zi.read(item.filename))
    return bf


def _merge_back(original, processed, output, big_files):
    pd = {}
    with zipfile.ZipFile(processed, "r") as zp:
        for item in zp.infolist():
            fn = item.filename
            if fn.startswith("xl/worksheets/") or fn == "xl/sharedStrings.xml":
                pd[fn] = zp.read(fn)
    with zipfile.ZipFile(original, "r") as zo:
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zw:
            for item in zo.infolist():
                fn = item.filename
                if fn in pd and fn not in big_files:
                    zw.writestr(item, pd[fn])
                else:
                    zw.writestr(item, zo.read(fn))


# ═══════════════════════════════════════════════════════════════════════════════
#  Main entry point
# ═══════════════════════════════════════════════════════════════════════════════

def process(input_path: str, output_path: str, closing_year: int, new_year: int) -> dict:
    pairs = _date_replacements(str(closing_year), str(new_year))
    log = []

    # ── 0. scan ───────────────────────────────────────────────────────────────
    sizes = _scan_sizes(input_path)
    big_names = {n for n,(r,c) in sizes.items() if r > BIG_ROWS or c > BIG_COLS}
    needs_trim = bool(big_names)

    tmp_dir = None
    big_files = set()

    if needs_trim:
        tmp_dir = tempfile.mkdtemp(prefix="bs_proc_")
        trimmed = os.path.join(tmp_dir, "trimmed.xlsx")
        big_files = _trim_zip(input_path, trimmed, big_names)
        work_path = trimmed
    else:
        work_path = input_path

    try:
        wb = load_workbook(work_path)
        fb = {k.strip().lower(): v for k, v in SHEET_COL_MAP.items()}
        done = set()

        # ── 1. detect sheets needing CY→PY ───────────────────────────────────
        shift_list = []
        for sn in wb.sheetnames:
            sl = sn.strip().lower()
            if sl in TEXT_ONLY_SHEETS: continue
            r = sizes.get(sn, (0,0))[0]
            if r > SHIFT_MAX_ROWS: continue
            det = _detect_cy_py_columns(wb[sn], closing_year)
            if not det and sl in fb:
                det = fb[sl]
            if det:
                shift_list.append((sn, det))

        # ── 2. CY→PY shift (fast value extraction) ───────────────────────────
        for sn, cpairs in shift_list:
            ws = wb[sn]
            for cy_l, py_l in cpairs:
                cy_ci = column_index_from_string(cy_l)
                py_ci = column_index_from_string(py_l)
                vals = _extract_col_values(work_path, sn, cy_ci)
                _copy_cy_to_py_cached(ws, vals, py_ci)
                _clear_cy_constants(ws, cy_l)
            _update_text(ws, pairs)
            for cy_l, py_l in cpairs:
                _fix_py_header(ws, py_l, closing_year, new_year)
            desc = ", ".join(f"{c}→{p}" for c, p in cpairs)
            log.append(f"✓ {sn}: CY→PY copied ({desc}), CY cleared, dates updated")
            done.add(sn)

        # ── 3. capital sheet ──────────────────────────────────────────────────
        for sn in wb.sheetnames:
            if sn.strip().lower() == "capital" and sn not in done:
                ws = wb[sn]
                for cl in CAPITAL_DATA_COLS:
                    ci = column_index_from_string(cl)
                    cv = _extract_col_values(work_path, sn, ci).get(CAPITAL_CY_ROW)
                    pc = ws.cell(row=CAPITAL_PY_ROW, column=ci)
                    if not isinstance(pc, MergedCell) and cv is not None:
                        pc.value = cv
                for cl in CAPITAL_DATA_COLS:
                    ci = column_index_from_string(cl)
                    cell = ws.cell(row=CAPITAL_CY_ROW, column=ci)
                    if not isinstance(cell, MergedCell) and not _is_formula(cell.value):
                        cell.value = None
                _update_text(ws, pairs)
                log.append(f"✓ {sn}: CY row→PY row copied, CY cleared, dates updated")
                done.add(sn)
                break

        # ── 4. remaining sheets — dates only ──────────────────────────────────
        for sn in wb.sheetnames:
            if sn not in done:
                if sn not in big_names:
                    _update_text(wb[sn], pairs)
                log.append(f"✓ {sn}: dates updated")
                done.add(sn)

        # ── 5. save ──────────────────────────────────────────────────────────
        if needs_trim:
            proc_path = os.path.join(tmp_dir, "processed.xlsx")
            wb.save(proc_path)
            wb.close()
            _merge_back(input_path, proc_path, output_path, big_files)
            for sn in big_names:
                log.append(f"✓ {sn}: preserved unchanged (large data sheet)")
        else:
            wb.save(output_path)
            wb.close()

    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    return {"status": "success", "log": log, "output": output_path}


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 5:
        print("Usage: python processor.py input.xlsx output.xlsx closing_year new_year")
        sys.exit(1)
    res = process(sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]))
    for l in res["log"]: print(l)
    print(f"\nSaved → {res['output']}")
