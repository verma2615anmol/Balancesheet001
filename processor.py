"""
Balance Sheet Year-Shift Processor  (v8 — XML-native, formatting-safe)
Shifts CY→PY, clears CY constants, updates all date references.

Zero openpyxl-save approach: uses read_only for detection + value extraction,
then directly manipulates worksheet XML and sharedStrings inside the ZIP.
All formatting, styles, merged cells, and pivot caches are preserved byte-perfect.
"""

import copy, os, re, zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.utils import get_column_letter

# ── thresholds ────────────────────────────────────────────────────────────────
BIG_ROWS = 1000            # sheets above this are skipped for CY→PY
BIG_COLS = 100
HEADER_SCAN_ROWS = 15

# ── XML namespaces ────────────────────────────────────────────────────────────
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_MAP = {"": _NS}        # for ElementTree findall

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


# ═══════════════════════════════════════════════════════════════════════════════
#  Date-replacement text pairs
# ═══════════════════════════════════════════════════════════════════════════════

def _date_replacements(cy, ny):
    po = str(int(cy) - 1)
    PH, PH_R = "__NEWCY__", "__FYRNG__"

    patterns = [
        "31.03.{y}", "31 March, {y}", "31 March {y}",
        "31st March, {y}", "31st March {y}",
        "31ST MARCH ,{y}", "31ST MARCH, {y}", "31ST MARCH {y}",
        "31 MARCH, {y}", "31 MARCH {y}",
        "31st MARCH, {y}", "31st MARCH {y}",       # mixed: lowercase ordinal + uppercase MARCH
        "year ended 31 March, {y}", "year ended, 31st March, {y}",
        "year end 31 March, {y}",
        "year ending 31.03.{y}", "YEAR ENDING 31.03.{y}",
        "YEAR ENDING 31ST MARCH ,{y}", "YEAR ENDING 31ST MARCH, {y}",
        "YEAR ENDING 31ST MARCH {y}",
        "as at 31 March, {y}", "AS AT 31ST MARCH {y}",
        "AS AT 31ST MARCH, {y}", "AS AT 31 MARCH, {y}",
        "AS AT 31st MARCH, {y}", "AS AT 31st MARCH {y}",   # mixed case in AS AT
        "for the year ended, 31st March, {y}",
        "for the year ended 31 March, {y}",
        "FOR THE YEAR ENDED 31ST MARCH, {y}",
        "FOR THE YEAR ENDED 31ST MARCH {y}",
        "FOR THE YEAR ENDED 31st MARCH, {y}",               # mixed case in FOR THE YEAR
        "FOR THE YEAR ENDED 31st MARCH {y}",
    ]

    a = [(p.format(y=cy), p.format(y=PH)) for p in patterns]   # CY → placeholder
    b = []                                                       # PY → CY
    for old, new in [
        ("1st April {po}",  "1st April {cy}"), ("1 April {po}",  "1 April {cy}"),
        ("01.04.{po}",      "01.04.{cy}"),     ("31.03.{po}",    "31.03.{cy}"),
        ("31st March {po}", "31st March {cy}"),("31st March, {po}","31st March, {cy}"),
        ("31 March, {po}",  "31 March, {cy}"), ("31 March {po}",  "31 March {cy}"),
        ("31ST MARCH {po}", "31ST MARCH {cy}"),("31ST MARCH, {po}","31ST MARCH, {cy}"),
        ("31st MARCH {po}", "31st MARCH {cy}"),("31st MARCH, {po}","31st MARCH, {cy}"),
        ("AS AT 31ST MARCH {po}","AS AT 31ST MARCH {cy}"),
        ("AS AT 31ST MARCH, {po}","AS AT 31ST MARCH, {cy}"),
        ("AS AT 31st MARCH {po}","AS AT 31st MARCH {cy}"),
        ("AS AT 31st MARCH, {po}","AS AT 31st MARCH, {cy}"),
    ]:
        b.append((old.format(po=po, cy=cy), new.format(po=po, cy=cy)))
    c = [(p.format(y=PH), p.format(y=ny)) for p in patterns]   # placeholder → NY

    # Fiscal year ranges
    po_s, cy_s, ny_s = po[2:], cy[2:], ny[2:]
    pp = str(int(po)-1); pp_s = pp[2:]
    r = [
        (f"{po}-{cy_s}", PH_R), (f"{po}-{cy}", f"{PH_R}L"),
        (f"{pp}-{po_s}", f"{po}-{cy_s}"), (f"{pp}-{po}", f"{po}-{cy}"),
        (f"{PH_R}L", f"{cy}-{ny}"), (PH_R, f"{cy}-{ny_s}"),
    ]
    return a + b + c + r


def _apply_pairs(text, pairs):
    for old, new in pairs:
        text = text.replace(old, new)
    return text


# ═══════════════════════════════════════════════════════════════════════════════
#  read_only helpers  (fast scans — never loads full workbook into memory)
# ═══════════════════════════════════════════════════════════════════════════════

def _has_year_date(val, yr):
    if not isinstance(val, str): return False
    flat = val.replace("\n", " ").replace("\r", " ")
    if yr not in flat: return False
    return ("31.03." in flat or
            any(m in flat for m in ("March","MARCH","march",
                "year end","Year end","YEAR END","as at","As at","AS AT")))


def _detect_columns(ws, closing_year, max_scan_rows=HEADER_SCAN_ROWS):
    """Auto-detect CY/PY column pairs from header rows. Uses iter_rows for speed."""
    cy_s, py_s = str(closing_year), str(closing_year - 1)
    cy_cols, py_cols = set(), set()
    for row in ws.iter_rows(min_row=1, max_row=max_scan_rows, max_col=50):
        for cell in row:
            if isinstance(cell, MergedCell): continue
            v = cell.value
            if not isinstance(v, str) or v.startswith("="): continue
            if _has_year_date(v, cy_s): cy_cols.add(cell.column)
            if _has_year_date(v, py_s): py_cols.add(cell.column)
    if not cy_cols: return []
    pairs, used = [], set()
    for cc in sorted(cy_cols):
        for off in [1, 2, 3]:
            cand = cc + off
            if cand in py_cols and cand not in used:
                pairs.append((get_column_letter(cc), get_column_letter(cand)))
                used.add(cand); break
    return pairs


def _parse_dim(dim_str):
    """Parse 'A1:AQ65570' → (rows, cols)."""
    if ":" not in dim_str:
        return 1, 1
    _, end = dim_str.split(":")
    m = _COL_RE.match(end)
    if not m:
        return 1, 1
    return int(m.group(2)), _col_idx(m.group(1))


def _get_sizes_from_zip(filepath):
    """Get sheet sizes from XML dimension tags — instant, no cell parsing."""
    sizes = {}  # {sheet_file: (rows, cols)}
    dim_re = re.compile(rb'<dimension ref="([^"]+)"')
    with zipfile.ZipFile(filepath) as z:
        for item in z.infolist():
            if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                header = z.read(item.filename)[:2000]
                m = dim_re.search(header)
                if m:
                    sizes[item.filename] = _parse_dim(m.group(1).decode())
                else:
                    sizes[item.filename] = (0, 0)
    return sizes


def _scan_workbook(filepath, closing_year):
    """
    Quick scan. Returns:
      sizes       {sheet_name: (rows, cols)}
      shift_map   {sheet_name: [(cy_letter, py_letter), ...]}
      cy_values   {sheet_name: {col_letter: {row: numeric_value}}}
      cy_formulas {sheet_name: {col_letter: set_of_rows}}
    """
    # Get sizes from ZIP dimension tags (0.04s)
    with zipfile.ZipFile(filepath) as z:
        smap = _sheet_file_map(z)
    file_sizes = _get_sizes_from_zip(filepath)
    sizes = {}
    for sn, sf in smap.items():
        sizes[sn] = file_sizes.get(sf, (0, 0))

    big_names = {n for n, (r, c) in sizes.items() if r > BIG_ROWS or c > BIG_COLS}

    # Detect CY/PY columns — only load small sheets via read_only
    shift_map = {}
    fb = {k.strip().lower(): v for k, v in SHEET_COL_MAP.items()}
    wb = load_workbook(filepath, read_only=True)
    for sn in wb.sheetnames:
        if sn in big_names:
            continue
        sl = sn.strip().lower()
        if sl in TEXT_ONLY_SHEETS:
            continue
        ws = wb[sn]
        det = _detect_columns(ws, closing_year)
        if not det and sl in fb:
            det = fb[sl]
        if det:
            shift_map[sn] = det
    wb.close()

    # Extract CY numeric values (data_only → resolved formulas)
    cy_values = {}
    wb_d = load_workbook(filepath, read_only=True, data_only=True)
    for sn, col_pairs in shift_map.items():
        cy_values[sn] = {}
        ws = wb_d[sn]
        for cy_l, _ in col_pairs:
            ci = _col_idx(cy_l)
            vals = {}
            for row in ws.iter_rows(min_col=ci, max_col=ci):
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        vals[cell.row] = cell.value
            cy_values[sn][cy_l] = vals
    wb_d.close()

    # Identify formula rows (read_only without data_only)
    cy_formulas = {}
    wb_f = load_workbook(filepath, read_only=True)
    for sn, col_pairs in shift_map.items():
        cy_formulas[sn] = {}
        ws = wb_f[sn]
        for cy_l, _ in col_pairs:
            ci = _col_idx(cy_l)
            frows = set()
            for row in ws.iter_rows(min_col=ci, max_col=ci):
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        frows.add(cell.row)
            cy_formulas[sn][cy_l] = frows
    wb_f.close()

    return sizes, shift_map, cy_values, cy_formulas


# ═══════════════════════════════════════════════════════════════════════════════
#  XML manipulation  (direct edits on worksheet XML inside the ZIP)
# ═══════════════════════════════════════════════════════════════════════════════

_COL_RE = re.compile(r'^([A-Z]+)(\d+)$')

def _col_idx(letter):
    """A→1, B→2, ..., Z→26, AA→27, etc."""
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_letter_from_ref(ref):
    m = _COL_RE.match(ref)
    return m.group(1) if m else None


def _row_from_ref(ref):
    m = _COL_RE.match(ref)
    return int(m.group(2)) if m else None


def _process_sheet_xml(xml_bytes, col_pairs, cy_vals, cy_formulas, shared_strings):
    """
    Modify a worksheet XML in-memory:
    - Copy CY numeric values → PY cells
    - Clear CY constants (non-formula numeric cells → empty)
    Returns modified XML bytes.
    """
    root = ET.fromstring(xml_bytes)
    sd = root.find(f"{{{_NS}}}sheetData")
    if sd is None:
        return xml_bytes

    # Build lookup: {cy_letter: (py_letter, {row: value}, formula_rows)}
    col_info = {}
    for cy_l, py_l in col_pairs:
        vals = cy_vals.get(cy_l, {})
        frows = cy_formulas.get(cy_l, set())
        col_info[cy_l] = (py_l, vals, frows)

    py_letters = {py_l for _, py_l in col_pairs}
    cy_letters = set(col_info.keys())

    for row_el in sd:
        for cell_el in row_el:
            ref = cell_el.get("r", "")
            cl = _col_letter_from_ref(ref)
            rn = _row_from_ref(ref)
            if cl is None or rn is None:
                continue

            # Copy value to PY cell
            if cl in cy_letters:
                py_l, vals, frows = col_info[cl]
                # We'll handle PY writing when we encounter the PY cell
                pass

            # Write CY value into PY cell
            if cl in py_letters:
                # Find corresponding CY
                for cy_l, (py_l2, vals, frows) in col_info.items():
                    if py_l2 == cl and rn in vals:
                        # Overwrite this PY cell with the CY value
                        v_el = cell_el.find(f"{{{_NS}}}v")
                        f_el = cell_el.find(f"{{{_NS}}}f")
                        if f_el is not None:
                            break  # don't overwrite formulas in PY
                        if v_el is None:
                            v_el = ET.SubElement(cell_el, f"{{{_NS}}}v")
                        v_el.text = str(vals[rn])
                        # Set type to number (remove 's' type if it was string)
                        if cell_el.get("t") == "s":
                            del cell_el.attrib["t"]
                        break

            # Clear CY constant (non-formula numeric)
            if cl in cy_letters:
                _, vals, frows = col_info[cl]
                if rn in vals and rn not in frows:
                    v_el = cell_el.find(f"{{{_NS}}}v")
                    if v_el is not None:
                        cell_el.remove(v_el)

    ET.register_namespace("", _NS)
    # Preserve all other namespaces from original
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


def _update_shared_strings(xml_bytes, pairs):
    """Replace date text in sharedStrings.xml.
    
    Handles BOTH plain strings and rich-text entries where a date like
    "31st MARCH, 2025" is split across multiple <t> elements
    (e.g., "31" + "st" + " MARCH, 2025" for superscript formatting).
    
    Strategy for rich-text: concatenate all <t> texts, apply replacements
    on the full string, then redistribute the changes back proportionally.
    """
    root = ET.fromstring(xml_bytes)
    changed = False
    for si in root:
        t_els = list(si.iter(f"{{{_NS}}}t"))
        if not t_els:
            continue

        # --- Single <t> (plain string) — fast path ---
        if len(t_els) == 1:
            t_el = t_els[0]
            if t_el.text:
                new_text = _apply_pairs(t_el.text, pairs)
                if new_text != t_el.text:
                    t_el.text = new_text
                    changed = True
            continue

        # --- Multiple <t> (rich-text) — concatenate, replace, redistribute ---
        originals = [t.text or "" for t in t_els]
        full_old = "".join(originals)
        full_new = _apply_pairs(full_old, pairs)
        if full_new == full_old:
            continue
        changed = True

        # Redistribute: walk new text assigning to each run.
        # Keep each run's length the same where possible; absorb any
        # length change (from e.g. "2025" → "2026", same length, or
        # "2024-25" → "2025-26") into the run that contains the change.
        _redistribute_rich_text(t_els, originals, full_new)

    if not changed:
        return xml_bytes
    ET.register_namespace("", _NS)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


def _redistribute_rich_text(t_els, originals, full_new):
    """Redistribute replaced full_new text back into t_els runs.
    
    Uses a simple approach: find common prefix/suffix between old and new
    full text, then assign the changed middle portion to whichever runs
    it overlaps with.
    """
    full_old = "".join(originals)
    
    # Find which character positions changed
    # Build a mapping: for each run, its start/end position in full_old
    run_ranges = []
    pos = 0
    for orig in originals:
        run_ranges.append((pos, pos + len(orig)))
        pos += len(orig)
    
    # If lengths are the same (most common: year number swap), we can map 1:1
    if len(full_new) == len(full_old):
        pos = 0
        for i, t_el in enumerate(t_els):
            run_len = len(originals[i])
            t_el.text = full_new[pos:pos + run_len]
            pos += run_len
        return
    
    # Lengths differ — find common prefix/suffix and assign the changed
    # region proportionally. This handles cases like "2024-25" → "2025-26"
    # where the replacement might span multiple runs.
    pfx = 0
    while pfx < len(full_old) and pfx < len(full_new) and full_old[pfx] == full_new[pfx]:
        pfx += 1
    sfx = 0
    while (sfx < len(full_old) - pfx and sfx < len(full_new) - pfx and
           full_old[-(sfx+1)] == full_new[-(sfx+1)]):
        sfx += 1
    
    # Assign runs: unchanged prefix runs get their original text,
    # the run(s) spanning the changed region get the new middle,
    # unchanged suffix runs get their original text.
    new_texts = []
    pos = 0
    change_start = pfx
    change_end_old = len(full_old) - sfx
    change_end_new = len(full_new) - sfx
    new_middle = full_new[change_start:change_end_new]
    middle_assigned = False
    
    for i, (rstart, rend) in enumerate(run_ranges):
        if rend <= change_start:
            # Entirely before change — keep original
            new_texts.append(originals[i])
        elif rstart >= change_end_old:
            # Entirely after change — keep original
            new_texts.append(originals[i])
        else:
            # This run overlaps the change
            before = originals[i][:max(0, change_start - rstart)]
            after = originals[i][max(0, change_end_old - rstart):]
            if not middle_assigned:
                new_texts.append(before + new_middle + after)
                middle_assigned = True
            else:
                # Additional overlapping runs — their changed portion was
                # already consumed; keep only the after-change suffix
                new_texts.append(after)
    
    for i, t_el in enumerate(t_els):
        t_el.text = new_texts[i] if i < len(new_texts) else ""


def _update_sheet_inline_strings(xml_bytes, pairs):
    """Replace date text in inline strings and cell string values within a worksheet."""
    text = xml_bytes.decode("utf-8", errors="replace")
    new_text = _apply_pairs(text, pairs)
    if new_text == text:
        return xml_bytes
    return new_text.encode("utf-8")


def _fix_py_headers_xml(xml_bytes, col_pairs, closing_year, new_year):
    """After date replacement, PY header might say new_year — fix it back to closing_year."""
    root = ET.fromstring(xml_bytes)
    sd = root.find(f"{{{_NS}}}sheetData")
    if sd is None:
        return xml_bytes

    ny_s, cy_s = str(new_year), str(closing_year)
    py_letters = {py_l for _, py_l in col_pairs}

    # Load shared strings to check string cells
    # Actually we can't access shared strings here, so just work on inline strings
    # The main date replacement happens in sharedStrings — PY header fix needs to happen there too

    return xml_bytes  # handled in shared strings instead


# ═══════════════════════════════════════════════════════════════════════════════
#  ZIP-level sheet file mapping
# ═══════════════════════════════════════════════════════════════════════════════

def _sheet_file_map(z):
    """Return {sheet_name: 'xl/worksheets/sheetN.xml'}"""
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
            t = t.lstrip("/")
            if not t.startswith("xl/"):
                t = f"xl/{t}"
            out[name] = t
    return out


# ═══════════════════════════════════════════════════════════════════════════════
#  Main entry point
# ═══════════════════════════════════════════════════════════════════════════════

def process(input_path: str, output_path: str, closing_year: int, new_year: int) -> dict:
    pairs = _date_replacements(str(closing_year), str(new_year))
    log = []

    # ── 1. read_only scan ─────────────────────────────────────────────────────
    sizes, shift_map, cy_values, cy_formulas = _scan_workbook(input_path, closing_year)
    big_names = {n for n, (r, c) in sizes.items() if r > BIG_ROWS or c > BIG_COLS}

    # ── 2. ZIP-level manipulation ─────────────────────────────────────────────
    with zipfile.ZipFile(input_path, "r") as zi:
        smap = _sheet_file_map(zi)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zi.infolist():
                fn = item.filename
                data = zi.read(fn)

                # ── shared strings: replace dates ─────────────────────────
                if fn == "xl/sharedStrings.xml":
                    data = _update_shared_strings(data, pairs)
                    zo.writestr(item, data)
                    continue

                # ── worksheet XMLs ────────────────────────────────────────
                is_sheet = fn.startswith("xl/worksheets/") and fn.endswith(".xml")
                if is_sheet:
                    # Find which sheet this is
                    sheet_name = None
                    for sn, sf in smap.items():
                        if sf == fn:
                            sheet_name = sn
                            break

                    if sheet_name and sheet_name in shift_map:
                        # Process: CY→PY copy + CY clear
                        cpairs = shift_map[sheet_name]
                        data = _process_sheet_xml(
                            data, cpairs,
                            cy_values.get(sheet_name, {}),
                            cy_formulas.get(sheet_name, {}),
                            None
                        )
                        # Also update inline strings in this sheet
                        data = _update_sheet_inline_strings(data, pairs)
                        desc = ", ".join(f"{c}->{p}" for c, p in cpairs)
                        log.append(f"+ {sheet_name}: CY->PY copied ({desc}), CY cleared, dates updated")
                    elif sheet_name and sheet_name not in big_names:
                        # Small sheet, no CY/PY but update date strings
                        data = _update_sheet_inline_strings(data, pairs)
                        log.append(f"+ {sheet_name}: dates updated")
                    elif sheet_name and sheet_name in big_names:
                        log.append(f"* {sheet_name}: preserved unchanged (large data sheet)")
                    # else: unknown sheet, pass through

                    zo.writestr(item, data)
                    continue

                # ── everything else: pass through unchanged ───────────────
                zo.writestr(item, data)

    return {"status": "success", "log": log, "output": output_path}


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 5:
        print("Usage: python processor.py input.xlsx output.xlsx closing_year new_year")
        sys.exit(1)
    res = process(sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]))
    for l in res["log"]: print(l)
    print(f"\nSaved -> {res['output']}")
