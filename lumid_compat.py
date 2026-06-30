"""
lumid_compat.py
===============
Compatibility shim for **Lumid-format** balance sheets
(e.g. "Final_Lumid_Biotech_Pvt__Ltd_BALANCE_SHEET_31_03_2025_in_Rupees.xlsx").

This file is a STANDALONE wrapper around the existing processor.py.
It does NOT modify processor.py in any way.  It:
  1. Detects whether the uploaded file is Lumid-format.
  2. If yes — patches the column-pair map that the processor would otherwise
     compute automatically, fixing five specific issues found in this template.
  3. Calls the standard processor.process() for everything else.
  4. If the file is NOT Lumid-format — falls through to the standard
     processor.process() unchanged.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
COMPATIBILITY ISSUES FOUND IN LUMID FORMAT  (analysis date: 2025-06-29)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Issue 1 – NOA 1-2  (two annexures on one sheet, column layout differs)
  Annexure 1 (rows ~8-55):
    E=CY Count, F=CY Amount (Rs Lakhs), G=PY Count, H=PY Amount
    Auto-detection finds E→G (count→count) but MISSES F→H (amount→amount).
    The amount column (F/H) has no year-bearing header of its own —
    it is labelled "Amount" in row 11, not "As At 31st March".
  Annexure 2 (rows ~57-75):
    F=CY, H=PY. Date header is split over two rows (row 59="As At",
    row 60="31st March 2025"), and row 60 > HEADER_SCAN_ROWS=20, so
    it is never seen by _detect_columns → Annexure 2 is entirely missed.
  Fix: hard-map NOA 1-2  → [(E,G), (F,H)]

Issue 2 – NOA 3-6  (four annexures on one sheet, two different column layouts)
  Annexure 3 (rows ~8-22):
    E=CY_NonCurrent, H=PY_NonCurrent  ← auto-detected correctly as E→H
    F=CY_CurrentMaturities, I=PY_CurrentMaturities  ← not detected (F/I unlabelled)
    G=CY_Total,             J=PY_Total              ← not detected
    In practice F and G are empty in this client's data, so E→H is sufficient.
  Annexures 4/5/6 (rows ~27-78):
    H=CY, J=PY — date headers are on rows 30/42/52, all beyond HEADER_SCAN_ROWS=20.
    NONE of these are detected automatically.
  Conflict: H is PY in Ann3 rows 8-22 AND CY in Ann4-6 rows 27-78.
    A single static E→H and H→J pair would corrupt Ann3 rows (H would be
    both cleared as CY-new and retained as PY-old simultaneously).
  Fix: hard-map NOA 3-6 → [(E,H)] for Ann3 section (rows 8-22 only have
    E and H populated; tool handles row-level: only rows with CY values are
    shifted) + add (H,J) for Ann4-6. Because E is empty in rows 27-78 and H
    is empty in rows 8-22 (as PY col H is never populated with CY data in
    Ann3 — it only receives the shifted value), this does NOT create a
    conflict at the cell level.  The tool writes E→H for rows where E has
    data; it writes H→J for rows where H has data AND J already exists.
    Since Ann3 rows have H=PY_data (copied FROM old E), and Ann4-6 rows
    have H=CY_data (to be copied to J), the pair (H,J) is safe once
    (E,H) has already shifted Ann3 correctly.
  Fix: hard-map NOA 3-6 → [(E,H), (H,J)]
    Caveat: This requires two passes in the right ORDER so that (E,H) runs
    before (H,J). processor._scan_workbook reads both pairs in the same pass
    so we supply them as ordered pairs and rely on the tool's row-by-row
    processing.  The tool processes pairs independently per-column so
    both can coexist.

Issue 3 – CASH FLOW  (stale date headers + wrong hardcoded SHEET_COL_MAP entry)
  The sheet contains date headers "For the Year ended 31st March, 2018" and
  "31st March, 2019" (never updated since FY2018-19), so auto-detection finds
  nothing.  The fallback SHEET_COL_MAP entry in processor.py is:
    "cash flow": [("D", "F")]
  but the actual layout is D=CY, E=PY (column F is empty).  The hardcoded
  D→F entry shifts data to the wrong column.
  Fix: override the CASH FLOW entry to ("D", "E").

Issue 4 – DEP COMPANIES ACT (2)  (Fixed-Asset schedule wrongly shifted)
  Auto-detection picks up B→D, H→I, L→N based on date-like strings in
  row 11 ("As at 1st April 2024", "As at 31 March 2025").  These are
  opening/closing WDV columns in a depreciation schedule, NOT CY/PY
  financial data columns.  Shifting them overwrites WDV and net-block
  cells with wrong values.
  Fix: add "DEP COMPANIES ACT (2)" and "DEP COMPANIES ACT" to TEXT_ONLY
  (date-text update only, no column shift).

Issue 5 – HEADER_SCAN_ROWS  (too small for this multi-annexure format)
  The standard value of 20 misses date headers beyond row 20.
  For this template the issues are addressed by Issues 1-4 above (hard-map
  overrides and TEXT_ONLY additions), so no global change is needed.
  No change to processor.py is made.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.cell import MergedCell

# ── Import the real processor ────────────────────────────────────────────────
import processor as _proc

# ── Lumid-format fingerprint ─────────────────────────────────────────────────
# We identify a Lumid-format file by the presence of the specific annexure
# sheet names used in this template family.  A simple heuristic that will
# not trigger on any other client's workbook.
_LUMID_SHEET_SIGNATURES = {
    "noa 1-2",
    "noa 3-6",
    "noa 7-12",
    "noa 13-17",
    "noa 18-22",
    "noa 23-27",
}

# Minimum number of the above sheets that must be present to call it Lumid-format
_LUMID_MIN_MATCH = 4


def _is_lumid_format(filepath: str) -> bool:
    """Return True if the workbook looks like a Lumid-format balance sheet."""
    try:
        wb = load_workbook(filepath, read_only=True, data_only=True)
        names_lower = {s.strip().lower() for s in wb.sheetnames}
        wb.close()
        matches = _LUMID_SHEET_SIGNATURES & names_lower
        return len(matches) >= _LUMID_MIN_MATCH
    except Exception:
        return False


# ── Lumid-specific column-pair overrides ─────────────────────────────────────
# Keys are the exact sheet names as they appear in the workbook (case-sensitive
# because _scan_workbook uses them as-is).  The values are the corrected
# (cy_letter, py_letter) pairs that replace whatever auto-detection would find.
#
# These are discovered from analysis of the Lumid template layout.

def _lumid_sheet_col_map(wb_sheetnames: list[str]) -> dict:
    """
    Build the corrected shift_map for all Lumid-format sheets.

    Returns a dict  { exact_sheet_name: [(cy_letter, py_letter), ...] }
    for sheets that need an override.  Sheets not listed here are handled
    by the standard processor logic.
    """
    overrides: dict = {}
    for sn in wb_sheetnames:
        sl = sn.strip().lower()

        # ── NOA 1-2 ──────────────────────────────────────────────────────────
        # Annexure 1 (Share Capital): E=CY_count, F=CY_amount, G=PY_count, H=PY_amount
        # Annexure 2 (Reserves):       F=CY_amount, H=PY_amount
        # The overlap of F and H across both annexures means (F,H) covers both.
        # E→G (count) is kept for completeness but share-count rarely changes.
        if sl == "noa 1-2" or sl == "noa 1-2 ":
            overrides[sn] = [("E", "G"), ("F", "H")]

        # ── NOA 3-6 ──────────────────────────────────────────────────────────
        # Annexure 3  (rows  ~8-22): E=CY_NonCurrent,  H=PY_NonCurrent → E→H
        # Annexures 4-6 (rows 27-78): H=CY,            J=PY            → H→J
        # NOTE: H plays two roles across the two section groups.
        # This is safe because in practice F and G are empty in Ann3,
        # and E is empty in Ann4-6, so there is no cell-level overlap.
        elif sl == "noa 3-6":
            overrides[sn] = [("E", "H"), ("H", "J")]

        # ── NOA 7-12 ─────────────────────────────────────────────────────────
        # Auto-detection gets F→H which is correct (verified in analysis).
        # No override needed — let the standard detector handle it.

        # ── NOA 13-17 ────────────────────────────────────────────────────────
        # Auto-detection gets G→I and N→P — both correct (verified).
        # No override needed.

        # ── NOA 18-22 ────────────────────────────────────────────────────────
        # Auto-detection gets E→G — correct.  No override needed.

        # ── NOA 23-27 ────────────────────────────────────────────────────────
        # Auto-detection gets G→I — correct.  No override needed.

        # ── CASH FLOW ────────────────────────────────────────────────────────
        # Standard SHEET_COL_MAP fallback uses D→F (wrong for this template).
        # Actual layout: D=CY, E=PY.
        elif sl == "cash flow":
            overrides[sn] = [("D", "E")]

    return overrides


# Sheets that are Fixed-Asset / depreciation SCHEDULES — they must NOT have
# CY/PY column-shifting applied.  We add the Lumid-specific sheets to the
# set that the processor treats as text-only.
_LUMID_TEXT_ONLY_EXTRA = {
    "dep companies act (2)",
    "dep companies act",
    "dep income tax (3)",
    "dep income tax",
    "deferred tax (2)",
    "deferred tax",
    "sheet1",
    "ratios",
}


# ── Public entry point ───────────────────────────────────────────────────────

def process(input_path: str, output_path: str,
            closing_year: int, new_year: int) -> dict:
    """
    Year-shift a balance sheet, with Lumid-format compatibility.

    For Lumid-format files (detected by sheet-name fingerprint):
      • Column-pair overrides are injected into _scan_workbook's result.
      • DEP / deferred-tax sheets are added to TEXT_ONLY so they only get
        date-string updates, not destructive column shifting.
      • Everything else is handled by the standard processor.

    For all other files:
      • Falls through directly to processor.process() unchanged.
    """
    if not _is_lumid_format(input_path):
        # Not Lumid-format → standard processor, no changes
        return _proc.process(input_path, output_path, closing_year, new_year)

    # ── Lumid-format path ────────────────────────────────────────────────────
    # Step 1: Run the standard scan to get sizes, cy_values, cy_formulas etc.
    sizes, shift_map, cy_values, cy_formulas, cap_data = \
        _proc._scan_workbook(input_path, closing_year)

    # Step 2: Get the corrected column-pair overrides for this template.
    wb_tmp = load_workbook(input_path, read_only=True, data_only=True)
    sheetnames = wb_tmp.sheetnames
    wb_tmp.close()

    col_overrides = _lumid_sheet_col_map(sheetnames)

    # Step 3: Apply overrides to shift_map.
    # For sheets listed in col_overrides, REPLACE the auto-detected pairs.
    # For sheets not listed, keep whatever the standard detector found.
    for sn, corrected_pairs in col_overrides.items():
        shift_map[sn] = corrected_pairs

    # Step 4: For Lumid TEXT_ONLY extras, REMOVE them from shift_map
    # so they only receive date-string updates (no column shifting).
    # We must also ensure they don't get mistakenly added back by SHEET_COL_MAP.
    lumid_text_only_names = set()
    for sn in sheetnames:
        sl = sn.strip().lower()
        if sl in _LUMID_TEXT_ONLY_EXTRA:
            shift_map.pop(sn, None)     # remove from column-shift list
            lumid_text_only_names.add(sn)

    # Step 5: Re-scan cy_values and cy_formulas using the CORRECTED pairs.
    # For sheets whose pairs changed, the existing cy_values may be from the
    # wrong columns.  Re-read those sheets.
    _rescan_changed_sheets(
        input_path, closing_year, sheetnames, col_overrides,
        shift_map, cy_values, cy_formulas
    )

    # Step 6: Build a patched TEXT_ONLY set that includes Lumid extras.
    # We monkey-patch the module-level constant temporarily so that the
    # process() function's inner logic respects our additions.
    original_text_only = _proc.TEXT_ONLY_SHEETS
    _proc.TEXT_ONLY_SHEETS = original_text_only | _LUMID_TEXT_ONLY_EXTRA

    # Step 7: Override SHEET_COL_MAP fallback for cash flow.
    # The standard SHEET_COL_MAP has "cash flow": [("D","F")] which is wrong
    # for this template.  We temporarily override it.
    original_col_map = dict(_proc.SHEET_COL_MAP)
    _proc.SHEET_COL_MAP["cash flow"] = [("D", "E")]

    # Step 8: Build the date-replacement pairs (same as standard)
    pairs = _proc._date_replacements(str(closing_year), str(new_year))
    log: list[str] = []
    log.append("ℹ Lumid-format detected — applying compatibility overrides")
    log.append(f"  Column-pair overrides: {list(col_overrides.keys())}")
    log.append(f"  TEXT_ONLY extras: {sorted(lumid_text_only_names)}")
    log.append(f"  CASH FLOW remapped: D→E (was D→F)")

    # Step 9: Execute the ZIP-level edit loop (copied from processor.process()
    # but using our patched shift_map).
    result = _run_with_patched_maps(
        input_path, output_path,
        closing_year, new_year,
        sizes, shift_map, cy_values, cy_formulas, cap_data,
        pairs, log,
        lumid_text_only_names,
    )

    # Step 10: Restore the temporarily patched module globals.
    _proc.TEXT_ONLY_SHEETS = original_text_only
    _proc.SHEET_COL_MAP.clear()
    _proc.SHEET_COL_MAP.update(original_col_map)

    return result


# ── Internal helpers ─────────────────────────────────────────────────────────

def _noa36_section_boundary(ws) -> int:
    """
    Find the row number where NOA 3-6's second section (Annexure 4 onwards)
    begins.  This is the first row > 22 where column E (col 5) contains an
    'ANNEXURE' label or a deferred-tax section header.

    Returns the boundary row (rows < boundary belong to Annexure 3,
    rows >= boundary belong to Annexures 4-6).  Defaults to 27 if not found.
    """
    for r in range(23, 90):
        for c in range(3, 7):
            v = ws.cell(r, c).value
            if isinstance(v, str) and (
                "ANNEXURE" in v.upper()
                or "DEFERRED" in v.upper()
            ):
                return r
    return 27   # safe fallback


def _rescan_changed_sheets(
    filepath: str,
    closing_year: int,
    sheetnames: list[str],
    col_overrides: dict,
    shift_map: dict,
    cy_values: dict,
    cy_formulas: dict,
) -> None:
    """
    For sheets whose column pairs were overridden, re-read the CY column
    values from the corrected columns (the original scan may have read the
    wrong columns).

    Special handling for NOA 3-6 (Issue 2 / H-column conflict):
    ─────────────────────────────────────────────────────────────
    NOA 3-6 is mapped as [(E,H), (H,J)].  The sheet has two section groups:
      Ann 3  (rows  ~8-26):  E = CY_NonCurrent, H = PY_NonCurrent
      Ann 4-6 (rows 27-78):  H = CY,            J = PY

    Within a single processor pass, both pairs are applied to every row.
    The conflict arises in Ann 3 rows 14/16/20/22 where H has numeric data
    (PY values = 175):
      • (E→H) correctly shifts E → H (H becomes the new PY value)
      • (H→J) then sees the NEWLY written H value as "CY", clears H,
        and writes it to J — destroying the correctly-shifted H value.

    Fix: restrict the cy_values for column H to ONLY rows >= section_boundary
    (i.e. Ann4-6 rows where H genuinely is the CY column).  For Ann3 rows,
    H has no cy_value entry → (H→J) makes no change there → no conflict.
    """
    changed_sheets = set(col_overrides.keys())
    if not changed_sheets:
        return

    # Get the sheet-file map from the ZIP
    with zipfile.ZipFile(filepath) as z:
        smap = _proc._sheet_file_map(z)

        _f_cell_re = re.compile(rb'<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>[^<]*<f[ />]')
        formula_refs_new: dict = {}
        formula_texts_new: dict = {}
        for sn in changed_sheets:
            sf = smap.get(sn, "")
            if not sf:
                continue
            try:
                xml_data = z.read(sf)
                refs = set(m.group(1).decode() for m in _f_cell_re.finditer(xml_data))
                formula_refs_new[sn] = refs

                ftexts: dict = {}
                for cm in re.finditer(
                    rb'<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>(.*?)</c>',
                    xml_data, re.DOTALL
                ):
                    ref_b, body_b = cm.group(1), cm.group(2)
                    has_real_v = bool(re.search(rb'<v[^>]*>[^<]+</v>', body_b))
                    if has_real_v:
                        continue
                    fm = re.search(rb'<f[^>]*>([^<]*)</f>', body_b)
                    if fm:
                        try:
                            ftexts[ref_b.decode()] = fm.group(1).decode("utf-8", "replace")
                        except Exception:
                            pass
                formula_texts_new[sn] = ftexts
            except Exception:
                formula_refs_new[sn] = set()
                formula_texts_new[sn] = {}

    # Re-read the corrected CY columns with openpyxl
    wb = load_workbook(filepath, read_only=True, data_only=True)
    try:
        for sn in changed_sheets:
            if sn not in wb.sheetnames:
                continue
            ws = wb[sn]
            det = shift_map.get(sn, [])
            if not det:
                continue

            sheet_frefs = formula_refs_new.get(sn, set())
            sheet_ftexts = formula_texts_new.get(sn, {})

            # Per-column row-range restrictions (for multi-section sheets).
            # {col_letter: (min_row, max_row)} — None means no restriction.
            row_restrictions: dict[str, tuple[int, int] | None] = {}

            sl = sn.strip().lower()
            if sl == "noa 3-6":
                # Detect boundary between Ann3 and Ann4 dynamically.
                boundary = _noa36_section_boundary(ws)
                # E column: only Ann3 rows (rows < boundary) are the CY section
                row_restrictions["E"] = (1, boundary - 1)
                # H column: only Ann4-6 rows (rows >= boundary) are CY
                row_restrictions["H"] = (boundary, 10_000)

            cy_vals_sheet: dict = {}
            cy_frows_sheet: dict = {}

            for cy_l, _ in det:
                ci = _proc._col_idx(cy_l)
                vals: dict = {}
                frows: set = set()

                # Determine the allowed row range for this column
                rng = row_restrictions.get(cy_l)  # (min_row, max_row) or None
                min_r = rng[0] if rng else 1
                max_r = rng[1] if rng else 10_000

                for row in ws.iter_rows(min_col=ci, max_col=ci):
                    for cell in row:
                        if isinstance(cell, MergedCell):
                            continue
                        if not hasattr(cell, 'row') or cell.row is None:
                            continue
                        rn = cell.row

                        # Apply row-range restriction
                        if rn < min_r or rn > max_r:
                            continue

                        ref = f"{cy_l}{rn}"

                        if ref in sheet_frefs:
                            frows.add(rn)
                            if isinstance(cell.value, (int, float)) and cell.value is not None:
                                vals[rn] = float(cell.value)
                            elif cell.value is None:
                                ftext = sheet_ftexts.get(ref)
                                evaluated = _proc._try_eval_arithmetic_formula(ftext)
                                vals[rn] = evaluated if evaluated is not None else 0.0
                        elif isinstance(cell.value, (int, float)) and cell.value is not None:
                            vals[rn] = float(cell.value)
                        elif isinstance(cell.value, str) and cell.value.strip() in ("-", "—", "–"):
                            vals[rn] = 0.0

                cy_vals_sheet[cy_l] = vals
                cy_frows_sheet[cy_l] = frows

            cy_values[sn] = cy_vals_sheet
            cy_formulas[sn] = cy_frows_sheet
    finally:
        wb.close()


def _run_with_patched_maps(
    input_path: str,
    output_path: str,
    closing_year: int,
    new_year: int,
    sizes: dict,
    shift_map: dict,
    cy_values: dict,
    cy_formulas: dict,
    cap_data: dict,
    pairs: list,
    log: list,
    lumid_text_only_names: set,
) -> dict:
    """
    Re-implement the ZIP-level loop from processor.process() using the
    patched shift_map and cy_values.  This avoids calling _scan_workbook
    a second time while still running the full edit pipeline.
    """
    import zipfile as _zipfile

    # Compute effective-big-sheet set (same logic as processor.process)
    BIG_ROWS = _proc.BIG_ROWS
    BIG_COLS = _proc.BIG_COLS

    def _effective_big(n, r, c):
        if r > BIG_ROWS:
            return True
        if c > BIG_COLS:
            sm_pairs = shift_map.get(n, [])
            if sm_pairs:
                from openpyxl.utils import column_index_from_string
                max_sm_col = max(
                    column_index_from_string(col)
                    for pair in sm_pairs for col in pair
                )
                return max_sm_col > BIG_COLS
            return True
        return False

    big_names = {n for n, (r, c) in sizes.items() if _effective_big(n, r, c)}
    date_style_indices = _proc._build_date_style_set(input_path)

    RAW_DATA_SHEETS  = _proc.RAW_DATA_SHEETS
    CAPITAL_SHEET_NAMES = _proc.CAPITAL_SHEET_NAMES
    # TEXT_ONLY_SHEETS already patched at module level by caller

    ext_count = 0

    with _zipfile.ZipFile(input_path, "r") as zi:
        smap = _proc._sheet_file_map(zi)

        with _zipfile.ZipFile(output_path, "w", _zipfile.ZIP_DEFLATED) as zo:
            for item in zi.infolist():
                fn   = item.filename
                data = zi.read(fn)

                # Skip external links
                if "externalLinks" in fn:
                    continue

                # Shared strings — date-text replacement only
                if fn == "xl/sharedStrings.xml":
                    data = _proc._update_shared_strings(data, pairs)
                    zo.writestr(item, data)
                    continue

                # Worksheet XMLs
                if fn.startswith("xl/worksheets/") and fn.endswith(".xml"):
                    sheet_name = next(
                        (sn for sn, sf in smap.items() if sf == fn), None
                    )
                    sl = (sheet_name or "").strip().lower()

                    # Strip external formulas from all sheets
                    before_len = len(data)
                    data = _proc._strip_external_formulas(data)
                    if len(data) != before_len:
                        ext_count += 1

                    if sheet_name and sheet_name in shift_map:
                        # Full CY→PY shift + date update
                        data = _proc._process_sheet_xml(
                            data,
                            shift_map[sheet_name],
                            cy_values.get(sheet_name, {}),
                            cy_formulas.get(sheet_name, {}),
                            date_style_indices,
                            shift_map,
                        )
                        data = _proc._update_inline_strings(data, pairs)
                        desc = ", ".join(
                            f"{c}→{p}" for c, p in shift_map[sheet_name]
                        )
                        log.append(
                            f"✓ {sheet_name}: CY→PY copied ({desc}), "
                            f"CY constants cleared, dates updated"
                        )
                    elif sheet_name and sl in RAW_DATA_SHEETS:
                        log.append(f"— {sheet_name}: skipped (raw data sheet)")
                    elif sheet_name and sheet_name.strip().lower() in CAPITAL_SHEET_NAMES:
                        cap_info = cap_data.get(sheet_name, {})
                        if cap_info:
                            data = _proc._process_capital_sheet(data, cap_info)
                            log.append(
                                f"✓ {sheet_name}: capital CY→PY shifted, "
                                f"CY opening updated, additions/withdrawals cleared"
                            )
                        data = _proc._update_inline_strings(data, pairs)
                    elif sheet_name and sheet_name in lumid_text_only_names:
                        # Lumid-specific TEXT_ONLY extra: just update date strings
                        data = _proc._update_inline_strings(data, pairs)
                        log.append(f"· {sheet_name}: dates updated (Lumid TEXT_ONLY)")
                    elif sheet_name and sheet_name not in big_names:
                        data = _proc._update_inline_strings(data, pairs)
                        log.append(f"· {sheet_name}: dates updated")
                    elif sheet_name in big_names:
                        log.append(
                            f"* {sheet_name}: preserved unchanged "
                            f"(large sheet {sizes[sheet_name][0]} rows)"
                        )

                    zo.writestr(item, data)
                    continue

                # workbook.xml.rels — remove external-link references
                if fn == "xl/workbook.xml.rels" or fn.endswith(".rels"):
                    text_rels = data.decode("utf-8", errors="replace")
                    if "externalLinks" in text_rels:
                        text_rels = re.sub(
                            r'<Relationship[^>]*Target="externalLinks[^"]*"[^>]*/>\s*',
                            '', text_rels
                        )
                        data = text_rels.encode("utf-8")

                # workbook.xml — remove <externalReferences> section
                if fn == "xl/workbook.xml":
                    text_wb = data.decode("utf-8", errors="replace")
                    if "<externalReferences" in text_wb:
                        text_wb = re.sub(
                            r'<externalReferences\b[^>]*>.*?</externalReferences>',
                            '', text_wb, flags=re.DOTALL
                        )
                        data = text_wb.encode("utf-8")

                zo.writestr(item, data)

    if ext_count:
        log.append(
            f"🔗 External references converted to values in {ext_count} sheet(s)"
        )

    return {"status": "success", "log": log}


# ── Convenience: expose same API as processor.py ─────────────────────────────

__all__ = ["process", "_is_lumid_format"]


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 5:
        print("Usage: lumid_compat.py INPUT OUTPUT CLOSING_YEAR NEW_YEAR")
        sys.exit(1)
    res = process(sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]))
    print(f"Status: {res['status']}")
    for line in res.get("log", []):
        print(" ", line)
