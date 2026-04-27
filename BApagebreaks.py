"""
BApagebreaks.py
================
Post-processing for the BA Rate Pages workbook.

PIPELINE
  1. openpyxl: truncate sheet names > 31 chars, move Index to position 0,
     apply per-rule page breaks and print settings (in-memory; no COM).
  2. XML sanitize pass on the saved .xlsx zip: removes the few openpyxl
     output quirks that make Excel show "Open and Repair" on file open.

Pure Python, no Excel.exe. Typical run: ~1-2 seconds end-to-end.

----------------------------------------------------------------------------
HOW TO ADD A PAGE-BREAK RULE FOR A NEW SHEET
----------------------------------------------------------------------------

Step 1 - Write a handler function. Signature: (ws, dest_filename) -> None
         where `ws` is an openpyxl Worksheet and `dest_filename` is the
         absolute file path (use it for state/file-specific branching).

    def _handle_rule_999(ws, dest_filename):
        ws.print_area = f"A1:H{ws.max_row}"
        disable_fit_to_page(ws)
        add_break_after(ws, 37)

Step 2 - Register it in SHEET_RULES below (one line):

    ("Rule 999", _handle_rule_999),

Order matters: more-specific prefixes BEFORE less-specific ones
(e.g., "Rule 239 C" must come before "Rule 239 ").

----------------------------------------------------------------------------
HELPERS YOU CAN USE INSIDE A HANDLER
----------------------------------------------------------------------------

  fit_single_page(ws)         entire sheet onto one printed page
  fit_width_only(ws)          fit width to 1 page; height grows with content
  disable_fit_to_page(ws)     turn off fit-to-page; manual breaks rule
  add_break_after(ws, row)    add a horizontal page break AFTER the given row

OPENPYXL API CHEATSHEET
  ws.page_setup.orientation     = "portrait" or "landscape"
  ws.page_setup.fitToWidth      = 1
  ws.page_setup.fitToHeight     = 1   (or 0 for unlimited tall)
  ws.print_area                 = "A1:H50"
  ws.print_title_rows           = "1:3"   (or None to clear)
  ws.print_options.horizontalCentered = False
  ws.print_options.verticalCentered   = False
  ws.page_margins.top           = 1.00     (inches)
  ws.row_breaks.append(Break(id=N))         break ABOVE row N
"""

import os
import re
import shutil
import zipfile
from io import BytesIO

import openpyxl
from openpyxl.worksheet.pagebreak import Break


# ============================================================================
#  HELPERS  -  use these inside rule handlers
# ============================================================================

def fit_single_page(ws):
    """Fit the entire content of the sheet onto a single printed page."""
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


def fit_width_only(ws):
    """Fit width to 1 page; height grows as needed (manual breaks honored)."""
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def disable_fit_to_page(ws):
    """Turn off fit-to-page entirely; manual breaks are honored as-is."""
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.fitToWidth = 0
    ws.page_setup.fitToHeight = 0


def add_break_after(ws, row):
    """Add a horizontal page break AFTER the given row (1-indexed)."""
    ws.row_breaks.append(Break(id=row))


# ============================================================================
#  RULE HANDLERS
#  Signature: (ws, dest_filename) -> None
# ============================================================================

def _handle_index(ws, dest_filename):
    # NOTE: print_title_rows = None (NOT "0:0") so we don't write an invalid
    # definedName like Index!$0:$0 - that is what triggers Open-and-Repair.
    ws.print_title_rows = None
    ws.print_area = f"A1:J{ws.max_row}"
    fit_width_only(ws)


def _handle_rule_222b(ws, dest_filename):
    disable_fit_to_page(ws)
    add_break_after(ws, 25)
    add_break_after(ws, 49)


def _handle_rule_222ttt(ws, dest_filename):
    fit_single_page(ws)
    ws.print_title_rows = "1:3"


def _handle_rule_223b5(ws, dest_filename):
    ws.page_setup.orientation = "landscape"


def _handle_rule_223c(ws, dest_filename):
    fit_single_page(ws)


def _handle_rule_225_zone(ws, dest_filename):
    ws.print_area = f"A1:M{ws.max_row}"
    ws.print_options.horizontalCentered = False
    ws.print_options.verticalCentered = False
    disable_fit_to_page(ws)
    for row in range(52, ws.max_row, 51):
        add_break_after(ws, row)


def _handle_rule_225c3(ws, dest_filename):
    fit_single_page(ws)


def _handle_rule_232ppt(ws, dest_filename):
    fit_single_page(ws)
    ws.print_title_rows = "1:3"


def _handle_rule_239_general(ws, dest_filename):
    fit_single_page(ws)
    ws.print_title_rows = "1:3"


def _handle_rule_239c(ws, dest_filename):
    fit_single_page(ws)
    ws.page_margins.top = 1.00


def _handle_rule_240(ws, dest_filename):
    fit_single_page(ws)
    ws.print_options.verticalCentered = True
    ws.print_title_rows = "1:3"
    ws.print_area = f"A1:M{ws.max_row}"
    ws.page_margins.top = 1.00


def _handle_rule_255(ws, dest_filename):
    ws.print_area = f"A1:H{ws.max_row}"
    ws.print_options.horizontalCentered = False
    ws.print_options.verticalCentered = False
    disable_fit_to_page(ws)
    add_break_after(ws, 37)


def _handle_rule_275(ws, dest_filename):
    if ws["A10"].value == "275.B.1.(b). Short Term - Autos Leased by the Hour, Day, or Week":
        ws.print_title_rows = "1:1"
        fit_single_page(ws)


def _handle_rule_283(ws, dest_filename):
    ws.print_area = f"A1:P{ws.max_row}"
    targets = {
        "283.B Limited Specified Causes of Loss",
        "283.B Comprehensive",
        "283.B Blanket Collision",
    }
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value)
        if cell_value in targets and row > 3:
            # break BEFORE this row (= after row-1)
            ws.row_breaks.append(Break(id=row - 1))
    fit_width_only(ws)


def _handle_rule_289(ws, dest_filename):
    ws.print_area = f"A1:H{ws.max_row}"
    disable_fit_to_page(ws)
    add_break_after(ws, 37)


def _handle_rule_297(ws, dest_filename):
    ws.print_area = f"A1:P{ws.max_row}"
    disable_fit_to_page(ws)
    occurrence_count = 0
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value)
        if cell_value.startswith("Single") or cell_value.startswith("Uninsured"):
            occurrence_count += 1
        if (occurrence_count % 3 == 0) and (occurrence_count != 0):
            occurrence_count += 1
            ws.row_breaks.append(Break(id=row - 1))


def _handle_rule_298(ws, dest_filename):
    ws.print_area = f"A1:K{ws.max_row}"
    disable_fit_to_page(ws)
    occurrence_count = 0
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value)
        if cell_value.startswith("298"):
            occurrence_count += 1
        if occurrence_count == 4:
            occurrence_count += 1
            ws.row_breaks.append(Break(id=row - 1))
        if occurrence_count == 8:
            break


_VA_VEHICLE_TYPES = {
    "Extra Heavy Truck-Tractor", "Extra-Heavy Truck", "Heavy Truck",
    "Heavy Truck-Tractor", "Light Truck", "Medium Truck",
    "Private Passenger Types", "Semitrailer",
    "Service or Utility Trailer", "Trailer",
}


def _handle_rule_301ab(ws, dest_filename):
    if ws["B4"].value in _VA_VEHICLE_TYPES:
        return
    fit_width_only(ws)
    if ws.title.startswith("Rule 301.B"):
        ws.print_area = f"A1:T{ws.max_row}"
    for row in range(46, ws.max_row, 45):
        add_break_after(ws, row)
    ws.print_options.horizontalCentered = False
    ws.print_options.verticalCentered = False
    ws.page_setup.orientation = "landscape"
    ws.page_margins.top = 1.00


def _handle_rule_301cd(ws, dest_filename):
    if ws["B4"].value not in _VA_VEHICLE_TYPES:
        return
    if "FL" not in dest_filename:
        ws.page_margins.top = 1.00
    fit_single_page(ws)


def _handle_rule_306(ws, dest_filename):
    fit_width_only(ws)
    ws.print_title_rows = "1:4"


def _handle_rule_315(ws, dest_filename):
    fit_width_only(ws)
    add_break_after(ws, 23)


def _handle_rule_r1(ws, dest_filename):
    ws.print_area = f"A1:M{ws.max_row}"
    disable_fit_to_page(ws)
    occurrence_count = 0
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value)
        if cell_value.startswith("R1"):
            occurrence_count += 1
        if occurrence_count in (3, 6):
            occurrence_count += 1
            ws.row_breaks.append(Break(id=row - 1))


# ============================================================================
#  RULE REGISTRY
#  Order matters: more-specific prefixes BEFORE less-specific ones.
#  To add a new rule: write the handler above, then add one line here.
# ============================================================================

SHEET_RULES = [
    ("Index",        _handle_index),

    ("Rule 222 B",   _handle_rule_222b),
    ("Rule 222 TTT", _handle_rule_222ttt),

    ("Rule 223 B.5", _handle_rule_223b5),
    ("Rule 223 C",   _handle_rule_223c),

    ("Rule 225 Zone", _handle_rule_225_zone),
    ("Rule 225.C.3", _handle_rule_225c3),

    ("Rule 232 PPT", _handle_rule_232ppt),

    ("Rule 239 C",   _handle_rule_239c),       # specific BEFORE generic
    ("Rule 239 ",    _handle_rule_239_general),

    ("Rule 240 ",    _handle_rule_240),

    ("Rule 255",     _handle_rule_255),

    ("Rule 275",     _handle_rule_275),

    ("Rule 283",     _handle_rule_283),
    ("Rule 289",     _handle_rule_289),
    ("Rule 297",     _handle_rule_297),
    ("Rule 298",     _handle_rule_298),

    ("Rule 301.C",   _handle_rule_301cd),
    ("Rule 301.D",   _handle_rule_301cd),
    ("Rule 301.A",   _handle_rule_301ab),
    ("Rule 301.B",   _handle_rule_301ab),

    ("Rule 306",     _handle_rule_306),
    ("Rule 315",     _handle_rule_315),

    ("Rule R1",      _handle_rule_r1),
]


def _apply_matching_rule(sheet_name, ws, dest_filename):
    """Walk SHEET_RULES and run the first handler whose prefix matches."""
    for prefix, handler in SHEET_RULES:
        if sheet_name.startswith(prefix):
            handler(ws, dest_filename)
            return True
    return False


# ============================================================================
#  PUBLIC ENTRY POINT
# ============================================================================

def process_pagebreaks(dest_filename1, dest_filename2):
    """
    Apply page breaks / print settings to dest_filename1.

    dest_filename2 is accepted for backward compatibility (was a PDF path).
    """
    print(f"[BApagebreaks] Processing: {dest_filename1}")
    dest_filename1 = os.path.normpath(os.path.abspath(dest_filename1))

    workbook = openpyxl.load_workbook(dest_filename1)

    # Truncate sheet names exceeding Excel's 31-character limit
    for original_name in list(workbook.sheetnames):
        if len(original_name) > 31:
            workbook[original_name].title = original_name[:31]

    # Apply defaults + rules to every sheet
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        ws.print_title_rows = "1:1"
        fit_single_page(ws)
        _apply_matching_rule(sheet_name, ws, dest_filename1)

    # Index goes to position 0, visible, active on open
    if "Index" in workbook.sheetnames:
        ws_index = workbook["Index"]
        ws_index.sheet_state = "visible"
        if workbook.sheetnames.index("Index") != 0:
            workbook._sheets.remove(ws_index)
            workbook._sheets.insert(0, ws_index)
        workbook.active = 0

    workbook.save(dest_filename1)
    workbook.close()

    # Sanitize the saved zip to prevent Excel's "Open and Repair" popup
    _sanitize_xlsx(dest_filename1)
    print("[BApagebreaks] Done.")


# ============================================================================
#  XML SANITIZE PASS
# ============================================================================
#
#  Why: openpyxl's xlsx output sometimes contains tiny XML quirks that make
#  Excel show "We found a problem with some content" on open. The safest
#  fixes Excel's repair would apply, applied here directly:
#
#    * Drop <definedName> entries whose ref contains "$0" - rows are 1-indexed
#      in Excel; "$0" never matches anything and Excel flags it as invalid.
#
#    * Drop <definedName> entries whose ref is empty.
#
#    * Drop empty <definedNames/> wrappers if they end up with no children.
#
#  Operates directly on the .xlsx zip. ~0.5 seconds even for large workbooks.
#  No Excel process needed.
# ============================================================================

# Match an entire <definedName ...>...</definedName> element whose body
# contains "$0" anywhere - the only practical way openpyxl produces an
# invalid range like Sheet1!$0:$0.
_BAD_DEFINED_NAME_RE = re.compile(
    rb"<definedName\b[^>]*>[^<]*\$0[^<]*</definedName>"
)
# Also catch self-closing <definedName ... ref="..."/> variants and
# definedName elements with empty body.
_EMPTY_DEFINED_NAME_RE = re.compile(
    rb"<definedName\b[^>]*>\s*</definedName>"
)
# If the wrapper ends up empty, drop the whole <definedNames/> element.
_EMPTY_DEFINED_NAMES_WRAPPER_RE = re.compile(
    rb"<definedNames\s*/?>\s*</definedNames>|<definedNames\s*/>"
)


def _sanitize_xlsx(filename):
    """
    Open the xlsx as a zip, rewrite xl/workbook.xml in-place to drop any
    <definedName> elements with invalid "$0" or empty refs, and save the
    zip back. Pure Python, no Excel.
    """
    tmp = filename + ".sanitize.tmp"
    changed = False

    with zipfile.ZipFile(filename, "r") as zin:
        infos = zin.infolist()
        contents = {info.filename: zin.read(info.filename) for info in infos}

    if "xl/workbook.xml" in contents:
        original = contents["xl/workbook.xml"]
        cleaned = _BAD_DEFINED_NAME_RE.sub(b"", original)
        cleaned = _EMPTY_DEFINED_NAME_RE.sub(b"", cleaned)
        cleaned = _EMPTY_DEFINED_NAMES_WRAPPER_RE.sub(b"", cleaned)
        if cleaned != original:
            contents["xl/workbook.xml"] = cleaned
            changed = True

    if not changed:
        # Nothing to fix - keep original file as-is
        return

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, contents[info.filename])

    os.replace(tmp, filename)


# Example usage:
# process_pagebreaks(r"C:\path\to\workbook.xlsx", "ignored.pdf")
