"""
BApagebreaks.py
===============
Post-processing step: apply page breaks / print settings to the generated
Excel workbook via Excel COM automation, then leave it open for the user.

WHAT CHANGED vs previous version
─────────────────────────────────
1.  ELIMINATED the openpyxl load → modify → save cycle.
    That cycle corrupted the XML because openpyxl incorrectly merges the
    fitToPage / pageSetUpPr attributes that ExcelSettingsBA already wrote.
    All page-break and print-setup changes are now applied directly via
    Excel COM, which always writes well-formed XML.

2.  ELIMINATED the CorruptLoad repair step:
    Before: openpyxl resave → corrupt xlsx → COM CorruptLoad → _repaired.xlsx
    After:  COM opens the clean file directly → applies rules → saves clean

3.  SHEET_RULES registry is preserved — adding a new rule is still
    one function + one line in SHEET_RULES.

4.  IDENTICAL OUTPUT — every page_setup, row break, print_area, and
    page_margins value is the same; only the implementation mechanism changed.

HOW TO ADD A NEW RULE
──────────────────────
  1.  Write a handler function:
          def _handle_rule_XXX(ws_com, xl_app, dest_filename):
              ws_com.PageSetup.FitToPagesWide = 1
              # ... etc.

  2.  Register it in SHEET_RULES (one line):
          ("Rule XXX", _handle_rule_XXX),

  That is ALL.

COM API TRANSLATION REFERENCE
──────────────────────────────
  openpyxl                              COM equivalent
  ─────────────────────────────────     ──────────────────────────────────────
  ws.page_setup.fitToWidth  = 1        ws_com.PageSetup.Zoom = False
  ws.page_setup.fitToHeight = 1          ws_com.PageSetup.FitToPagesWide = 1
                                          ws_com.PageSetup.FitToPagesTall = 1
  ws.page_setup.fitToHeight = False     ws_com.PageSetup.FitToPagesTall = False
  ws.page_setup.orientation="landscape" ws_com.PageSetup.Orientation = 2
  ws.print_title_rows = "1:3"          ws_com.PageSetup.PrintTitleRows = "$1:$3"
  ws.print_area = "A1:Jn"             ws_com.PageSetup.PrintArea = "$A$1:$J$n"
  ws.page_margins.top = 1.00          ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)
  ws.print_options.horizontalCentered  ws_com.PageSetup.CenterHorizontally
  ws.print_options.verticalCentered    ws_com.PageSetup.CenterVertically
  sheet.row_breaks.append(Break(N))    ws_com.HPageBreaks.Add(ws_com.Rows(N + 1))
  Break(row - 1)  [break before row]  ws_com.HPageBreaks.Add(ws_com.Rows(row))
  workbook[title]["A10"].value         ws_com.Range("A10").Value
  ws.max_row                           ws_com.UsedRange.Rows.Count
  ws.title                             ws_com.Name
"""

import gc
import os
import time


try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None   # allows importing this module on non-Windows without crash


# ===========================================================================
#  COM CONSTANTS
# ===========================================================================

_XL_LANDSCAPE = 2   # xlLandscape


# ===========================================================================
#  SMART PAGE-BREAK HELPERS
# ===========================================================================

def _clear_page_breaks(ws_com):
    """Remove all existing horizontal page breaks (e.g. written by openpyxl)."""
    count = ws_com.HPageBreaks.Count
    for _ in range(count, 0, -1):
        ws_com.HPageBreaks(1).Delete()


def _page_height_after_block(block_h, page_h):
    """Return how much of the last page is occupied after placing a block."""
    rem = block_h % page_h
    return rem if rem > 0 else (page_h if block_h >= page_h else block_h)


def _smart_page_breaks(ws_com, xl_app, sheet_name=""):
    """
    Insert horizontal page breaks that keep each logical table block on one
    page wherever possible.

    For most sheets, blocks are delimited by blank rows in column A.
    Rule R1 is a special case: its sections start at rows whose column-A
    value has an "R1" prefix (no blank separators exist between sections).

    Algorithm — greedy page packing:
      1.  Identify raw block-start rows using the appropriate detection method.
      2.  Merge "tiny" lead blocks (≤ SUBTITLE_ROWS rows) with the next block
          so subtitle rows are never stranded alone at the bottom of a page.
          generateWorksheetTablesX writes:  subtitle → blank → headers → data
          Blank-row detection fires twice per sub-table (at the subtitle row
          and again at the headers row after the within-table blank), creating
          a 2-row "subtitle+blank" stub that must travel with its data block.
      3.  Measure each merged group's total height in points.
      4.  If the next group fits in the remaining page space, keep it there;
          otherwise insert a break before it and open a new page.
      A group larger than one full page is placed as-is — Excel will break
      it internally rather than us splitting it artificially.
    """
    # Only skip when Excel's own fit-to-N-tall scaling is active AND zoom is
    # disabled.  When Zoom=False;FitToPagesWide=1 is set without explicitly
    # setting FitToPagesTall=False, Excel silently defaults FitToPagesTall=1
    # even though the intent was "fit width only".  We guard against that by
    # requiring BOTH conditions before bailing out.
    try:
        zoom_off = not ws_com.PageSetup.Zoom   # True when Zoom is disabled (fit-to-pages mode)
        ftt = int(ws_com.PageSetup.FitToPagesTall)
        if zoom_off and ftt >= 1:
            return
    except (TypeError, ValueError, Exception):
        pass

    max_row = ws_com.UsedRange.Rows.Count
    if max_row < 2:
        return

    # Available vertical space per page (points)
    is_landscape = (ws_com.PageSetup.Orientation == _XL_LANDSCAPE)
    paper_h_pts  = xl_app.InchesToPoints(8.5 if is_landscape else 11.0)
    avail_pts    = (paper_h_pts
                   - ws_com.PageSetup.TopMargin
                   - ws_com.PageSetup.BottomMargin)

    # Row heights (one COM call per row)
    row_h = [ws_com.Rows(r).RowHeight for r in range(1, max_row + 1)]

    # Column A values for boundary detection
    col_a = ws_com.Range(f"A1:A{max_row}").Value  # tuple of (value,) tuples

    # ── Step 1: raw block-start rows ──────────────────────────────────────────
    if sheet_name.startswith("Rule R1"):
        # R1 has no blank-row separators; each row whose A value starts with
        # "R1" is treated as the opening of a new section.
        block_starts = [1]
        for r in range(2, max_row + 1):
            v = col_a[r - 1][0]
            if v is not None and str(v).startswith("R1"):
                block_starts.append(r)
        block_starts.append(max_row + 1)  # sentinel

        # R1 blocks carry no subtitle stub; use them as groups directly
        groups = [
            (block_starts[i], block_starts[i + 1] - 1)
            for i in range(len(block_starts) - 1)
        ]
    else:
        # Generic: a block starts at row 1 and at every non-blank row that
        # immediately follows a blank row.
        def blank(r):
            v = col_a[r - 1][0]
            return v is None or str(v).strip() == ""

        block_starts = [1]
        for r in range(2, max_row + 1):
            if blank(r - 1) and not blank(r):
                block_starts.append(r)
        block_starts.append(max_row + 1)  # sentinel

        # ── Step 2: merge subtitle lead-blocks with the next block ────────────
        # Any raw block whose row span is ≤ SUBTITLE_ROWS is a subtitle stub
        # (subtitle + within-table blank, possibly plus the sheet title row).
        # Merge it forward so the subtitle and its data are one unit for packing.
        SUBTITLE_ROWS = 3
        groups = []
        i = 0
        while i < len(block_starts) - 1:
            g_start   = block_starts[i]
            g_end_idx = i + 1   # exclusive index into block_starts (points past group end)
            # Extend while the accumulated lead is still a stub
            while (g_end_idx < len(block_starts) - 1 and
                   block_starts[g_end_idx] - g_start <= SUBTITLE_ROWS):
                g_end_idx += 1
            g_end = block_starts[g_end_idx] - 1
            groups.append((g_start, g_end))
            i = g_end_idx

    # ── Step 3: greedy page fill ──────────────────────────────────────────────
    page_used = 0.0
    for (g_start, g_end) in groups:
        g_h = sum(row_h[r - 1] for r in range(g_start, g_end + 1))

        if page_used == 0.0:
            # First group on this page — always accept (even if it overflows)
            page_used = _page_height_after_block(g_h, avail_pts)
        elif page_used + g_h > avail_pts:
            # Group doesn't fit → start a new page before it
            ws_com.HPageBreaks.Add(ws_com.Rows(g_start))
            page_used = _page_height_after_block(g_h, avail_pts)
        else:
            page_used += g_h


# ===========================================================================
#  RULE HANDLER FUNCTIONS
#  Each handles one sheet-name prefix.
#
#  Signature: (ws_com, xl_app, dest_filename) → None
#    ws_com        — COM Worksheet object
#    xl_app        — COM Excel.Application  (needed for InchesToPoints)
#    dest_filename — absolute path to the .xlsx file (for filename-dependent rules)
# ===========================================================================

def _handle_index(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$J${max_row}"
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = False   # False = automatic height (unlimited pages)



def _handle_rule_222ttt(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = "$1:$3"


def _handle_rule_223b5(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Orientation = _XL_LANDSCAPE


def _handle_rule_223c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_225zone(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$M${max_row}"
    ws_com.PageSetup.CenterHorizontally = False
    ws_com.PageSetup.CenterVertically = False


def _handle_rule_225c3(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_232ppt(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = "$1:$3"


def _handle_rule_239_not_c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = "$1:$3"


def _handle_rule_239c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_240(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.CenterVertically = True
    ws_com.PageSetup.PrintTitleRows = "$1:$3"
    ws_com.PageSetup.PrintArea = f"$A$1:$M${max_row}"
    ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_255(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$H${max_row}"
    ws_com.PageSetup.CenterHorizontally = False
    ws_com.PageSetup.CenterVertically = False


def _handle_rule_275(ws_com, xl_app, dest_filename):
    if ws_com.Range("A10").Value == (
        "275.B.1.(b). Short Term - Autos Leased by the Hour, Day, or Week"
    ):
        ws_com.PageSetup.PrintTitleRows = "$1:$1"
        ws_com.PageSetup.Zoom = False
        ws_com.PageSetup.FitToPagesWide = 1
        ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_283(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$P${max_row}"
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = False  # explicit: fit width only, not height


def _handle_rule_289(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$H${max_row}"


def _handle_rule_297(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$P${max_row}"


def _handle_rule_298(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$K${max_row}"


def _handle_rule_301ab(ws_com, xl_app, dest_filename):
    """Rule 301.A and 301.B — VA has special vehicle-type exceptions."""
    _va_vehicle_types = {
        "Extra Heavy Truck-Tractor", "Extra-Heavy Truck", "Heavy Truck",
        "Heavy Truck-Tractor", "Light Truck", "Medium Truck",
        "Private Passenger Types", "Semitrailer",
        "Service or Utility Trailer", "Trailer",
    }
    if ws_com.Range("B4").Value in _va_vehicle_types:
        return   # VA vehicle-type pages: no custom breaks

    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = False  # explicit: fit width only, not height
    if ws_com.Name.startswith("Rule 301.B"):
        ws_com.PageSetup.PrintArea = f"$A$1:$T${max_row}"
    ws_com.PageSetup.CenterHorizontally = False
    ws_com.PageSetup.CenterVertically = False
    ws_com.PageSetup.Orientation = _XL_LANDSCAPE
    ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_301cd(ws_com, xl_app, dest_filename):
    """Rule 301.C and 301.D — VA vehicle-type exception or fit-to-page."""
    _va_vehicle_types = {
        "Extra Heavy Truck-Tractor", "Extra-Heavy Truck", "Heavy Truck",
        "Heavy Truck-Tractor", "Light Truck", "Medium Truck",
        "Private Passenger Types", "Semitrailer",
        "Service or Utility Trailer", "Trailer",
    }
    if ws_com.Range("B4").Value in _va_vehicle_types:
        # FL gets a top-margin exception for its stamping space
        if "FL" not in dest_filename:
            ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)
        ws_com.PageSetup.Zoom = False
        ws_com.PageSetup.FitToPagesWide = 1
        ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_306(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = False  # explicit: fit width only, not height
    ws_com.PageSetup.PrintTitleRows = "$1:$4"


def _handle_rule_315(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_r1(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f"$A$1:$M${max_row}"


# ===========================================================================
#  RULE REGISTRY
#  Order matters: more-specific prefixes must come before less-specific ones.
#  Each tuple: (sheet-name prefix, handler function)
#
#  HOW TO ADD A RULE:
#    1. Write the handler function above.
#    2. Insert a new tuple here — one line, done.
# ===========================================================================

SHEET_RULES = [
    # Index sheet — special print area
    ("Index",           _handle_index),

    # Rule 222
    ("Rule 222 TTT",    _handle_rule_222ttt),

    # Rule 223
    ("Rule 223 B.5",    _handle_rule_223b5),
    ("Rule 223 C",      _handle_rule_223c),

    # Rule 225
    ("Rule 225 Zone",   _handle_rule_225zone),
    ("Rule 225.C.3",    _handle_rule_225c3),

    # Rule 232
    ("Rule 232 PPT",    _handle_rule_232ppt),

    # Rule 239 — 239 C must come BEFORE generic 239 to prevent wrong match
    ("Rule 239 C",      _handle_rule_239c),
    ("Rule 239 ",       _handle_rule_239_not_c),

    # Rule 240
    ("Rule 240 ",       _handle_rule_240),

    # Rule 255
    ("Rule 255",        _handle_rule_255),

    # Rule 275
    ("Rule 275",        _handle_rule_275),

    # Rule 283
    ("Rule 283",        _handle_rule_283),

    # Rule 289
    ("Rule 289",        _handle_rule_289),

    # Rule 297
    ("Rule 297",        _handle_rule_297),

    # Rule 298
    ("Rule 298",        _handle_rule_298),

    # Rule 301 — 301.C/D before 301.A/B so the more-specific match fires first
    ("Rule 301.C",      _handle_rule_301cd),
    ("Rule 301.D",      _handle_rule_301cd),
    ("Rule 301.A",      _handle_rule_301ab),
    ("Rule 301.B",      _handle_rule_301ab),

    # Rule 306
    ("Rule 306",        _handle_rule_306),

    # Rule 315
    ("Rule 315",        _handle_rule_315),

    # Rule R1
    ("Rule R1",         _handle_rule_r1),
]


def _apply_sheet_rules(sheet_name: str, ws_com, xl_app, dest_filename: str) -> None:
    """
    Walk SHEET_RULES and apply the FIRST matching handler.
    A sheet that matches no prefix keeps only the default PrintTitleRows
    already set in the outer loop.
    """
    for prefix, handler in SHEET_RULES:
        if sheet_name.startswith(prefix):
            handler(ws_com, xl_app, dest_filename)
            return


# ===========================================================================
#  EXCEL COM HELPER
# ===========================================================================

def _kill_excel_instances() -> None:
    """Force-kill any running Excel processes so COM dispatch can proceed."""
    import subprocess
    # /t 2>NUL suppresses the "process not found" error message
    subprocess.call("taskkill /f /im excel.exe 2>NUL", shell=True)


# ===========================================================================
#  PUBLIC ENTRY POINT
# ===========================================================================

def process_pagebreaks(dest_filename1: str, dest_filename2: str) -> None:
    """
    Apply page breaks and print settings to the generated Excel workbook
    via Excel COM automation, then leave it open and visible for the user.

    No openpyxl resave occurs here — the file written by buildBAPages is
    opened directly by Excel COM, which keeps the XML intact.

    Args:
        dest_filename1: Path to the .xlsx file produced by buildBAPages.
        dest_filename2: Intended PDF output path (currently not written;
                        the export block is commented out).

    Pipeline:
        1.  Kill running Excel instances.
        2.  Open file directly via COM (no CorruptLoad — file is already clean).
        3.  Truncate any sheet names exceeding Excel's 31-char limit.
        4.  Set default PrintTitleRows = "$1:$1" for every sheet.
        5.  Apply rule-specific settings via SHEET_RULES registry.
        6.  Hide the Index sheet.
        7.  Save, make Excel visible, leave it open for the user.
    """
    dest_filename1 = os.path.normpath(os.path.abspath(dest_filename1))
    dest_filename2 = os.path.normpath(os.path.abspath(dest_filename2))

    _kill_excel_instances()
    time.sleep(2)

    if pythoncom:
        pythoncom.CoInitialize()

    xl_app = None
    try:
        xl_app = win32com.client.DispatchEx("Excel.Application")
        xl_app.Visible = False
        xl_app.DisplayAlerts = False

        xl_book = xl_app.Workbooks.Open(dest_filename1)

        # Truncate tab names exceeding Excel's 31-character sheet-name limit
        for ws_com in xl_book.Sheets:
            if len(ws_com.Name) > 31:
                ws_com.Name = ws_com.Name[:31]

        # Apply default + rule-specific print settings to every sheet,
        # then replace any hard-coded or openpyxl-set page breaks with
        # content-aware breaks that keep table blocks intact.
        for ws_com in xl_book.Sheets:
            ws_com.PageSetup.PrintTitleRows = "$1:$1"   # universal default
            _apply_sheet_rules(ws_com.Name, ws_com, xl_app, dest_filename1)
            _clear_page_breaks(ws_com)
            _smart_page_breaks(ws_com, xl_app, ws_com.Name)

        # Save without hiding the index
        xl_book.Save()

        # Export all visible sheets as PDF (Index is now visible)
        xl_book.ExportAsFixedFormat(0, dest_filename2, Quality=0)
        print(f"PDF saved: {dest_filename2}")

        xl_book.Close(True)
        xl_app.Quit()
        print("Stage 3: Page Breaks applied and file saved.")

    except Exception as exc:
        print(f"Error during page break processing: {exc}")
        if xl_app:
            try:
                xl_app.Quit()
            except Exception:
                pass
            del xl_app
        gc.collect()
        raise
    finally:
        if pythoncom:
            pythoncom.CoUninitialize()
