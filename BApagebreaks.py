\"\"\"
BApagebreaks.py
===============
Post-processing step: apply page breaks / print settings to the generated
Excel workbook via Excel COM automation, then leave it open for the user.

VERSION: Ported from BApagebreaks_old.py with COM-direct implementation.
\"\"\"

import gc
import os
import time


try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None


# ===========================================================================
#  COM CONSTANTS
# ===========================================================================

_XL_LANDSCAPE = 2   # xlLandscape


# ===========================================================================
#  PAGE BREAK HELPERS
# ===========================================================================

def _clear_page_breaks(ws_com):
    \"\"\"Remove all existing horizontal page breaks.\"\"\"
    count = ws_com.HPageBreaks.Count
    for _ in range(count, 0, -1):
        ws_com.HPageBreaks(1).Delete()


def _break_before_subtitles(ws_com, max_row, col_a):
    \"\"\"Insert a page break before every sub-table subtitle row except the first.
    
    A subtitle row is defined as a non-blank cell in Column A preceded by a blank cell.
    \"\"\"
    for row_num in range(2, max_row + 1):
        prev_val = col_a[row_num - 2][0]
        curr_val = col_a[row_num - 1][0]
        
        prev_blank = prev_val is None or str(prev_val).strip() == \"\"
        curr_has_val = curr_val is not None and str(curr_val).strip() != \"\"
        
        if curr_has_val and prev_blank:
            # Skip the first few rows (title/first subtitle)
            if row_num > 4:
                ws_com.HPageBreaks.Add(ws_com.Rows(row_num))


# ===========================================================================
#  RULE HANDLER FUNCTIONS
#  Ported from BApagebreaks_old.py logic.
# ===========================================================================

def _handle_index(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintTitleRows = \"\"
    ws_com.PageSetup.PrintArea = f\"$A$1:$J${max_row}\"
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = False


def _handle_rule_222b(ws_com, xl_app, dest_filename):
    # Old logic: Break(25), Break(49)
    ws_com.HPageBreaks.Add(ws_com.Rows(26))
    ws_com.HPageBreaks.Add(ws_com.Rows(50))


def _handle_rule_222ttt(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"


def _handle_rule_223b5(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Orientation = _XL_LANDSCAPE


def _handle_rule_223c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_225zone(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$M${max_row}\"
    ws_com.PageSetup.CenterHorizontally = False
    ws_com.PageSetup.CenterVertically = False
    # Old logic: range(52, max_row, 51) -> Break(row)
    for row in range(52, max_row, 51):
        ws_com.HPageBreaks.Add(ws_com.Rows(row + 1))


def _handle_rule_225c3(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_232ppt(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"


def _handle_rule_239_not_c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"


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
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"
    ws_com.PageSetup.PrintArea = f\"$A$1:$M${max_row}\"
    ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_255(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$H${max_row}\"
    ws_com.PageSetup.CenterHorizontally = False
    ws_com.PageSetup.CenterVertically = False
    # Old logic: Break(37)
    ws_com.HPageBreaks.Add(ws_com.Rows(38))


def _handle_rule_275(ws_com, xl_app, dest_filename):
    # Old logic: check A10
    if ws_com.Range(\"A10\").Value == \"275.B.1.(b). Short Term - Autos Leased by the Hour, Day, or Week\":
        ws_com.PageSetup.PrintTitleRows = \"$1:$1\"
        ws_com.PageSetup.Zoom = False
        ws_com.PageSetup.FitToPagesWide = 1
        ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_283(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$P${max_row}\"
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    
    target_values = [\"283.B Limited Specified Causes of Loss\",
                     \"283.B Comprehensive\",
                     \"283.B Blanket Collision\"]
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    for row_num in range(1, max_row + 1):
        cell_val = str(col_a[row_num - 1][0])
        if cell_val in target_values and row_num > 3:
            ws_com.HPageBreaks.Add(ws_com.Rows(row_num))


def _handle_rule_289(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$H${max_row}\"
    # Old logic: Break(37)
    ws_com.HPageBreaks.Add(ws_com.Rows(38))


def _handle_rule_297(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$P${max_row}\"
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    occurrence_count = 0
    for row_num in range(1, max_row + 1):
        cell_val = str(col_a[row_num - 1][0])
        if cell_val.startswith('Single') or cell_val.startswith('Uninsured'):
            occurrence_count += 1
            if (occurrence_count % 3 == 0) and (occurrence_count != 0):
                occurrence_count += 1
                ws_com.HPageBreaks.Add(ws_com.Rows(row_num))


def _handle_rule_298(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$K${max_row}\"
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    occurrence_count = 0
    for row_num in range(1, max_row + 1):
        cell_val = str(col_a[row_num - 1][0])
        if cell_val.startswith('298'):
            occurrence_count += 1
            if occurrence_count == 4:
                occurrence_count += 1
                ws_com.HPageBreaks.Add(ws_com.Rows(row_num))
            if occurrence_count == 8:
                break


def _handle_rule_301abc(ws_com, xl_app, dest_filename):
    va_types = [\"Extra Heavy Truck-Tractor\", \"Extra-Heavy Truck\", \"Heavy Truck\",
                \"Heavy Truck-Tractor\", \"Light Truck\", \"Medium Truck\", \"Private Passenger Types\",
                \"Semitrailer\", \"Service or Utility Trailer\", \"Trailer\"]
    
    b4_val = ws_com.Range(\"B4\").Value
    if b4_val in va_types:
        # Rule 301.C / D logic for VA types
        if \"FL\" not in dest_filename:
            ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)
        ws_com.PageSetup.Zoom = False
        ws_com.PageSetup.FitToPagesWide = 1
        ws_com.PageSetup.FitToPagesTall = 1
    else:
        # Generic 301 logic
        ws_com.PageSetup.Zoom = False
        ws_com.PageSetup.FitToPagesWide = 1
        if ws_com.Name.startswith(\"Rule 301.B\"):
            ws_com.PageSetup.PrintArea = f\"$A$1:$T${ws_com.UsedRange.Rows.Count}\"
        
        # Old logic: range(46, max_row, 45) -> Break(row)
        max_row = ws_com.UsedRange.Rows.Count
        for row in range(46, max_row, 45):
            ws_com.HPageBreaks.Add(ws_com.Rows(row + 1))
            
        ws_com.PageSetup.CenterHorizontally = False
        ws_com.PageSetup.CenterVertically = False
        ws_com.PageSetup.Orientation = _XL_LANDSCAPE
        ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_306(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.PrintTitleRows = \"$1:$4\"


def _handle_rule_315(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    ws_com.PageSetup.FitToPagesTall = 1
    # Old logic: Break(23)
    ws_com.HPageBreaks.Add(ws_com.Rows(24))


def _handle_rule_r1(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$M${max_row}\"
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    occurrence_count = 0
    for row_num in range(1, max_row + 1):
        cell_val = str(col_a[row_num - 1][0])
        if cell_val.startswith('R1'):
            occurrence_count += 1
            if occurrence_count == 3 or occurrence_count == 6:
                occurrence_count += 1
                ws_com.HPageBreaks.Add(ws_com.Rows(row_num))


def _handle_generic_rule(ws_com, xl_app, dest_filename):
    \"\"\"Catch-all for any 'Rule' sheet not explicitly handled.
    Applies FitToPagesWide and Subtitle detection to keep tables together.
    \"\"\"
    ws_com.PageSetup.Zoom = False
    ws_com.PageSetup.FitToPagesWide = 1
    max_row = ws_com.UsedRange.Rows.Count
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    _break_before_subtitles(ws_com, max_row, col_a)


# ===========================================================================
#  RULE REGISTRY
# ===========================================================================

SHEET_RULES = [
    (\"Index\",           _handle_index),
    (\"Rule 222 B\",      _handle_rule_222b),
    (\"Rule 222 TTT\",    _handle_rule_222ttt),
    (\"Rule 223 B.5\",    _handle_rule_223b5),
    (\"Rule 223 C\",      _handle_rule_223c),
    (\"Rule 225 Zone\",   _handle_rule_225zone),
    (\"Rule 225.C.3\",    _handle_rule_225c3),
    (\"Rule 232 PPT\",    _handle_rule_232ppt),
    (\"Rule 239 C\",      _handle_rule_239c),
    (\"Rule 239 \",       _handle_rule_239_not_c),
    (\"Rule 240 \",       _handle_rule_240),
    (\"Rule 255\",        _handle_rule_255),
    (\"Rule 275\",        _handle_rule_275),
    (\"Rule 283\",        _handle_rule_283),
    (\"Rule 289\",        _handle_rule_289),
    (\"Rule 297\",        _handle_rule_297),
    (\"Rule 298\",        _handle_rule_298),
    (\"Rule 301.A\",      _handle_rule_301abc),
    (\"Rule 301.B\",      _handle_rule_301abc),
    (\"Rule 301.C\",      _handle_rule_301abc),
    (\"Rule 301.D\",      _handle_rule_301abc),
    (\"Rule 306\",        _handle_rule_306),
    (\"Rule 315\",        _handle_rule_315),
    (\"Rule R1\",         _handle_rule_r1),
    (\"Rule \",           _handle_generic_rule), # Catch-all
]


def _apply_sheet_rules(sheet_name, ws_com, xl_app, dest_filename):
    for prefix, handler in SHEET_RULES:
        if sheet_name.startswith(prefix):
            handler(ws_com, xl_app, dest_filename)
            return


# ===========================================================================
#  PUBLIC ENTRY POINT
# ===========================================================================

def process_pagebreaks(dest_filename1, dest_filename2):
    dest_filename1 = os.path.normpath(os.path.abspath(dest_filename1))
    dest_filename2 = os.path.normpath(os.path.abspath(dest_filename2))

    # Kill Excel
    import subprocess
    subprocess.call(\"taskkill /f /im excel.exe 2>NUL\", shell=True)
    time.sleep(2)

    if pythoncom:
        pythoncom.CoInitialize()

    xl_app = None
    try:
        xl_app = win32com.client.DispatchEx(\"Excel.Application\")
        xl_app.Visible = False
        xl_app.DisplayAlerts = False

        xl_book = xl_app.Workbooks.Open(dest_filename1)

        for ws_com in xl_book.Sheets:
            # Universal defaults
            ws_com.PageSetup.PrintTitleRows = \"$1:$1\"
            _clear_page_breaks(ws_com)
            
            # Apply rules
            _apply_sheet_rules(ws_com.Name, ws_com, xl_app, dest_filename1)

        xl_book.Save()
        xl_book.ExportAsFixedFormat(0, dest_filename2, Quality=0)
        
        xl_book.Close(True)
        xl_app.Quit()
        print(\"Stage 3: Page Breaks applied successfully.\")

    except Exception as exc:
        print(f\"Error: {exc}\")
        if xl_app:
            xl_app.Quit()
        raise
    finally:
        if pythoncom:
            pythoncom.CoUninitialize()
