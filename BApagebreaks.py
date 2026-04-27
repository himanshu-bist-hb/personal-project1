\"\"\"
BApagebreaks.py
===============
Post-processing step: apply page breaks / print settings to the generated
Excel workbook via Excel COM automation, then leave it open for the user.

This version uses dynamic detection to keep tables and their notes together,
preventing mid-table cuts and unwanted vertical page breaks.
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
        try:
            ws_com.HPageBreaks(1).Delete()
        except Exception:
            pass


def _break_at_text(ws_com, col_a, search_text, offset=0):
    \"\"\"Search column A for search_text and add a horizontal break at that row + offset.\"\"\"
    for i, val in enumerate(col_a):
        if val[0] and search_text in str(val[0]):
            try:
                ws_com.HPageBreaks.Add(ws_com.Rows(i + 1 + offset))
                return True
            except Exception:
                pass
    return False


def _break_before_subtitles(ws_com, max_row, col_a):
    \"\"\"Insert a page break before every sub-table subtitle row.
    
    A subtitle is identified as a non-blank cell in Column A preceded by a blank cell.
    \"\"\"
    for row_num in range(2, max_row + 1):
        prev_val = col_a[row_num - 2][0]
        curr_val = col_a[row_num - 1][0]
        
        prev_blank = prev_val is None or str(prev_val).strip() == \"\"
        curr_has_val = curr_val is not None and str(curr_val).strip() != \"\"
        
        if curr_has_val and prev_blank:
            # Skip titles at the very top
            if row_num > 4:
                try:
                    ws_com.HPageBreaks.Add(ws_com.Rows(row_num))
                except Exception:
                    pass


# ===========================================================================
#  RULE HANDLER FUNCTIONS
# ===========================================================================

def _handle_index(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintTitleRows = \"\"
    ws_com.PageSetup.PrintArea = f\"$A$1:$J${max_row}\"


def _handle_rule_222b(ws_com, xl_app, dest_filename):
    # Dynamic break before the second and third tables
    max_row = ws_com.UsedRange.Rows.Count
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    _break_before_subtitles(ws_com, max_row, col_a)


def _handle_rule_222ttt(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"


def _handle_rule_223b5(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Orientation = _XL_LANDSCAPE


def _handle_rule_223c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_225zone(ws_com, xl_app, dest_filename):
    \"\"\"Force breaks before each Zone base rate table to prevent cutting.\"\"\"
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$M${max_row}\"
    ws_com.PageSetup.CenterHorizontally = False
    ws_com.PageSetup.CenterVertically = False
    
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    for i, val in enumerate(col_a):
        text = str(val[0]) if val[0] else \"\"
        # Break before headings like \"ZONE 1\", \"ZONE 2\", or the factors at the end
        if \"ZONE\" in text.upper() or \"MEDICAL PAYMENTS\" in text.upper() or \"PERSONAL INJURY\" in text.upper():
            if i > 5: # Don't break at the very top
                try:
                    ws_com.HPageBreaks.Add(ws_com.Rows(i + 1))
                except Exception:
                    pass


def _handle_rule_225c3(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_232ppt(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"


def _handle_rule_239c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_240(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.CenterVertically = True
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"
    ws_com.PageSetup.PrintArea = f\"$A$1:$M${ws_com.UsedRange.Rows.Count}\"
    ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_255(ws_com, xl_app, dest_filename):
    \"\"\"Ensures the Deductibles table and its notes are on a fresh page.\"\"\"
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$H${max_row}\"
    ws_com.PageSetup.CenterHorizontally = False
    ws_com.PageSetup.CenterVertically = False
    
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    # Search for the notes preceding the Deductibles section to keep them with the table
    if not _break_at_text(ws_com, col_a, \"Apply a factor of 0.74\"):
        # Fallback to the heading itself if notes aren't found
        _break_at_text(ws_com, col_a, \"255.D. Deductibles\")
    
    # Also break before the Enhancement Endorsement section
    _break_at_text(ws_com, col_a, \"255.E.2\")


def _handle_rule_275(ws_com, xl_app, dest_filename):
    if ws_com.Range(\"A10\").Value == \"275.B.1.(b). Short Term - Autos Leased by the Hour, Day, or Week\":
        ws_com.PageSetup.PrintTitleRows = \"$1:$1\"
        ws_com.PageSetup.FitToPagesTall = 1


def _handle_rule_283(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$P${max_row}\"
    
    target_values = [\"283.B Limited Specified Causes of Loss\",
                     \"283.B Comprehensive\",
                     \"283.B Blanket Collision\"]
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    for row_num in range(1, max_row + 1):
        cell_val = str(col_a[row_num - 1][0])
        if cell_val in target_values and row_num > 3:
            try:
                ws_com.HPageBreaks.Add(ws_com.Rows(row_num))
            except Exception:
                pass


def _handle_rule_289(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$H${max_row}\"
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    # Dynamic break before the second table
    _break_before_subtitles(ws_com, max_row, col_a)


def _handle_rule_301abc(ws_com, xl_app, dest_filename):
    va_types = [\"Extra Heavy Truck-Tractor\", \"Extra-Heavy Truck\", \"Heavy Truck\",
                \"Heavy Truck-Tractor\", \"Light Truck\", \"Medium Truck\", \"Private Passenger Types\",
                \"Semitrailer\", \"Service or Utility Trailer\", \"Trailer\"]
    
    b4_val = ws_com.Range(\"B4\").Value
    if b4_val in va_types:
        if \"FL\" not in dest_filename:
            ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)
        ws_com.PageSetup.FitToPagesTall = 1
    else:
        if ws_com.Name.startswith(\"Rule 301.B\"):
            ws_com.PageSetup.PrintArea = f\"$A$1:$T${ws_com.UsedRange.Rows.Count}\"
        
        # Periodic breaks for large tables
        max_row = ws_com.UsedRange.Rows.Count
        for row in range(46, max_row, 45):
            try:
                ws_com.HPageBreaks.Add(ws_com.Rows(row + 1))
            except Exception:
                pass
            
        ws_com.PageSetup.CenterHorizontally = False
        ws_com.PageSetup.CenterVertically = False
        ws_com.PageSetup.Orientation = _XL_LANDSCAPE
        ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)


def _handle_rule_306(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.PrintTitleRows = \"$1:$4\"


def _handle_rule_315(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1
    # Dynamic breaks before duration tables
    max_row = ws_com.UsedRange.Rows.Count
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    _break_before_subtitles(ws_com, max_row, col_a)


def _handle_generic_rule(ws_com, xl_app, dest_filename):
    \"\"\"Catch-all to keep any Rule tables together.\"\"\"
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
    (\"Rule 239 \",       _handle_generic_rule),
    (\"Rule 240 \",       _handle_rule_240),
    (\"Rule 255\",        _handle_rule_255),
    (\"Rule 275\",        _handle_rule_275),
    (\"Rule 283\",        _handle_rule_283),
    (\"Rule 289\",        _handle_rule_289),
    (\"Rule 297\",        _handle_generic_rule),
    (\"Rule 298\",        _handle_generic_rule),
    (\"Rule 301.A\",      _handle_rule_301abc),
    (\"Rule 301.B\",      _handle_rule_301abc),
    (\"Rule 301.C\",      _handle_rule_301abc),
    (\"Rule 301.D\",      _handle_rule_301abc),
    (\"Rule 306\",        _handle_rule_306),
    (\"Rule 315\",        _handle_rule_315),
    (\"Rule R1\",         _handle_generic_rule),
    (\"Rule \",           _handle_generic_rule),
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
            # Universal Defaults: Eliminate vertical breaks and force scaling
            ws_com.PageSetup.PrintTitleRows = \"$1:$1\"
            ws_com.PageSetup.Zoom = False
            ws_com.PageSetup.FitToPagesWide = 1
            ws_com.PageSetup.FitToPagesTall = False
            
            _clear_page_breaks(ws_com)
            
            # Apply specialized and dynamic rules
            _apply_sheet_rules(ws_com.Name, ws_com, xl_app, dest_filename1)

        xl_book.Save()
        xl_book.ExportAsFixedFormat(0, dest_filename2, Quality=0)
        
        xl_book.Close(True)
        xl_app.Quit()
        print(\"Stage 3: Professional Page Breaks applied and tables kept together.\")

    except Exception as exc:
        print(f\"Error: {exc}\")
        if xl_app:
            xl_app.Quit()
        raise
    finally:
        if pythoncom:
            pythoncom.CoUninitialize()
