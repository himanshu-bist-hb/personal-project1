\"\"\"
BApagebreaks.py
===============
Post-processing step: apply page breaks / print settings to the generated
Excel workbook via Excel COM automation.

This version uses Section-Aware grouping and dynamic marker detection to 
keep tables and their notes together on a single page, preventing splits.
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
    \"\"\"Completely reset all manual and automatic page breaks.\"\"\"
    try:
        ws_com.ResetAllPageBreaks()
    except Exception:
        pass

def _break_at_text(ws_com, col_a, search_text, offset=0):
    \"\"\"Search column A for search_text and add a horizontal break at that row + offset.\"\"\"
    for i, val in enumerate(col_a):
        if val[0] and search_text in str(val[0]):
            try:
                ws_com.HPageBreaks.Add(Before=ws_com.Rows(i + 1 + offset))
                return i + 1 + offset
            except Exception:
                pass
    return None

def _break_before_subtitles(ws_com, max_row, col_a, cooldown=8):
    \"\"\"
    Intelligently adds page breaks before new sections.
    A section is identified as a non-blank cell preceded by a blank cell.
    \"\"\"
    last_break_row = 0
    for i in range(4, max_row):
        row_num = i + 1
        prev_val = col_a[i - 1][0]
        curr_val = col_a[i][0]
        
        is_prev_blank = prev_val is None or str(prev_val).strip() == \"\"
        is_curr_data  = curr_val is not None and str(curr_val).strip() != \"\"
        
        if is_curr_data and is_prev_blank:
            if (row_num - last_break_row) > cooldown:
                try:
                    ws_com.HPageBreaks.Add(Before=ws_com.Rows(row_num))
                    last_break_row = row_num
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
    max_row = ws_com.UsedRange.Rows.Count
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    _break_before_subtitles(ws_com, max_row, col_a, cooldown=10)

def _handle_rule_222ttt(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1
    ws_com.PageSetup.PrintTitleRows = \"$1:$3\"

def _handle_rule_223b5(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.Orientation = _XL_LANDSCAPE

def _handle_rule_223c(ws_com, xl_app, dest_filename):
    ws_com.PageSetup.FitToPagesTall = 1

def _handle_rule_225zone(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$M${max_row}\"
    ws_com.PageSetup.CenterHorizontally = False
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    for i, val in enumerate(col_a):
        text = str(val[0]).upper() if val[0] else \"\"
        if i > 5 and any(marker in text for marker in [\"ZONE\", \"MEDICAL PAYMENTS\", \"PERSONAL INJURY\"]):
            try:
                ws_com.HPageBreaks.Add(Before=ws_com.Rows(i + 1))
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
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$H${max_row}\"
    ws_com.PageSetup.CenterHorizontally = False
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    
    # Force break before the notes for 255.D to keep them with the table
    if not _break_at_text(ws_com, col_a, \"Apply a factor of 0.74\"):
        _break_at_text(ws_com, col_a, \"255.D. Deductibles\")
    
    # Enhancement Endorsement break
    _break_at_text(ws_com, col_a, \"255.E.2\")

def _handle_rule_283(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    ws_com.PageSetup.PrintArea = f\"$A$1:$P${max_row}\"
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    _break_before_subtitles(ws_com, max_row, col_a, cooldown=8)

def _handle_rule_301abc(ws_com, xl_app, dest_filename):
    va_types = [\"Extra Heavy Truck-Tractor\", \"Extra-Heavy Truck\", \"Heavy Truck\",
                \"Heavy Truck-Tractor\", \"Light Truck\", \"Medium Truck\", \"Private Passenger Types\",
                \"Semitrailer\", \"Service or Utility Trailer\", \"Trailer\"]
    
    if ws_com.Range(\"B4\").Value in va_types:
        if \"FL\" not in dest_filename:
            ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)
        ws_com.PageSetup.FitToPagesTall = 1
    else:
        max_row = ws_com.UsedRange.Rows.Count
        if ws_com.Name.startswith(\"Rule 301.B\"):
            ws_com.PageSetup.PrintArea = f\"$A$1:$T${max_row}\"
        for row in range(46, max_row, 45):
            try:
                ws_com.HPageBreaks.Add(Before=ws_com.Rows(row + 1))
            except Exception:
                pass
        ws_com.PageSetup.CenterHorizontally = False
        ws_com.PageSetup.Orientation = _XL_LANDSCAPE
        ws_com.PageSetup.TopMargin = xl_app.InchesToPoints(1.00)

def _handle_generic_rule(ws_com, xl_app, dest_filename):
    max_row = ws_com.UsedRange.Rows.Count
    col_a = ws_com.Range(f\"A1:A{max_row}\").Value
    _break_before_subtitles(ws_com, max_row, col_a, cooldown=8)

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
    (\"Rule 275\",        _handle_generic_rule),
    (\"Rule 283\",        _handle_rule_283),
    (\"Rule 289\",        _handle_generic_rule),
    (\"Rule 301.A\",      _handle_rule_301abc),
    (\"Rule 301.B\",      _handle_rule_301abc),
    (\"Rule 301.C\",      _handle_rule_301abc),
    (\"Rule 301.D\",      _handle_rule_301abc),
    (\"Rule 306\",        _handle_generic_rule),
    (\"Rule 315\",        _handle_generic_rule),
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
            ws_com.PageSetup.Zoom = False
            ws_com.PageSetup.FitToPagesWide = 1
            ws_com.PageSetup.FitToPagesTall = False
            ws_com.PageSetup.PrintTitleRows = \"$1:$1\"
            
            _clear_page_breaks(ws_com)
            _apply_sheet_rules(ws_com.Name, ws_com, xl_app, dest_filename1)

        xl_book.Save()
        xl_book.ExportAsFixedFormat(0, dest_filename2, Quality=0)
        
        xl_book.Close(True)
        xl_app.Quit()
        print(\"Stage 3: Professional Page Breaks applied successfully.\")

    except Exception as exc:
        print(f\"Error: {exc}\")
        if xl_app:
            xl_app.Quit()
        raise
    finally:
        if pythoncom:
            pythoncom.CoUninitialize()
