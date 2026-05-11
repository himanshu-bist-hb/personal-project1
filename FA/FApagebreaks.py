"""
FA/FApagebreaks.py
==================
Page-break / print-settings post-processing for Farm Auto rate pages.

By default FA uses the IDENTICAL page-break rules as BA.  FA_SHEET_RULES starts
as a copy of BA's rules so any future FA-only rules can be added here without
touching BA code.

HOW TO ADD AN FA-SPECIFIC PAGE-BREAK RULE
------------------------------------------
Step 1 – Write a handler function in this file:

    def _handle_fa_rule_xxx(ws, dest_filename):
        ws.print_area = f"A1:H{ws.max_row}"
        disable_fit_to_page(ws)
        add_break_after(ws, 30)

Step 2 – Append it to FA_SHEET_RULES below:

    FA_SHEET_RULES.append(("Rule FA XXX", _handle_fa_rule_xxx))

    (Add more-specific prefixes BEFORE less-specific ones if needed.)

Step 3 – Done.  FARatePages.py already calls this module's process_pagebreaks,
         which walks FA_SHEET_RULES automatically.

HELPERS (imported from BA for use in any handlers you add here)
-------------------------------------------------------------------
  fit_single_page(ws)       – fit entire sheet onto one printed page
  fit_width_only(ws)        – fit width to 1 page; height grows with content
  disable_fit_to_page(ws)   – turn off fit-to-page; manual breaks rule
  add_break_after(ws, row)  – add a horizontal page break AFTER the given row

OPENPYXL API CHEATSHEET (same as BApagebreaks.py)
  ws.page_setup.orientation     = "portrait" or "landscape"
  ws.print_area                 = "A1:H50"
  ws.print_title_rows           = "1:3"
  ws.page_margins.top           = 1.00   (inches)
"""

import os
import openpyxl

from BA.BApagebreaks import (
    SHEET_RULES as _BA_SHEET_RULES,
    fit_single_page,
    fit_width_only,
    disable_fit_to_page,
    add_break_after,
    _sanitize_xlsx,
    export_to_pdf,        # re-exported: FARatePages.py imports this
)

# ---------------------------------------------------------------------------
# FA_SHEET_RULES: start with all BA rules, then append FA-only rules below.
# ---------------------------------------------------------------------------
FA_SHEET_RULES = list(_BA_SHEET_RULES)

# ── Add FA-only rules here ──────────────────────────────────────────────────
# Example (uncomment and fill in when needed):
#
#   def _handle_fa_farm_machinery(ws, dest_filename):
#       fit_single_page(ws)
#       ws.print_title_rows = "1:3"
#
#   FA_SHEET_RULES.append(("Rule FA Farm Machinery", _handle_fa_farm_machinery))


# ---------------------------------------------------------------------------
# process_pagebreaks: FA version — uses FA_SHEET_RULES instead of BA default
# ---------------------------------------------------------------------------

def process_pagebreaks(dest_filename1, dest_filename2):
    """
    Apply page breaks / print settings to dest_filename1 using FA_SHEET_RULES.

    dest_filename2 is accepted for backward compatibility (was a PDF path).
    FA_SHEET_RULES contains all BA rules plus any FA-only additions above,
    so adding a FA rule here automatically takes effect — no other file changes.
    """
    print(f"[FApagebreaks] Processing: {dest_filename1}")
    dest_filename1 = os.path.normpath(os.path.abspath(dest_filename1))

    workbook = openpyxl.load_workbook(dest_filename1)

    # Truncate sheet names exceeding Excel's 31-character limit
    for original_name in list(workbook.sheetnames):
        if len(original_name) > 31:
            workbook[original_name].title = original_name[:31]

    # Apply defaults + FA rules to every sheet
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        ws.print_title_rows = "1:1"
        fit_single_page(ws)
        for prefix, handler in FA_SHEET_RULES:
            if sheet_name.startswith(prefix):
                handler(ws, dest_filename1)
                break

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
    print("[FApagebreaks] Done.")
