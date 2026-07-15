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
import subprocess
import time
import zipfile
from io import BytesIO

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break, ColBreak, RowBreak


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
    ws.print_area = f"A1:J{ws.max_row}"
    fit_width_only(ws)
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
    ws.print_area = f"A1:G{ws.max_row}"
    ws.print_options.horizontalCentered = False
    ws.print_options.verticalCentered = False
    fit_width_only(ws)
    add_break_after(ws, 51)
    add_break_after(ws, 103)


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
    ws.print_area = f"A1:G{ws.max_row}"
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
    ws.print_area = f"A1:G{ws.max_row}"
    ws.col_breaks = ColBreak()
    ws.row_breaks = RowBreak()
    fit_width_only(ws)
    occurrence_count = 0
    for row in range(1, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=1).value)
        if cell_value.startswith("Single") or cell_value.startswith("Uninsured"):
            occurrence_count += 1
        if (occurrence_count % 3 == 0) and (occurrence_count != 0):
            occurrence_count += 1
            add_break_after(ws, row - 1)


def _handle_rule_298(ws, dest_filename):
    ws.print_area = f"A1:G{ws.max_row}"
    ws.col_breaks = ColBreak()
    ws.row_breaks = RowBreak()
    fit_width_only(ws)
    add_break_after(ws, 33)
    add_break_after(ws, 57)
    add_break_after(ws, 88)


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


def _handle_rule_301c(ws, dest_filename):
    if ws["B4"].value in _VA_VEHICLE_TYPES:
        if "FL" not in dest_filename:
            ws.page_margins.top = 1.00
        fit_single_page(ws)
        return

    # Factor columns (D onwards) default to ~12 char-units wide (0.94" each).
    # With 26 such columns, Set 2 is ~28" wide — forcing fit_width_only to
    # scale down to ~33%.  At 33%, Set 1 (7 cols) fills only ~26% of the page.
    # Width 6 reduces Set 2 to ~16.6", giving a ~57% scale.  At 57% scale,
    # the automatic page break falls at row ~46, coinciding with the manual
    # break — no spurious dashed lines appear within any section.  Width 5
    # gave 65% scale but that put the auto break at row ~41 (before the
    # manual break at 46), creating unwanted intermediate pages.
    for col_idx in range(4, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 6

    ws.col_breaks = ColBreak()
    ws.row_breaks = RowBreak()
    ws.page_setup.orientation = "landscape"
    ws.print_area = f"A1:AC{ws.max_row}"
    ws.print_options.horizontalCentered = False
    ws.print_options.verticalCentered = False
    fit_width_only(ws)

    for row in [46, 91, 136, 181, 226, 271, 316, 361]:
        add_break_after(ws, row)
    for row in [406, 451, 496, 541, 586, 631, 676, 720]:
        add_break_after(ws, row)


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

    ("Rule 301.C",   _handle_rule_301c),
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


# ============================================================================
#  PDF EXPORT  -  drives Excel via COM to preserve every print setting
#  (page breaks, fit-to-page, print areas, orientation, margins, headers).
# ============================================================================

_XL_TYPE_PDF        = 0   # xlTypePDF
_XL_QUALITY_STD     = 0   # xlQualityStandard


def _compress_pdf(src_path, dest_path):
    """
    Write a compressed copy of src_path to dest_path.

    Excel's ExportAsFixedFormat produces a "tagged" PDF: every table cell
    gets an accessibility structure element, which for rate pages (hundreds
    of pages of dense tables) is most of the file — a 10 MB workbook can
    export as a 170 MB PDF. The DOI filing copy doesn't need the tag tree,
    so drop it and recompress every stream (measured ~5x smaller).

    Best-effort: if pikepdf is not installed or anything fails, the raw
    PDF is moved to dest_path unchanged so export never breaks.
    """
    try:
        import pikepdf
    except ImportError:
        print("[pagebreaks] pikepdf not installed — PDF left uncompressed "
              "(pip install pikepdf to shrink output ~5x)")
        shutil.move(src_path, dest_path)
        return

    try:
        with pikepdf.open(src_path) as pdf:
            root = pdf.Root
            for key in ("/StructTreeRoot", "/MarkInfo", "/Metadata",
                        "/PieceInfo", "/Lang"):
                if key in root:
                    del root[key]
            for page in pdf.pages:
                for key in ("/StructParents", "/Tabs", "/PieceInfo"):
                    if key in page:
                        del page[key]
            pdf.remove_unreferenced_resources()
            pdf.save(dest_path,
                     compress_streams=True,
                     recompress_flate=True,
                     object_stream_mode=pikepdf.ObjectStreamMode.generate)
        raw_mb  = os.path.getsize(src_path) / 1e6
        out_mb  = os.path.getsize(dest_path) / 1e6
        print(f"[pagebreaks] PDF compressed {raw_mb:.1f} MB -> {out_mb:.1f} MB")
        os.remove(src_path)
    except Exception as exc:
        print(f"[pagebreaks] PDF compression failed ({exc}) — keeping raw PDF")
        shutil.move(src_path, dest_path)


def _kill_excel_instances():
    """Best-effort: terminate any orphan EXCEL.EXE so COM gets a clean slate."""
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "excel.exe"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
        )
    except Exception:
        pass


# ============================================================================
#  PARALLEL PDF EXPORT
#
#  Excel's PDF renderer is single-threaded at ~8k cells/second, so a sheet
#  like BOP's TRDEF (82k rows x 21 cols = 1.7M cells) dominates the export.
#  When a workbook contains a sheet that big, several Excel instances each
#  export a different slice of the workbook — a huge sheet is sliced by ROW
#  RANGE via an in-memory print area, so the .xlsx keeps its single tab and
#  the printed footer/tab name is unchanged. Page numbers stay continuous:
#  every worker first counts its own pages (phase 1), then the cumulative
#  start is stamped via FirstPageNumber before exporting (phase 2).
#  Measured on TRDEF-scale data: 224s single-instance -> 86s with 4 workers.
#
#  Any failure falls back to the normal single-instance export.
# ============================================================================

_HUGE_SHEET_ROWS  = 10_000   # a sheet taller than this triggers the parallel path
_PARALLEL_WORKERS = 4


def _plan_parallel_export(xlsx_path):
    """
    Return a list of worker groups, or None when a plain single-instance
    export is the right choice. Each group is a list of
    (sheet_name, first_row, last_row, max_col) units in workbook order;
    first_row/last_row are None for "whole sheet".
    """
    try:
        import pikepdf  # required to merge the per-worker PDFs
    except ImportError:
        return None

    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    try:
        sheets = []
        for ws in wb.worksheets:
            if ws.title == "Index" or ws.sheet_state != "visible":
                continue
            sheets.append((ws.title, ws.max_row or 1, ws.max_column or 1))
    finally:
        wb.close()

    if not any(rows > _HUGE_SHEET_ROWS for _, rows, _ in sheets):
        return None

    # Units: whole small sheets; huge sheets sliced into row ranges of
    # roughly equal weight so the groups balance.
    total_weight = sum(rows * cols for _, rows, cols in sheets)
    target = total_weight / _PARALLEL_WORKERS
    units = []
    for name, rows, cols in sheets:
        weight = rows * cols
        if rows <= _HUGE_SHEET_ROWS:
            units.append((name, None, None, cols, weight))
            continue
        n_slices = max(2, min(_PARALLEL_WORKERS * 2, round(weight / target) or 2))
        step = -(-rows // n_slices)  # ceil
        r = 1
        while r <= rows:
            r2 = min(r + step - 1, rows)
            units.append((name, r, r2, cols, (r2 - r + 1) * cols))
            r = r2 + 1

    # Contiguous partition into up to _PARALLEL_WORKERS groups, coalescing
    # adjacent slices of the same sheet that land in the same group (a
    # worker can only print one area per sheet).
    groups, cur, cur_weight = [], [], 0
    for unit in units:
        name, r1, r2, cols, weight = unit
        if cur and cur_weight >= target and len(groups) < _PARALLEL_WORKERS - 1:
            groups.append(cur)
            cur, cur_weight = [], 0
        if cur and cur[-1][0] == name and cur[-1][2] is not None and r1 is not None:
            prev = cur[-1]
            cur[-1] = (name, prev[1], r2, cols)
        else:
            cur.append((name, r1, r2, cols))
        cur_weight += weight
    if cur:
        groups.append(cur)

    return groups if len(groups) >= 2 else None


def _parallel_worker(xlsx_path, group, out_pdf, idx, counts, counted,
                     starts, go, errors):
    """Phase 1: isolate this group's slice and count its pages. Phase 2 (after
    the coordinator computes cumulative starts): stamp FirstPageNumber and
    export. Runs in its own thread with its own Excel instance."""
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(xlsx_path, ReadOnly=True, UpdateLinks=0)

        by_name = {u[0]: u for u in group}
        first_sheet = None
        n_pages = 0
        for sheet in workbook.Sheets:
            unit = by_name.get(sheet.Name)
            if unit is None:
                sheet.Visible = 0  # xlSheetHidden — in-memory only (ReadOnly)
                continue
            _, r1, r2, cols = unit
            if r1 is not None:
                area = f"A{r1}:{get_column_letter(cols)}{r2}"
                sheet.PageSetup.PrintArea = area
            if first_sheet is None:
                first_sheet = sheet
            n_pages += sheet.PageSetup.Pages.Count

        counts[idx] = n_pages
        counted[idx].set()

        go.wait()
        if errors:          # another worker failed during phase 1 — abort
            return
        first_sheet.PageSetup.FirstPageNumber = starts[idx]
        first_sheet.Activate()
        workbook.ExportAsFixedFormat(
            Type=_XL_TYPE_PDF,
            Filename=out_pdf,
            Quality=_XL_QUALITY_STD,
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    except Exception as exc:
        errors.append(exc)
        counted[idx].set()
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def _export_pdf_parallel(xlsx_path, groups, pdf_path, progress):
    """Run the worker pool, then merge + strip + compress into pdf_path.
    Raises on any failure (caller falls back to the sequential export)."""
    import tempfile
    import threading
    import pikepdf

    k = len(groups)
    part_pdfs = [
        os.path.join(tempfile.gettempdir(),
                     f"~rate_pages_part{i}_{os.getpid()}.pdf")
        for i in range(k)
    ]
    counts  = [None] * k
    counted = [threading.Event() for _ in range(k)]
    starts  = [1] * k
    go      = threading.Event()
    errors  = []

    progress(f"PDF — rendering with {k} Excel instances in parallel...")
    threads = [
        threading.Thread(
            target=_parallel_worker,
            args=(xlsx_path, groups[i], part_pdfs[i], i,
                  counts, counted, starts, go, errors),
            daemon=True,
        )
        for i in range(k)
    ]
    try:
        for t in threads:
            t.start()
        for evt in counted:
            evt.wait()
        if errors:
            raise errors[0]
        for i in range(1, k):
            starts[i] = starts[i - 1] + counts[i - 1]
        go.set()
        for t in threads:
            t.join()
        if errors:
            raise errors[0]
    finally:
        go.set()  # never leave workers blocked

    for p in part_pdfs:
        if not (os.path.exists(p) and os.path.getsize(p) > 0):
            raise RuntimeError(f"parallel export part missing: {p}")

    progress("PDF — merging and compressing...")
    try:
        merged = pikepdf.new()
        parts = [pikepdf.open(p) for p in part_pdfs]
        try:
            for part in parts:
                merged.pages.extend(part.pages)
            root = merged.Root
            for key in ("/StructTreeRoot", "/MarkInfo", "/Metadata",
                        "/PieceInfo", "/Lang"):
                if key in root:
                    del root[key]
            for page in merged.pages:
                for key in ("/StructParents", "/Tabs", "/PieceInfo"):
                    if key in page:
                        del page[key]
            merged.remove_unreferenced_resources()
            merged.save(pdf_path,
                        compress_streams=True,
                        recompress_flate=True,
                        object_stream_mode=pikepdf.ObjectStreamMode.generate)
        finally:
            for part in parts:
                part.close()
        raw_mb = sum(os.path.getsize(p) for p in part_pdfs) / 1e6
        print(f"[pagebreaks] PDF compressed {raw_mb:.1f} MB -> "
              f"{os.path.getsize(pdf_path)/1e6:.1f} MB")
    finally:
        for p in part_pdfs:
            try: os.remove(p)
            except OSError: pass


def export_to_pdf(xlsx_path, pdf_path, progress_callback=None):
    """
    Convert an .xlsx file to PDF using Excel via COM.

    All sheets except 'Index' are included. The xlsx is opened read-only
    and never modified. Raises RuntimeError if the PDF is not produced.
    """
    import win32com.client
    import pythoncom
    import tempfile

    def _progress(msg):
        print(f"[pagebreaks] {msg}")
        if progress_callback:
            progress_callback(msg)

    xlsx_path = os.path.normpath(os.path.abspath(xlsx_path))
    pdf_path  = os.path.normpath(os.path.abspath(pdf_path))

    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Excel file not found: {xlsx_path}")

    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    if os.path.exists(pdf_path):
        try: os.remove(pdf_path)
        except PermissionError:
            raise RuntimeError(f"PDF is open in another program: {pdf_path}")

    # Excel writes the raw (tagged, ~5x larger) PDF to the local temp dir so
    # the oversized intermediate never lands in a OneDrive-synced folder;
    # only the compressed copy is written to pdf_path.
    raw_pdf = os.path.join(
        tempfile.gettempdir(),
        f"~rate_pages_export_{os.getpid()}.pdf",
    )
    if os.path.exists(raw_pdf):
        try: os.remove(raw_pdf)
        except OSError: pass

    # Opening the workbook from a OneDrive-synced folder stalls on hydration
    # and sync locks; give Excel a local temp copy to read instead.
    src_for_excel = xlsx_path
    local_xlsx = None
    onedrive = os.environ.get("OneDrive", "")
    if onedrive and xlsx_path.lower().startswith(os.path.normpath(onedrive).lower()):
        local_xlsx = os.path.join(
            tempfile.gettempdir(),
            f"~rate_pages_src_{os.getpid()}.xlsx",
        )
        t0 = time.perf_counter()
        shutil.copy2(xlsx_path, local_xlsx)
        src_for_excel = local_xlsx
        print(f"[pagebreaks] copied source out of OneDrive in {time.perf_counter()-t0:0.1f}s")

    _kill_excel_instances()

    # Workbooks with a huge sheet (e.g. BOP's 82k-row TRDEF) render ~2.6x
    # faster split across several Excel instances; anything going wrong here
    # falls through to the normal single-instance export below.
    try:
        plan = _plan_parallel_export(src_for_excel)
    except Exception as exc:
        print(f"[pagebreaks] parallel planning skipped ({exc})")
        plan = None
    if plan:
        t_export = time.perf_counter()
        try:
            _export_pdf_parallel(src_for_excel, plan, pdf_path, _progress)
            print(f"[pagebreaks] parallel export took "
                  f"{time.perf_counter()-t_export:0.1f}s")
            if local_xlsx is not None:
                try: os.remove(local_xlsx)
                except OSError: pass
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                return pdf_path
            raise RuntimeError("merged PDF missing or empty")
        except Exception as exc:
            print(f"[pagebreaks] parallel export failed ({exc}) — "
                  f"falling back to single-instance export")
            try: os.remove(pdf_path)
            except OSError: pass

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    t_export = time.perf_counter()
    try:
        _progress("PDF — launching Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        _progress("PDF — opening workbook...")
        workbook = excel.Workbooks.Open(src_for_excel, ReadOnly=True, UpdateLinks=0)

        # Hide Index from the PDF without touching the source file.
        for sheet in workbook.Sheets:
            if sheet.Name == "Index":
                sheet.Visible = 0  # xlSheetHidden
            else:
                # Make sure the first non-Index sheet is active so Excel
                # picks a sensible page-1.
                try: sheet.Activate()
                except Exception: pass
                break

        _progress("PDF — Excel is rendering pages (the long step)...")
        workbook.ExportAsFixedFormat(
            Type=_XL_TYPE_PDF,
            Filename=raw_pdf,
            Quality=_XL_QUALITY_STD,
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        del workbook, excel
        pythoncom.CoUninitialize()
        if local_xlsx is not None:
            try: os.remove(local_xlsx)
            except OSError: pass
    print(f"[pagebreaks] Excel export took {time.perf_counter()-t_export:0.1f}s")

    # Verify Excel produced the raw PDF, then shrink it into pdf_path.
    for _ in range(10):
        if os.path.exists(raw_pdf) and os.path.getsize(raw_pdf) > 0:
            break
        time.sleep(0.2)
    else:
        raise RuntimeError(f"PDF was not created: {raw_pdf}")

    _progress("PDF — compressing...")
    t_comp = time.perf_counter()
    _compress_pdf(raw_pdf, pdf_path)
    print(f"[pagebreaks] compression took {time.perf_counter()-t_comp:0.1f}s")

    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
        return pdf_path
    raise RuntimeError(f"PDF was not created: {pdf_path}")


# Example usage:
# process_pagebreaks(r"C:\path\to\workbook.xlsx", "ignored.pdf")
# export_to_pdf(r"C:\path\to\workbook.xlsx", r"C:\path\to\workbook.pdf")
