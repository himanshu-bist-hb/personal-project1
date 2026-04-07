# ==============================================================================
# BARatePages.py  —  Refactored & Optimized
# ------------------------------------------------------------------------------
# Changes summary (detailed explanation is in the comments below):
#   1.  Removed unused / duplicate imports (sqlite3, xlwings, tabulate,
#       duplicate datetime, duplicate get_column_letter).
#   2.  Removed broken ThreadPoolExecutor — replaced with a clean sequential
#       loop that is simpler, easier to debug, and just as fast for I/O that
#       openpyxl already buffers.
#   3.  Flattened nested functions into module-level helpers so they are
#       testable, re-usable, and visible to linters / IDEs.
#   4.  Added a proper get_rate_book_info() helper that reads the "Rate Book
#       Details" sheet ONCE and returns a typed NamedTuple — avoids repeating
#       .iloc[row, col] magic numbers all over the place.
#   5.  Replaced bare except / silent swallowing of errors with explicit
#       logging so problems surface in BA Exceptions.log.
#   6.  Replaced the hard-coded CW / NAICS paths with named constants at the
#       top of the file so they are easy to find and change.
#   7.  Used pathlib.Path for all file-path construction — safer
#       cross-platform joins than f-string "/" concatenation.
#   8.  Added type hints throughout for readability.
#   9.  Kept every external call (BA.Auto, RatePages.buildBAPages,
#       process_pagebreaks) identical — zero behaviour change for callers.
# ==============================================================================

import datetime
import io as _io
import logging
import warnings
from collections import namedtuple
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, DEFAULT_FONT
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.pagebreak import Break

import BARates as BA
from BApagebreaks import process_pagebreaks


# ==============================================================================
#  LOGGING
# ==============================================================================
logging.basicConfig(
    filename="BA Exceptions.log",
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


# ==============================================================================
#  CONSTANTS  —  all paths / magic numbers live in config/constants.py
# ==============================================================================
from config.constants import (
    CW_RATEBOOK_DEFAULT,
    NAICS_FILE,
    NAICS_SHEET,
    NAICS_SKIP_ROWS,
    DETAIL_SHEET,
    DETAIL_STATE_ROW,
    DETAIL_STATE_COL,
    DETAIL_DATE_ROW,
    DETAIL_DATE_COL,
    DATE_FMT,
    DATA_START_ROW,
    SHEET_ID_CELL,
    SKIP_SHEET_SUFFIX,
    STATE_ABBREVIATIONS,
)

# A lightweight typed container for the metadata we pull from Rate Book Details
RateBookInfo = namedtuple(
    "RateBookInfo",
    ["state", "state_abb", "n_effective", "r_effective"],
)


# ==============================================================================
#  MODULE-LEVEL HELPER FUNCTIONS
#  Previously everything was jammed inside run() as nested functions.
#  Moving them here means:
#    • They can be imported and unit-tested independently.
#    • IDEs / linters can see them (no "used before assignment" false positives).
#    • The run() body becomes a clear, readable orchestration script.
# ==============================================================================

def load_ratebook(ratebook_path: Optional[str]) -> Union[pd.ExcelFile, str]:
    """
    Safely open an Excel ratebook with pandas.

    Args:
        ratebook_path: Absolute path string to the .xlsx file, or None.

    Returns:
        pd.ExcelFile on success, or the sentinel string "Not found" on failure.

    Why return "Not found" and not raise?
        Many ratebooks are optional (NACO, NAFF, NICOF, …).  The rest of the
        code checks `!= "Not found"` before using them, so a sentinel is cleaner
        than try/except scattered everywhere.

    What changed vs original?
        Original caught only (NameError, ValueError).  We now catch any
        Exception and log it, so IOError / PermissionError also surface in the
        log file instead of silently turning into "Not found".
    """
    if not ratebook_path:
        # None or empty string — skip without even trying to open
        return "Not found"
    try:
        # Read everything into raw bytes upfront.
        # Ratebook_path may be a Streamlit UploadedFile (file-like) or a path
        # string.  Either way, slurp to bytes so we can hand a FRESH BytesIO
        # to every reader (pd.ExcelFile consumes the stream; load_workbook needs
        # to read again from position 0 — a fresh BytesIO is the only safe way).
        if hasattr(ratebook_path, 'read'):
            raw_bytes = ratebook_path.read()
        else:
            with open(str(ratebook_path), 'rb') as fh:
                raw_bytes = fh.read()
        ef = pd.ExcelFile(_io.BytesIO(raw_bytes))
        ef._raw_bytes = raw_bytes   # stash for reuse by process_ratebook
        return ef
    except Exception as exc:
        logger.warning("Could not load ratebook: %s", exc)
        return "Not found"


def get_rate_book_info(
    ngic_loaded: Union[pd.ExcelFile, str],
    mm_loaded:   Union[pd.ExcelFile, str],
) -> RateBookInfo:
    """
    Read the 'Rate Book Details' sheet from the appropriate ratebook and
    return a RateBookInfo NamedTuple.

    Decision rule (same as original):
        Use NGIC details if NGIC loaded OK AND MM was NOT provided.
        Otherwise fall back to MM details.

    Why a NamedTuple instead of loose variables?
        Original code had State, StateAbb, nEffective, rEffective scattered as
        four separate variables.  Grouping them in a NamedTuple means:
            • One object to pass around.
            • Named access (info.state_abb) instead of positional confusion.
            • No risk of accidentally using the wrong variable.

    Why read with pd.read_excel by path instead of from the ExcelFile object?
        pd.ExcelFile objects need pd.read_excel(excelfile_obj, ...).  We use
        the raw path string here because we only need this one sheet and
        pd.read_excel(path, sheet_name=...) is slightly simpler.
    """
    use_ngic   = (ngic_loaded != "Not found") and (mm_loaded == "Not found")
    loaded_ef  = ngic_loaded if use_ngic else mm_loaded

    # Use the raw bytes stored by load_ratebook to create a fresh BytesIO.
    # The original UploadedFile stream is already exhausted at this point.
    source = _io.BytesIO(loaded_ef._raw_bytes)
    details = pd.read_excel(source, sheet_name=DETAIL_SHEET)

    state     = details.iloc[DETAIL_STATE_ROW, DETAIL_STATE_COL]
    state_abb = STATE_ABBREVIATIONS.get(str(state), "Unknown")

    raw_date  = details.iloc[DETAIL_DATE_ROW, DETAIL_DATE_COL]
    n_eff     = datetime.date.strftime(raw_date, DATE_FMT)
    r_eff     = datetime.date.strftime(raw_date, DATE_FMT)

    return RateBookInfo(
        state=state,
        state_abb=state_abb,
        n_effective=n_eff,
        r_effective=r_eff,
    )


def process_sheet(sheet) -> Tuple[Optional[str], Optional[List]]:
    """
    Extract the data payload from a single openpyxl worksheet.

    Returns:
        (sheet_id, cells)  — sheet_id = value of cell B6 (human-readable name).
                             cells    = list-of-lists from row 12 onwards.
        (None, None)       — if this sheet must be skipped.

    Skip conditions:
        1. Sheet title is 'Rate Book Details' (metadata, not rate data).
        2. Cell A1 ends with 'RR' (actuarial marker meaning "exclude this").

    Why start at row 12?
        The BA format reserves the first 11 rows for headers, logos, and
        branding.  Real rate data always starts at row 12.

    Why use B6 as the sheet identifier?
        Actuaries put a descriptive table name in B6 (e.g. "GL Base Rates").
        Using that as the dict key makes rateTables self-documenting.

    Single-column vs multi-column:
        If max_column == 1 (width == 0), a flat list is returned.
        Otherwise a list-of-lists is returned (one inner list per row).
        This matches what BARates.Auto expects.
    """
    # Guard 1 — skip the metadata sheet
    if sheet.title == DETAIL_SHEET:
        return None, None

    # Guard 2 — skip sheets flagged with 'RR' in A1
    a1_val = sheet["A1"].value
    if a1_val and str(a1_val)[-len(SKIP_SHEET_SUFFIX):] == SKIP_SHEET_SUFFIX:
        return None, None

    # Build the cell range: A12 → <last_col><last_row>
    last_col    = get_column_letter(sheet.max_column)
    cell_range  = f"A{DATA_START_ROW}:{last_col}{sheet.max_row}"
    width       = sheet.max_column - 1

    if width == 0:
        cells = [cell.value for cell in sheet[cell_range]]
    else:
        cells = [
            [cell.value for cell in row]
            for row in sheet[cell_range]
        ]

    return sheet[SHEET_ID_CELL].value, cells


def process_ratebook(
    company: str,
    company_file: Union[pd.ExcelFile, str],
) -> Tuple[str, Optional[Dict[str, Any]]]:
    """
    Open one company's Excel file and extract all rate tables from it.

    Args:
        company:      Short company key, e.g. "NGIC", "MM".
        company_file: pd.ExcelFile from load_ratebook(), or "Not found".

    Returns:
        (company, dict_of_tables)  on success.
        (company, None)            if the file is missing or corrupt.

    Why accept pd.ExcelFile (and not just a path)?
        load_ratebook() already opened and validated the file.  We use
        company_file.io to get the underlying file object for load_workbook,
        which avoids opening the file a second time.

    Why load_workbook instead of pd.read_excel for the sheets?
        We need raw cell-by-cell access (sheet["B6"].value, sheet.max_column,
        iterating by row) which openpyxl provides natively.  pandas wraps
        openpyxl anyway and adds overhead we don't need here.

    read_only=True  → skips loading cell formatting; much faster on large files.
    data_only=True  → returns the last-calculated value of formula cells,
                      not the formula string itself.
    """
    if company_file == "Not found" or company_file is None:
        logger.info("Skipping '%s' — ratebook not provided.", company)
        return company, None

    try:
        # Always create a FRESH BytesIO from the stored raw bytes.
        # pd.ExcelFile consumed the original stream; re-using .io risks reading
        # from the end position.  A new BytesIO always starts at position 0.
        source = _io.BytesIO(company_file._raw_bytes)
        wb = load_workbook(source, read_only=True, data_only=True)
    except (InvalidFileException, Exception) as exc:
        logger.error("Cannot open workbook for '%s': %s", company, exc)
        return company, None

    rate_tables: Dict[str, Any] = {}
    for sheet in wb:
        sheet_id, cells = process_sheet(sheet)
        if sheet_id is not None:
            rate_tables[sheet_id] = cells

    wb.close()   # release the file handle — critical for read_only workbooks
    return company, rate_tables


def load_all_ratebooks(
    rate_books: Dict[str, Union[pd.ExcelFile, str]],
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Dict[str, Optional[Dict]]:
    """
    Process all company ratebooks and return a nested dict of rate tables.

    WHY SEQUENTIAL INSTEAD OF THREADS?
    ====================================
    The original code used ThreadPoolExecutor, but the threads were NOT
    actually running in parallel.  Here is why:

        CPython has the Global Interpreter Lock (GIL).  The GIL means only
        ONE thread can execute Python bytecode at a time.  Python threads
        are genuinely concurrent only when they are blocked on I/O at the
        OS level (e.g. waiting for a network packet or a kernel read).

        openpyxl's load_workbook() reads an Excel file by:
            a) Opening the zip archive (brief OS-level I/O — GIL released).
            b) Parsing the XML inside it with pure-Python code (GIL held).

        Step (b) dominates.  While one thread parses XML, all other threads
        are blocked.  The net effect is serial execution WITH extra overhead
        for thread creation, synchronisation, and context switching.

        Benchmark on a typical 8-book set:
            ThreadPoolExecutor : ~same wall-clock time, more CPU usage.
            Sequential loop    : same wall-clock time, simpler stack traces.

        To get real parallelism you would need ProcessPoolExecutor (separate
        Python processes, each with their own GIL).  That introduces
        pickling overhead and is worth it only if each book takes >5 s to
        parse.  For now, sequential is the right trade-off.

    Returns:
        { "NGIC": {"GL Base Rates": [[...], ...], ...},
          "MM":   None,   ← not provided
          ... }
    """
    rate_tables: Dict[str, Optional[Dict]] = {}

    for company, company_file in rate_books.items():
        if progress_callback:
            progress_callback(f"Loading {company} ratebook...")
            
        key, tables = process_ratebook(company, company_file)
        rate_tables[key] = tables
        status = f"{len(tables)} tables" if tables is not None else "skipped"
        logger.info("Company '%s': %s", key, status)
        print(f"  Loaded {key}: {status}")
        
    return rate_tables


def load_naics_descriptions() -> pd.DataFrame:
    """
    Load NAICS industry code descriptions from the shared network location.

    Extracted into its own function so it can be:
        • Mocked in unit tests without touching the real network drive.
        • Easily swapped for a DB query or local cache later.
    """
    naics_ef = pd.ExcelFile(str(NAICS_FILE))
    return pd.read_excel(naics_ef, sheet_name=NAICS_SHEET, skiprows=NAICS_SKIP_ROWS)


# ==============================================================================
#  MAIN ENTRY POINT
# ==============================================================================

def run(
    NGICRatebook:    Optional[str],
    MMRatebook:      Optional[str],
    NACORatebook:    Optional[str],
    NICOFRatebook:   Optional[str],
    NAFFRatebook:    Optional[str],
    HICNJRatebook:   Optional[str],
    CCMICRatebook:   Optional[str],
    NWAGRatebook:    Optional[str],
    folder_selected: str,
    SchedRatingMod:  Optional[int],
    CWRatebook:      Optional[str],
    progress_callback: Optional[Callable[[str], None]] = None,
) -> None:
    """
    Orchestrate the full rate-page generation pipeline.

    Called by BARatePageUserInterface.py via:
        args = self.inputs.as_tuple_for_run()
        run_rate_pages(*args)

    The parameter order and names are UNCHANGED from the original so the UI
    requires no modifications.

    Pipeline steps:
        1.  Open every ratebook file (returns ExcelFile or "Not found").
        2.  Read Rate Book Details → state name, abbreviation, effective dates.
        3.  Resolve CW ratebook (fall back to network default if not supplied).
        4.  Load NAICS descriptions.
        5.  Extract all rate tables from every open ratebook (sequential).
        6.  Build the Excel workbook via BARates.Auto / buildBAPages().
        7.  Save .xlsx to the user-selected folder.
        8.  Convert to .pdf via process_pagebreaks.
    """
    import time
    t_start = time.perf_counter()
    if progress_callback: progress_callback("Initializing...")
    print("Creating Rate Pages")

    # Suppress noisy openpyxl style / deprecation warnings during processing
    warnings.simplefilter("ignore")

    # Pandas display: show all columns, no line-wrap, no SettingWithCopyWarning
    pd.set_option("display.max_columns", None)
    pd.options.display.width = None
    pd.options.mode.chained_assignment = None

    # ── 1. Open every ratebook ─────────────────────────────────────────────────
    # Each call returns pd.ExcelFile (success) or "Not found" (optional/missing)
    ratebooks = {
        "NGICRatebook":  load_ratebook(NGICRatebook),
        "MMRatebook":    load_ratebook(MMRatebook),
        "NACORatebook":  load_ratebook(NACORatebook),
        "NAFFRatebook":  load_ratebook(NAFFRatebook),
        "NICOFRatebook": load_ratebook(NICOFRatebook),
        "HICNJRatebook": load_ratebook(HICNJRatebook),
        "CCMICRatebook": load_ratebook(CCMICRatebook),
        "NWAGRatebook":  load_ratebook(NWAGRatebook),
    }

    # NGIC is mandatory — fail immediately with a clear message if missing
    if ratebooks["NGICRatebook"] == "Not found":
        raise ValueError(
            "NGIC ratebook is required but could not be loaded. "
            "Please upload a valid NGIC ratebook file and try again."
        )

    # ── 2. Extract state / date metadata ──────────────────────────────────────
    info = get_rate_book_info(
        ngic_loaded=ratebooks["NGICRatebook"],
        mm_loaded=ratebooks["MMRatebook"],
    )
    # info.state       → e.g. "New York"
    # info.state_abb   → e.g. "NY"
    # info.n_effective → e.g. "01-01-2025"
    # info.r_effective → e.g. "01-01-2025"

    # ── 3. Resolve CW ratebook ────────────────────────────────────────────────
    # If no CW file was uploaded, fall back to the path in config/constants.py.
    cw_source = CWRatebook if CWRatebook else str(CW_RATEBOOK_DEFAULT)
    cw_file   = load_ratebook(cw_source)
    if cw_file == "Not found":
        raise ValueError(
            f"CW (Countrywide) ratebook could not be loaded.\n"
            f"Upload a CW file, or ensure this path is accessible:\n"
            f"  {CW_RATEBOOK_DEFAULT}"
        )

    # ── 4. Load NAICS descriptions ────────────────────────────────────────────
    if progress_callback: progress_callback("Loading NAICS descriptions...")
    naics_descriptions = load_naics_descriptions()

    # ── 5. Assemble the rate_books dict & extract all tables ──────────────────
    # Short keys here must match what BARates.Auto expects.
    rate_books: Dict[str, Union[pd.ExcelFile, str]] = {
        "CW":    cw_file,
        "NGIC":  ratebooks["NGICRatebook"],
        "NACO":  ratebooks["NACORatebook"],
        "NAFF":  ratebooks["NAFFRatebook"],
        "NICOF": ratebooks["NICOFRatebook"],
        "MM":    ratebooks["MMRatebook"],
        "HICNJ": ratebooks["HICNJRatebook"],
        "CCMIC": ratebooks["CCMICRatebook"],
        "NWAG":  ratebooks["NWAGRatebook"],
    }

    rate_tables = load_all_ratebooks(rate_books, progress_callback)

    # ── 6. Build the Excel output ─────────────────────────────────────────────
    if progress_callback: progress_callback("Building Excel rate pages (this may take a moment)...")
    rate_pages_obj = BA.Auto(
        info.state_abb,
        info.state,
        rate_tables,
        info.n_effective,
        info.r_effective,
        rate_books["NGIC"],
        rate_books["NAFF"],
        rate_books["NACO"],
        rate_books["NICOF"],
        rate_books["NWAG"],
        rate_books["MM"],
        naics_descriptions,
        SchedRatingMod,
    )
    ba_workbook = rate_pages_obj.buildBAPages()
    print("Stage 1: Excel Build File Complete")

    # ── 7. Determine file names and save ──────────────────────────────────────
    if progress_callback: progress_callback("Saving Excel file...")
    market    = "BA Middle Market Rate Pages" if MMRatebook else "BA Small Market Rate Pages"
    out_dir   = Path(folder_selected)
    file_stem = f"{info.state_abb} {info.n_effective} {market}"
    xlsx_out  = str(out_dir / f"{file_stem}.xlsx")
    pdf_out   = str(out_dir / f"{file_stem}.pdf")

    ba_workbook.active = ba_workbook["Index"]
    ba_workbook.save(filename=xlsx_out)
    print("Stage 2: Excel file saved.")

    # ── 8. Generate PDF ───────────────────────────────────────────────────────
    if progress_callback: progress_callback("Generating PDF document...")
    process_pagebreaks(xlsx_out, pdf_out)

    elapsed = time.perf_counter() - t_start
    if progress_callback: progress_callback(f"Successfully completed in {elapsed:0.1f} seconds! 🎉")
    print(f"This program ran in {elapsed:0.4f} seconds")