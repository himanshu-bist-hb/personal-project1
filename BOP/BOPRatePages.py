"""
BOPRatePages.py
================
Orchestration entry point for BOP (Business Owners Policy) — All Programs.

Mirrors BA/BARatePages.py:run() closely: same ratebook-loading helpers
(reused directly from BA.BARatePages, since they're already company-agnostic),
same "load -> extract metadata -> build workbook -> save -> page-break" flow,
same (xlsx_out, pdf_out) return shape so app.py's wiring is a straight copy
of the Farm Auto block.
"""

import time
import warnings
from pathlib import Path
from typing import Callable, Dict, Optional, Tuple, Union

import pandas as pd

from BA.BARatePages import load_ratebook, get_rate_book_info, load_all_ratebooks
from config.constants import BOP_TERRITORY_DEFS_PATH, BOP_CW_RATEBOOK_DEFAULT
from . import AllProgramsPage
from . import AllProgramsPageCurrent
from .bop_config import load_bop_config
from .BOPpagebreaks import process_pagebreaks, export_to_pdf

# "2.0" -> AllProgramsPage.AllPrograms (needs Territory Defs)
# "pre2.0" -> AllProgramsPageCurrent.AllPrograms (no Territory Defs at all —
#   that program and its build/format methods predate the territory tables)
VALID_VERSIONS = ("2.0", "pre2.0")


def load_territory_defs(state_abb: str) -> pd.DataFrame:
    """
    Load the Territory Definitions workbook (network drive) and return the
    sheet for the given state. Required for the All Programs page.
    """
    territory_ef = pd.ExcelFile(str(BOP_TERRITORY_DEFS_PATH))
    return pd.read_excel(territory_ef, sheet_name=state_abb)


def run(
    NGICRatebook: Optional[str],
    folder_selected: str,
    CWRatebook: Optional[str] = None,
    MMRatebook: Optional[str] = None,
    NACORatebook: Optional[str] = None,
    NAFFRatebook: Optional[str] = None,
    NICOFRatebook: Optional[str] = None,
    HICNJRatebook: Optional[str] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    skip_pdf: bool = True,
    version: str = "2.0",
) -> Tuple[str, str]:
    """
    Orchestrate the BOP All Programs rate-page generation pipeline.

    Args:
        version: "2.0" (default) or "pre2.0" — selects which generation of
            the All Programs rating logic and page layout to build.

    Returns:
        (xlsx_out, pdf_out) paths.
    """
    if version not in VALID_VERSIONS:
        raise ValueError(f"version must be one of {VALID_VERSIONS}, got {version!r}")

    t_start = time.perf_counter()
    if progress_callback: progress_callback("Initializing...")
    print(f"Creating BOP All Programs Rate Pages ({version})")

    warnings.simplefilter("ignore")
    pd.set_option("display.max_columns", None)
    pd.options.display.width = None

    # ── 1. Open every ratebook ─────────────────────────────────────────────
    ratebooks = {
        "NGIC":  load_ratebook(NGICRatebook),
        "MM":    load_ratebook(MMRatebook),
        "NACO":  load_ratebook(NACORatebook),
        "NAFF":  load_ratebook(NAFFRatebook),
        "NICOF": load_ratebook(NICOFRatebook),
        "HICNJ": load_ratebook(HICNJRatebook),
    }
    if ratebooks["NGIC"] == "Not found":
        raise ValueError("NGIC ratebook is required.")

    # CW is optional — fall back to the static network copy when the user
    # doesn't upload their own, same pattern as Business Auto's CW handling.
    cw_source = CWRatebook if CWRatebook else str(BOP_CW_RATEBOOK_DEFAULT)
    cw_file = load_ratebook(cw_source)
    if cw_file == "Not found":
        raise ValueError(f"CW ratebook could not be loaded: {BOP_CW_RATEBOOK_DEFAULT}")

    # ── 2. Extract state / date metadata (same 'Rate Book Details' layout BA uses) ──
    info = get_rate_book_info(ngic_loaded=ratebooks["NGIC"], mm_loaded=ratebooks["MM"])

    # ── 3. Load Territory Definitions (2.0 only — pre2.0 never used them) ──
    territory_defs_by_st = None
    if version == "2.0":
        if progress_callback: progress_callback("Loading Territory Definitions...")
        territory_defs_by_st = load_territory_defs(info.state_abb)

    # ── 4. Load config-driven rating lookup tables ─────────────────────────
    cfg = load_bop_config()
    if info.state_abb not in cfg.perils_by_state:
        raise ValueError(
            f"No 'Perils By State' entry for '{info.state_abb}' in BOP Input File.xlsx — "
            "add a row there before generating this state's rate pages."
        )
    perils = cfg.perils_by_state[info.state_abb]

    # ── 5. Assemble the rate_books dict & extract all tables ───────────────
    rate_books: Dict[str, Union[pd.ExcelFile, str]] = {
        "CW":    cw_file,
        "NGIC":  ratebooks["NGIC"],
        "NACO":  ratebooks["NACO"],
        "NAFF":  ratebooks["NAFF"],
        "NICOF": ratebooks["NICOF"],
        "HICNJ": ratebooks["HICNJ"],
        "MM":    ratebooks["MM"],
    }
    rate_tables_raw = load_all_ratebooks(rate_books, progress_callback)
    # Drop companies that were not provided so AllPrograms.buildDataFrame's
    # "'NACO' in self.rateTables.keys()" optional-company checks behave
    # correctly (a present-but-None entry would otherwise crash on
    # None.keys()).
    rate_tables = {k: v for k, v in rate_tables_raw.items() if v is not None}

    # ── 6. Build the Excel output ───────────────────────────────────────────
    if progress_callback: progress_callback("Building Excel rate pages...")
    if version == "2.0":
        rate_pages_obj = AllProgramsPage.AllPrograms(
            info.state_abb, rate_tables, perils,
            cfg.peril_conversions, cfg.protection_class_conversions,
            cfg.building_codes_by_state,
            info.n_effective, info.r_effective, territory_defs_by_st,
        )
    else:
        rate_pages_obj = AllProgramsPageCurrent.AllPrograms(
            info.state_abb, rate_tables, perils,
            cfg.peril_conversions, cfg.protection_class_conversions,
            cfg.building_codes_by_state,
            info.n_effective, info.r_effective,
        )
    bop_workbook = rate_pages_obj.buildAllProgramsPage()

    # ── 7. Determine file names and save ────────────────────────────────────
    if progress_callback: progress_callback("Saving Excel file...")
    out_dir     = Path(folder_selected)
    version_tag = "" if version == "2.0" else " (Pre 2.0)"
    file_stem   = f"{info.state_abb} {info.n_effective} BOP All Programs Rate Pages{version_tag}"
    xlsx_out  = str(out_dir / f"{file_stem}.xlsx")
    pdf_out   = str(out_dir / f"{file_stem}.pdf")

    bop_workbook.active = bop_workbook["Index"]
    bop_workbook.save(filename=xlsx_out)
    print("Stage 2: Excel file saved.")

    # ── 8. Apply page breaks / print settings ───────────────────────────────
    if progress_callback: progress_callback("Applying page breaks...")
    process_pagebreaks(xlsx_out, pdf_out)

    elapsed = time.perf_counter() - t_start
    if progress_callback: progress_callback(f"Successfully completed in {elapsed:0.1f} seconds! 🎉")
    print(f"This program ran in {elapsed:0.4f} seconds")

    return xlsx_out, pdf_out


def generate_pdf_only(xlsx_path: str, pdf_path: str, progress_callback: Optional[Callable[[str], None]] = None) -> str:
    """
    Convert an existing rate-pages .xlsx into a PDF using Excel's own print
    engine, so every page-break / fit-to-page / print-area setting already
    written to the workbook is honored.
    """
    import os
    if progress_callback: progress_callback("Launching Excel...")
    t0 = time.perf_counter()
    out = export_to_pdf(xlsx_path, pdf_path)
    if not (os.path.exists(out) and os.path.getsize(out) > 0):
        raise RuntimeError(f"PDF was not created at {out}")
    elapsed = time.perf_counter() - t0
    if progress_callback:
        progress_callback(f"PDF created in {elapsed:0.1f}s — {os.path.basename(out)} 🎉")
    print(f"[BOPRatePages] PDF generated: {out}")
    return out
