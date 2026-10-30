# ==============================================================================
# CPRatePages.py  —  Commercial Property orchestration
# ------------------------------------------------------------------------------
# Same shape as BA/BARatePages.py and BOP/BOPRatePages.py:
#     run()               ratebooks (uploaded bytes) -> saved SM / MM .xlsx files
#     generate_pdf_only() .xlsx -> PDF through Excel's own print engine
# The table/sheet logic is in CPRates.py; PDF export is the shared exporter in
# BA/BApagebreaks.py (not BA-specific).
# ==============================================================================

import io as _io
import logging
import os
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

from BA.BApagebreaks import _sanitize_xlsx, export_to_pdf
from config.constants import CP_CW_RATEBOOK_DEFAULT
from . import CPRates
from .CPRates import MarketOptions, NOT_FOUND

logging.basicConfig(
    filename="CP Exceptions.log",
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# Upload slots the engine reads (CW is passed separately — it has a network default).
RATEBOOK_KEYS = ("NGIC", "NACO", "NAFF", "NICOF", "MM", "SM", "ATA", "HEX")
MARKET_NAMES = {"SM": "Small Market", "MM": "Middle Market"}


def load_ratebook(source) -> Any:
    """Open an uploaded file (bytes, file-like or path) as a pd.ExcelFile; "Not found" if absent/unreadable."""
    if source is None or (isinstance(source, str) and not source):
        return NOT_FOUND
    try:
        if isinstance(source, (bytes, bytearray)):
            return pd.ExcelFile(_io.BytesIO(source))
        if hasattr(source, "read"):
            return pd.ExcelFile(_io.BytesIO(source.read()))
        return pd.ExcelFile(str(source))
    except Exception as exc:
        logger.warning("Could not load ratebook: %s", exc)
        return NOT_FOUND


def _check_same_state(ratebooks: Dict[str, Any]) -> None:
    """Company ratebooks from different states would silently produce nonsense pages."""
    seen = {}
    for key in ("NGIC", "NACO", "NAFF", "NICOF", "MM", "SM", "ATA"):
        rb = ratebooks.get(key, NOT_FOUND)
        if isinstance(rb, str) or "Rate Book Details" not in rb.sheet_names:
            continue
        try:
            seen[key] = str(pd.read_excel(rb, sheet_name="Rate Book Details").iloc[3, 4])
        except Exception:
            continue
    if len(set(seen.values())) > 1:
        detail = ", ".join(f"{k}: {v}" for k, v in seen.items())
        raise ValueError(f"The uploaded ratebooks are not all for the same state ({detail}).")


def validate_inputs(ratebooks: Dict[str, Any], sm: MarketOptions, mm: MarketOptions) -> None:
    """Fail early, in plain words, instead of with a NameError 2,000 lines into the build."""
    if not (sm.enabled or mm.enabled):
        raise ValueError("Select at least one market (Small Market or Middle Market).")
    state_level, _ = CPRates._resolve_sources(ratebooks)
    if sm.enabled and state_level == NOT_FOUND:
        raise ValueError("Small Market pages need an NGIC, Small Market or Applies to All ratebook.")
    if mm.enabled and ratebooks.get("MM", NOT_FOUND) == NOT_FOUND:
        raise ValueError("Middle Market pages need the MM ratebook.")
    if (sm.needs_hex or mm.needs_hex) and ratebooks.get("HEX", NOT_FOUND) == NOT_FOUND:
        raise ValueError("Territory Adjustment pages need the Hex ratebook.")
    if mm.enabled and mm.segmentation_removal and not mm.full_manual:
        logger.info("Segmentation Removal only changes the MM Full Rate Manual pages.")
    _check_same_state(ratebooks)


def run(
    folder_selected: str,
    small_market: Optional[MarketOptions] = None,
    middle_market: Optional[MarketOptions] = None,
    NGICRatebook=None, NACORatebook=None, NAFFRatebook=None, NICOFRatebook=None,
    MMRatebook=None, SMRatebook=None, ATARatebook=None, HexRatebook=None,
    CWRatebook=None, TerrCoordsRatebook=None,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> List[Dict[str, str]]:
    """
    Build the selected market workbooks and save them to folder_selected.

    Ratebook arguments are uploaded bytes (or paths); anything omitted is "not found".
    Returns one {"market", "xlsx", "pdf"} entry per market built — "pdf" is where
    generate_pdf_only() should write it.
    """
    t0 = time.perf_counter()
    sm = small_market or MarketOptions()
    mm = middle_market or MarketOptions()
    say = progress_callback or (lambda _msg: None)

    say("Opening ratebooks...")
    ratebooks = {
        "NGIC": load_ratebook(NGICRatebook), "NACO": load_ratebook(NACORatebook),
        "NAFF": load_ratebook(NAFFRatebook), "NICOF": load_ratebook(NICOFRatebook),
        "MM": load_ratebook(MMRatebook), "SM": load_ratebook(SMRatebook),
        "ATA": load_ratebook(ATARatebook), "HEX": load_ratebook(HexRatebook),
    }
    validate_inputs(ratebooks, sm, mm)
    info = CPRates.get_rate_book_info(ratebooks)
    state, state_abb, effective = info

    say("Opening Countrywide ratebook...")
    cw = load_ratebook(CWRatebook if CWRatebook else str(CP_CW_RATEBOOK_DEFAULT))
    if cw == NOT_FOUND:
        raise ValueError(f"CW ratebook could not be loaded: {CP_CW_RATEBOOK_DEFAULT}")
    coords = load_ratebook(TerrCoordsRatebook)

    books = CPRates.build_workbooks(
        ratebooks, cw, info, sm, mm,
        terr_coords=None if coords == NOT_FOUND else coords,
        progress=say,
    )

    out_dir = Path(folder_selected)
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    for market in ("SM", "MM"):
        book = books.get(market)
        if book is None:
            continue
        stem = f"{state_abb} {effective} CP {MARKET_NAMES[market]} Rate Pages"
        xlsx = str(out_dir / f"{stem}.xlsx")
        say(f"Saving {MARKET_NAMES[market]} workbook...")
        try:
            book.save(xlsx)
        except PermissionError:
            raise PermissionError(f"Close '{Path(xlsx).name}' in Excel and try again.")
        _sanitize_xlsx(xlsx)
        results.append({"market": market, "xlsx": xlsx, "pdf": str(out_dir / f"{stem}.pdf")})

    say(f"Successfully completed in {time.perf_counter() - t0:0.1f} seconds! 🎉")
    return results


def generate_pdf_only(xlsx_path: str, pdf_path: str,
                      progress_callback: Optional[Callable[[str], None]] = None) -> str:
    """Convert a saved rate-page workbook to PDF (honours every print setting in the file)."""
    t0 = time.perf_counter()
    out = export_to_pdf(xlsx_path, pdf_path, progress_callback=progress_callback)
    if not (os.path.exists(out) and os.path.getsize(out) > 0):
        raise RuntimeError(f"PDF was not created at {out}")
    if progress_callback:
        progress_callback(f"PDF created in {time.perf_counter() - t0:0.1f}s — {os.path.basename(out)} 🎉")
    return out
