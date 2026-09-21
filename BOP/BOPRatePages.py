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
from typing import Callable, Dict, List, Optional, Sequence, Tuple, Union

import pandas as pd

from BA.BARatePages import load_ratebook, get_rate_book_info, load_all_ratebooks
from config.constants import BOP_TERRITORY_DEFS_PATH, BOP_CW_RATEBOOK_DEFAULT, BOP_EQ_TERRITORY_DEFS_DIR
from . import AllPerilPage
from . import AllPerilPageCurrent
from . import AllProgramsPage
from . import AllProgramsPageCurrent
from . import HabPage
from . import HabPageCurrent
from . import HabPageAppetite
from . import AutoServicePage
from . import AutoServicePageCurrent
from . import RetailPage
from . import RetailPageCurrent
from . import ServicePage
from . import ServicePageCurrent
from . import OfficePage
from . import OfficePageCurrent
from . import WholesalePage
from . import WholesalePageCurrent
from . import FoodServicePage
from . import FoodServicePageCurrent
from . import OptionalCoveragesPage
from . import RatingPlansPage
from . import ClassModifierPage
from . import CommonRulesPage
from . import AdditionalRulesPage
from .bop_config import load_bop_config, resolve_default_version, resolve_default_appetite
from .BOPpagebreaks import (
    process_pagebreaks, export_to_pdf, export_single_sheet_pdf, split_pdf_by_size,
)

# "2.0" -> AllProgramsPage.AllPrograms (needs Territory Defs)
# "pre2.0" -> AllProgramsPageCurrent.AllPrograms (no Territory Defs at all —
#   that program and its build/format methods predate the territory tables)
#
# Appetite is NOT a version of its own — it's an add-on flag (`appetite`
# below) layered on top of whichever version ("2.0" or "pre2.0") is built.
# A program with Appetite-specific pages for a given base version gets an
# entry in APPETITE_CLASSES keyed by (program, that version), pointing at a
# subclass of that version's own page class (see HabPageAppetite for the
# pattern); a program/version pair with no entry there just builds the
# plain base-version class even when `appetite` is True, since Appetite has
# no changes for it yet.
VALID_VERSIONS = ("2.0", "pre2.0", "Default")

# (program, base_version) -> Appetite subclass of that version's page class.
# Add an entry here (and the matching *PageAppetite.py subclass) as more
# programs/versions get Appetite pages; anything not listed here falls back
# to the plain base-version class when `appetite=True` — see _version_cls().
APPETITE_CLASSES = {
    ("Hab", "2.0"): HabPageAppetite.Hab,
}


def _version_cls(version, appetite, cls_20, cls_pre20, program=None):
    """
    Resolve which page class to build for the given base `version`
    ("2.0"/"pre2.0"), applying the Appetite add-on: when `appetite` is True
    and (program, version) is listed in APPETITE_CLASSES, its Appetite
    subclass is used (built on top of that version's own class); otherwise
    the plain base-version class is used, since Appetite pages don't exist
    for every program/version yet.
    """
    base_cls = cls_20 if version == "2.0" else cls_pre20
    if appetite:
        return APPETITE_CLASSES.get((program, version), base_cls)
    return base_cls

# "All Programs" -> AllProgramsPage / AllProgramsPageCurrent (by-peril tables)
# "All Peril"    -> AllPerilPage / AllPerilPageCurrent (by-program tables,
#   "allperil" peril only; never needs the Territory Definitions workbook)
# "Hab" / "Auto Service" / "Retail" / "Service" / "Office" / "Wholesale" ->
#   the individual program pages. Like All Peril, none ever need the
#   Territory Definitions workbook — 2.0 versions don't print a Territory
#   Multiplier table at all (dropped when the All Programs Territory page
#   took over); pre2.0 versions build theirs straight from each ratebook's
#   own BP7_Peril_TerritorialFactor table.
# "Optional Coverages", "Rating Plans", "Class Modifier", "Common Rules" and
#   "Additional Rules" are the five programs with no "2.0"/"pre2.0" split at
#   all: there's only ever one OptionalCoveragesPage.py / RatingPlansPage.py
#   / ClassModifierPage.py / CommonRulesPage.py / AdditionalRulesPage.py (no
#   *Current variant), so each is built identically regardless of `version`.
#   Optional Coverages is also the only program needing the separate
#   Earthquake Territory Definitions file (see load_eq_territory_defs), not
#   the All Programs Territory Definitions workbook.
VALID_PROGRAMS = ("All Programs", "All Peril", "Hab", "Auto Service", "Retail", "Service", "Office", "Wholesale", "Food Service", "Optional Coverages", "Rating Plans", "Class Modifier", "Common Rules", "Additional Rules")

# The 2.0 "All Programs" workbook's last sheet — its 82k-row Territory
# Definitions table dominates PDF export time, so the main PDF export
# excludes it (see run()/generate_pdf_only's exclude_sheets) and it is
# generated separately/optionally via generate_territory_defs_pdf.
TERRITORY_DEFS_SHEET = "TRDEF"


def load_territory_defs(state_abb: str) -> pd.DataFrame:
    """
    Load the Territory Definitions workbook (network drive) and return the
    sheet for the given state. Required for the All Programs page.
    """
    territory_ef = pd.ExcelFile(str(BOP_TERRITORY_DEFS_PATH))
    return pd.read_excel(territory_ef, sheet_name=state_abb)


def load_eq_territory_defs(state_abb: str) -> pd.DataFrame:
    """
    Load the per-state Earthquake Territory Definitions file (network drive,
    tab-separated) — required only for Optional Coverages' "Earthquake
    Territory Definitions" table (OC Table C.4.E.3).
    """
    path = BOP_EQ_TERRITORY_DEFS_DIR / f"NWCE_{state_abb}_ZIP_1223.txt"
    return pd.read_csv(str(path), sep="\t", header=0)


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
    irpm_credit: float = 0.0,
    irpm_debit: float = 0.0,
    version: str = "Default",
    appetite: Optional[bool] = None,
    program: Union[str, Sequence[str]] = "All Programs",
) -> Tuple[Union[str, List[str]], Union[str, List[str]]]:
    """
    Orchestrate the BOP rate-page generation pipeline.

    Args:
        version: "Default" (default), "2.0" or "pre2.0" — selects which
            generation of the rating logic and page layout to build.
            "Default" looks the ratebook's state up in the "Version By
            State" tab of BOP Input File.xlsx and resolves to "2.0" or
            "pre2.0" from there (see resolve_default_version); it is
            resolved after the state is known, once ratebook metadata has
            been extracted below.
        appetite: whether to add each program's Appetite-only pages (see
            APPETITE_CLASSES) on top of the resolved `version` — a program
            with no Appetite pages for that version builds plain `version`
            regardless. When `version` is "Default", this is ignored and
            instead resolved per-state from the "Appetite" column of the
            same "Version By State" tab (see resolve_default_appetite).
            When `version` is explicitly "2.0"/"pre2.0", defaults to False
            if not given.
        program: which BOP program(s) to build — a single name ("All
            Programs" or "All Peril") or a list of names. The ratebooks are
            opened and extracted ONCE and every requested program is built
            from the same tables, each saved as its own file.

    Returns:
        (xlsx_out, pdf_out, resolved_version, resolved_appetite) when
        program is a single name; ([xlsx_outs], [pdf_outs],
        resolved_version, resolved_appetite) in the same order when it is a
        list. resolved_version/resolved_appetite are the input
        version/appetite, except when version was "Default" — then both are
        resolved per-state from the "Version By State" tab.
    """
    single = isinstance(program, str)
    programs = [program] if single else list(program)
    if version not in VALID_VERSIONS:
        raise ValueError(f"version must be one of {VALID_VERSIONS}, got {version!r}")
    if not programs:
        raise ValueError("program list is empty — select at least one program")
    for prog in programs:
        if prog not in VALID_PROGRAMS:
            raise ValueError(f"program must be one of {VALID_PROGRAMS}, got {prog!r}")

    t_start = time.perf_counter()
    if progress_callback: progress_callback("Initializing...")
    print(f"Creating BOP {', '.join(programs)} Rate Pages ({version})")

    warnings.simplefilter("ignore")
    pd.set_option("display.max_columns", None)
    pd.options.display.width = None

    # ── 1. Open every ratebook ─────────────────────────────────────────────
    t_stage = time.perf_counter()
    if progress_callback: progress_callback("Opening uploaded ratebooks...")
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
    if not CWRatebook and progress_callback:
        progress_callback("Fetching default CW ratebook (network drive)...")
    cw_source = CWRatebook if CWRatebook else str(BOP_CW_RATEBOOK_DEFAULT)
    cw_file = load_ratebook(cw_source)
    if cw_file == "Not found":
        raise ValueError(f"CW ratebook could not be loaded: {BOP_CW_RATEBOOK_DEFAULT}")
    print(f"Stage 1: Ratebooks opened in {time.perf_counter() - t_stage:0.1f}s")

    # ── 2. Extract state / date metadata (same 'Rate Book Details' layout BA uses) ──
    info = get_rate_book_info(ngic_loaded=ratebooks["NGIC"], mm_loaded=ratebooks["MM"])

    # ── 3. Load config-driven rating lookup tables ─────────────────────────
    cfg = load_bop_config()
    if info.state_abb not in cfg.perils_by_state:
        raise ValueError(
            f"No 'Perils By State' entry for '{info.state_abb}' in BOP Input File.xlsx — "
            "add a row there before generating this state's rate pages."
        )
    perils = cfg.perils_by_state[info.state_abb]

    # "Default" resolves to this state's version and Appetite flag per the
    # "Version By State" tab (see resolve_default_version /
    # resolve_default_appetite) — must happen before the Territory
    # Definitions load below, since that depends on the resolved version.
    # An explicit version ("2.0"/"pre2.0") keeps the caller's `appetite`
    # value (defaulting to False when not given).
    if version == "Default":
        version = resolve_default_version(cfg, info.state_abb)
        appetite = resolve_default_appetite(cfg, info.state_abb)
    elif appetite is None:
        appetite = False

    # ── 4. Load Territory Definitions (2.0 All Programs only — pre2.0 and
    #       All Peril never use them) ──
    territory_defs_by_st = None
    if version == "2.0" and "All Programs" in programs:
        if progress_callback: progress_callback("Loading Territory Definitions...")
        territory_defs_by_st = load_territory_defs(info.state_abb)

    eq_territory_defs = None
    if "Optional Coverages" in programs:
        if progress_callback: progress_callback("Loading Earthquake Territory Definitions...")
        eq_territory_defs = load_eq_territory_defs(info.state_abb)

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
    t_stage = time.perf_counter()
    if progress_callback: progress_callback("Extracting rate tables from ratebooks...")
    rate_tables_raw = load_all_ratebooks(rate_books, progress_callback)
    print(f"Stage 2: Rate tables extracted in {time.perf_counter() - t_stage:0.1f}s")
    # Drop companies that were not provided so AllPrograms.buildDataFrame's
    # "'NACO' in self.rateTables.keys()" optional-company checks behave
    # correctly (a present-but-None entry would otherwise crash on
    # None.keys()).
    rate_tables = {k: v for k, v in rate_tables_raw.items() if v is not None}

    # ── 6-8. Build, save and page-break each requested program ──────────────
    # The expensive part (opening + extracting the ratebooks above) is shared;
    # each program only costs its own workbook build and save.
    out_dir     = Path(folder_selected)
    base_tag    = "" if version == "2.0" else " (Pre 2.0)"
    version_tag = base_tag + (" (Appetite)" if appetite else "")
    xlsx_outs: List[str] = []
    pdf_outs:  List[str] = []

    for prog in programs:
        prefix = f"[{prog}] " if len(programs) > 1 else ""
        cb = (lambda msg, _p=prefix: progress_callback(f"{_p}{msg}")) if progress_callback else None

        t_stage = time.perf_counter()
        if cb: cb("Building Excel rate pages...")
        if prog == "All Peril":
            peril_cls = _version_cls(version, appetite, AllPerilPage.AllPeril, AllPerilPageCurrent.AllPeril, prog)
            rate_pages_obj = peril_cls(
                info.state_abb, rate_tables, cfg.class_codes,
                cfg.protection_class_conversions, cfg.building_codes_by_state,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildAllPerilPage(progress_callback=cb)
        elif prog == "Hab":
            hab_cls = _version_cls(version, appetite, HabPage.Hab, HabPageCurrent.Hab, prog)
            rate_pages_obj = hab_cls(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildHabPage(progress_callback=cb)
        elif prog == "Auto Service":
            auto_cls = _version_cls(version, appetite, AutoServicePage.Auto, AutoServicePageCurrent.Auto, prog)
            rate_pages_obj = auto_cls(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildAutoPage(progress_callback=cb)
        elif prog == "Retail":
            retail_cls = _version_cls(version, appetite, RetailPage.Retail, RetailPageCurrent.Retail, prog)
            rate_pages_obj = retail_cls(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildRetailPage(progress_callback=cb)
        elif prog == "Service":
            service_cls = _version_cls(version, appetite, ServicePage.Service, ServicePageCurrent.Service, prog)
            rate_pages_obj = service_cls(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildServicePage(progress_callback=cb)
        elif prog == "Office":
            office_cls = _version_cls(version, appetite, OfficePage.Office, OfficePageCurrent.Office, prog)
            rate_pages_obj = office_cls(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildOfficePage(progress_callback=cb)
        elif prog == "Wholesale":
            wholesale_cls = _version_cls(version, appetite, WholesalePage.Wholesale, WholesalePageCurrent.Wholesale, prog)
            rate_pages_obj = wholesale_cls(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildWholesalePage(progress_callback=cb)
        elif prog == "Food Service":
            food_cls = _version_cls(version, appetite, FoodServicePage.Food, FoodServicePageCurrent.Food, prog)
            rate_pages_obj = food_cls(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildFoodPage(progress_callback=cb)
        elif prog == "Optional Coverages":
            # No version split — same class regardless of `version`.
            rate_pages_obj = OptionalCoveragesPage.OptionalCoverages(
                info.state_abb, rate_tables, cfg.class_codes,
                info.n_effective, info.r_effective, eq_territory_defs,
            )
            bop_workbook = rate_pages_obj.buildOptionalCoveragesPage(progress_callback=cb)
        elif prog == "Rating Plans":
            # No version split — same class regardless of `version`.
            rate_pages_obj = RatingPlansPage.RatingPlans(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective, irpm_credit, irpm_debit,
            )
            bop_workbook = rate_pages_obj.buildRatingPlansPage(progress_callback=cb)
        elif prog == "Class Modifier":
            # No version split — same class regardless of `version`.
            rate_pages_obj = ClassModifierPage.ClassModifier(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildClassModifierPage(progress_callback=cb)
        elif prog == "Common Rules":
            # No version split — same class regardless of `version`.
            rate_pages_obj = CommonRulesPage.CommonRules(
                info.state_abb, rate_tables, perils, cfg.peril_conversions,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildCommonRulesPage(progress_callback=cb)
        elif prog == "Additional Rules":
            # No version split — same class regardless of `version`.
            rate_pages_obj = AdditionalRulesPage.AdditionalRules(
                info.state_abb, rate_tables, perils, cfg.peril_conversions, cfg.class_codes,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildAdditionalRulesPage(progress_callback=cb)
        elif version == "2.0":
            # No Appetite-specific "All Programs" page yet — builds plain
            # 2.0 regardless of `appetite` (see APPETITE_CLASSES).
            rate_pages_obj = AllProgramsPage.AllPrograms(
                info.state_abb, rate_tables, perils,
                cfg.peril_conversions, cfg.protection_class_conversions,
                cfg.building_codes_by_state,
                info.n_effective, info.r_effective, territory_defs_by_st,
            )
            bop_workbook = rate_pages_obj.buildAllProgramsPage(progress_callback=cb)
        else:
            rate_pages_obj = AllProgramsPageCurrent.AllPrograms(
                info.state_abb, rate_tables, perils,
                cfg.peril_conversions, cfg.protection_class_conversions,
                cfg.building_codes_by_state,
                info.n_effective, info.r_effective,
            )
            bop_workbook = rate_pages_obj.buildAllProgramsPage(progress_callback=cb)
        print(f"Stage 3: {prog} rate pages built in {time.perf_counter() - t_stage:0.1f}s")

        t_stage = time.perf_counter()
        if cb: cb("Saving Excel file...")
        file_stem = f"{info.state_abb} {info.n_effective} BOP {prog} Rate Pages{version_tag}"
        xlsx_out  = str(out_dir / f"{file_stem}.xlsx")
        pdf_out   = str(out_dir / f"{file_stem}.pdf")

        bop_workbook.active = bop_workbook["Index"]
        bop_workbook.save(filename=xlsx_out)
        print(f"Stage 4: {prog} Excel file saved in {time.perf_counter() - t_stage:0.1f}s")

        t_stage = time.perf_counter()
        if cb: cb("Applying page breaks...")
        process_pagebreaks(xlsx_out, pdf_out, progress_callback=cb)
        print(f"Stage 5: {prog} page breaks applied in {time.perf_counter() - t_stage:0.1f}s")

        xlsx_outs.append(xlsx_out)
        pdf_outs.append(pdf_out)

    elapsed = time.perf_counter() - t_start
    if progress_callback: progress_callback(f"Successfully completed in {elapsed:0.1f} seconds! 🎉")
    print(f"This program ran in {elapsed:0.4f} seconds")

    if single:
        return xlsx_outs[0], pdf_outs[0], version, appetite
    return xlsx_outs, pdf_outs, version, appetite


def generate_pdf_only(
    xlsx_path: str,
    pdf_path: str,
    progress_callback: Optional[Callable[[str], None]] = None,
    exclude_sheets: Optional[Sequence[str]] = None,
    max_pdf_mb: Optional[float] = None,
) -> List[str]:
    """
    Convert an existing rate-pages .xlsx into a PDF using Excel's own print
    engine, so every page-break / fit-to-page / print-area setting already
    written to the workbook is honored.

    exclude_sheets: sheet names to leave out of the PDF entirely — used by
    the "All Programs" (2.0) build to skip TERRITORY_DEFS_SHEET, which is
    generated separately/optionally via generate_territory_defs_pdf.

    max_pdf_mb: if given and the resulting PDF is larger than this, it is
    split into "<name>_part1.pdf", "<name>_part2.pdf", etc., each under the
    limit (see split_pdf_by_size) — a sheet's pages are kept together in one
    part whenever they fit.

    Returns the list of PDF paths actually produced — a single-item list
    unless max_pdf_mb forced a split.
    """
    import os
    if progress_callback: progress_callback("Launching Excel...")
    t0 = time.perf_counter()
    out = export_to_pdf(xlsx_path, pdf_path, progress_callback=progress_callback,
                         exclude_sheets=exclude_sheets)
    if not (os.path.exists(out) and os.path.getsize(out) > 0):
        raise RuntimeError(f"PDF was not created at {out}")
    outs = split_pdf_by_size(out, max_pdf_mb, xlsx_path=xlsx_path,
                              exclude_sheets=exclude_sheets,
                              progress_callback=progress_callback)
    elapsed = time.perf_counter() - t0
    if progress_callback:
        names = ", ".join(os.path.basename(o) for o in outs)
        progress_callback(f"PDF created in {elapsed:0.1f}s — {names} 🎉")
    print(f"[BOPRatePages] PDF generated: {outs}")
    return outs


def territory_defs_pdf_path(xlsx_path: str) -> str:
    """The output path generate_territory_defs_pdf writes to for a given
    "All Programs" xlsx, kept alongside it."""
    p = Path(xlsx_path)
    return str(p.with_name(f"{p.stem} - Territory Definitions.pdf"))


def generate_territory_defs_pdf(
    xlsx_path: str,
    progress_callback: Optional[Callable[[str], None]] = None,
    max_pdf_mb: Optional[float] = None,
) -> List[str]:
    """
    Export just the Territory Definitions ("TRDEF") sheet of an "All
    Programs" (2.0) workbook to its own PDF. Optional companion to
    generate_pdf_only, which excludes that sheet from the main PDF since it
    alone can dominate export time (82k rows).

    max_pdf_mb: if given and the resulting PDF is larger than this, it is
    split into "<name>_part1.pdf", "<name>_part2.pdf", etc. (see
    split_pdf_by_size) — TRDEF is a single sheet, so a split here is always
    a plain page-count cut, same as any other single-sheet PDF over the
    limit.

    Returns the list of PDF paths actually produced.
    """
    import os
    pdf_path = territory_defs_pdf_path(xlsx_path)
    if progress_callback: progress_callback("Launching Excel...")
    t0 = time.perf_counter()
    out = export_single_sheet_pdf(xlsx_path, pdf_path, TERRITORY_DEFS_SHEET,
                                   progress_callback=progress_callback)
    if not (os.path.exists(out) and os.path.getsize(out) > 0):
        raise RuntimeError(f"PDF was not created at {out}")
    outs = split_pdf_by_size(out, max_pdf_mb, progress_callback=progress_callback)
    elapsed = time.perf_counter() - t0
    if progress_callback:
        names = ", ".join(os.path.basename(o) for o in outs)
        progress_callback(f"Territory Definitions PDF created in {elapsed:0.1f}s — {names} 🎉")
    print(f"[BOPRatePages] Territory Definitions PDF generated: {outs}")
    return outs
