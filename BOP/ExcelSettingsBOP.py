"""
ExcelSettingsBOP.py
====================
Excel workbook / worksheet factory for BOP Rate Pages, mirroring the shape
of BA/ExcelSettingsBA.py.

WHY THIS FILE LOOKS DIFFERENT FROM BA'S VERSION
------------------------------------------------
BA has ~15 hand-written formatWorksheet* methods because its sheets have
genuinely different layouts (stacked tables, two-part Zone tables, the fixed
Rule 222 grid, ...). BOP's All Programs page has one layout used ~30 times
(title, optional merged sub-header row, one data table) with only the
column widths / number formats / sub-header text differing per table. So
instead of ~30 near-duplicate formatXxx() methods, there is ONE generic
format_table(ws, table_code) that looks up those per-table differences in
"BOP/BOP Input File.xlsx" (loaded via bop_config.load_bop_config()).

This is also what makes header/sub-header/title/table/footer sizing
consistent across every page: every sheet reads the exact same
Formatting Defaults values instead of each format method picking its own.
"""

from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.dataframe import dataframe_to_rows

from config.constants import COMPANY_NAMES
import re

import pandas as pd

from .bop_config import load_bop_config, is_no_hab_state

# Peril display names (see the "Peril Conversions" config tab / cat4, liability1-4)
# whose text after the hyphen is longer than their neighboring peril columns
# (e.g. "Products", "SlipFall", "OtherPrem" vs. "BINC", "Water", "Wind"). These
# share a uniform per-table column width with their shorter neighbors, so at
# that shared width they wrap to 3 lines ("L-" / "Product" / "s") instead of 2
# ("L-" / "Products"). Enforced as a floor on top of the configured Table
# Layout width, table-code-agnostic, so it self-applies to every BOP sheet
# these labels appear on instead of needing a per-table-per-column override.
_WIDE_PERIL_HEADERS = {"L-Products", "L-SlipFall", "L-Violence", "L-OtherMed", "L-OtherPrem"}
_WIDE_PERIL_HEADER_MIN_PX = 75


# Small Market context. Small Market has no per-company ratebooks (just
# SM -> ATA -> CW), so its footer lists the underwriting companies named in
# the ratebook's "Company Deviation Factor_Ext" sheet instead of rateTables
# keys, and its single merged base-rate sheet is titled "Small Market" rather
# than after NGIC. Set by BOPRatePages.run() for the duration of a Small
# Market build and always cleared afterwards (see set_small_market_context).
_SMALL_MARKET = {"footer_names": None}
_NGIC_TITLE_LABEL = "NW General Insurance Company"
_SM_TITLE_LABEL = "Small Market"
_FOOTER_LINE_INCHES = 0.14  # one 8pt footer line, with leading

# Excel rejects the whole workbook ("We found a problem with some content" /
# COM Open failure when exporting to PDF) if a sheet's header or footer string
# exceeds 255 characters in total: the "&L"/"&C"/"&R" markers plus the text of
# every section, font codes like &"Arial,Bold"&8 (15 chars) included, a line
# break counting as 1. Verified against Excel itself. Kept a little under.
_HF_STORED_LIMIT = 250

# Progressively shorter spellings used only when the full company names can't
# fit in the footer (see Excel._fit_footer_names). Applied cumulatively.
_ABBREVIATION_LEVELS = [
    [],
    [("Insurance Company", "Ins Co")],
    [("Property & Casualty", "P&C"), ("Company", "Co")],
    [("Nationwide", "NW"), ("Insurance", "Ins")],
]
_OVERFLOW_LINE = "+{n} more - see RP Table 96.C"


# ── "No Hab" state exception (see bop_config.NO_HAB_STATES) ────────────────
# Labels that mean the Habitational program. "H" alone is ambiguous (grade
# letters etc.), so it only counts in a column headed "Program".
_HAB_LABELS = {"hab", "habitational", "habitational (nrp)", "hab (nrp)"}
_HAB_CLASS_CODE_HEADERS = {"code", "class code", "classcode"}
_HAB_CLASS_CODE_RANGE = (10000, 19999)   # Class_Code_Min 10000 = Hab, 20000 = Auto
_HAB_AND_PREFIX = re.compile(r"^Habitational\s*(?:,|and|&)\s*(?:and\s+|&\s*)?", re.IGNORECASE)


def is_hab_only_title(title) -> bool:
    """True for a title that belongs to the Habitational program alone
    ("H Table 3.C.2.f. ..." with no other program's table in it)."""
    return (isinstance(title, str) and title.startswith("H Table ")
            and len(re.findall(r"\bTable\b", title)) == 1)


def strip_hab_from_title(title):
    """
    Remove the Habitational ("H") program from a table title, e.g.
      "AS, FS, H, O, R, S, W Table 3.C.2.b. ..."  -> "AS, FS, O, R, S, W Table 3.C.2.b. ..."
      "H Table 3.C.2.m., AS, FS, O, R, S, W Table 3.C.2.n. ..." -> "AS, FS, O, R, S, W Table 3.C.2.n. ..."
      "..., R Table 3.C.4.d., AS, H Table 3.C.4.e. ..." -> "..., R Table 3.C.4.d., AS Table 3.C.4.e. ..."
    A Hab-only title is returned unchanged; callers drop those sheets via
    is_hab_only_title().
    """
    if not isinstance(title, str) or is_hab_only_title(title):
        return title
    t = re.sub(r"^H Table [^\s,]+,\s*", "", title)      # leading H-only segment
    t = re.sub(r",\s*H(?= Table\b)", "", t)             # "AS, H Table x" -> "AS Table x"
    t = re.sub(r"(?<![\w.])H,\s*", "", t)               # "FS, H, O" -> "FS, O"
    return t


def _is_text(series) -> bool:
    return series.dtype == object or pd.api.types.is_string_dtype(series)


def scrub_hab_df(df, drop_rows=True):
    """
    Return `df` with every Habitational (Hab) trace removed:
      - columns headed Hab / Habitational / Habitational (NRP);
      - rows labelled Hab / Habitational (or "H" under a "Program" header) in
        either of the first two columns, and rows whose class-code column
        ("Code" / "Class Code") is a Hab class code (10000-19999);
      - the wording "Habitational and Office Programs" -> "Office Programs".
    drop_rows=False leaves rows alone (multi-table sheets, whose formatters
    index their blocks by the original row counts).
    """
    if df is None or len(df.columns) == 0:
        return df
    is_hab_col = [str(c).strip().lower() in _HAB_LABELS for c in df.columns]
    if any(is_hab_col):
        df = df.loc[:, [not h for h in is_hab_col]]
    if not drop_rows or len(df) == 0 or len(df.columns) == 0:
        return df

    drop = pd.Series(False, index=df.index)
    for pos in range(min(2, df.shape[1])):
        col = df.iloc[:, pos]
        header = str(df.columns[pos]).strip().lower()
        if _is_text(col):
            text = col.astype("string").str.strip()
            drop |= text.str.lower().isin(_HAB_LABELS).fillna(False).to_numpy()
            if header == "program":
                drop |= (text == "H").fillna(False).to_numpy()
            rewrite = text.str.match(_HAB_AND_PREFIX).fillna(False)
            if rewrite.any():
                df = df.copy()
                df.isetitem(pos, col.where(~rewrite, text.str.replace(_HAB_AND_PREFIX, "", regex=True)))
        if header in _HAB_CLASS_CODE_HEADERS:
            num = pd.to_numeric(col, errors="coerce")
            drop |= num.between(*_HAB_CLASS_CODE_RANGE).fillna(False).to_numpy()
    return df.loc[~drop] if drop.any() else df


def df_has_data(df) -> bool:
    """
    True if `df` has at least one real value. A table whose dataframe has
    columns but no rows (or only blank/NaN cells) prints as a bare header
    row — e.g. Florida has no Windstorm/Hail "Per Building" deductible rows,
    so that sheet came out as just a title and a header. Such tables are
    dropped (see Excel.removeBlankSheets and filter_nonempty_blocks).
    """
    if df is None or len(df) == 0 or len(df.columns) == 0:
        return False
    try:
        return bool(df.replace("", float("nan")).notna().any().any())
    except Exception:
        return True  # unsure -> keep the table rather than silently drop it


def _filled_cells(ws) -> int:
    return sum(1 for row in ws.iter_rows() for c in row if c.value is not None)


def filter_nonempty_blocks(ws, blocks):
    """
    For the hand-built "labeled block" sheets (a label + a dataframe, stacked
    on one sheet — see each program page's _appendLabeledBlocks): return only
    the (label, df) blocks that actually have data, and record on the sheet
    how many blocks were offered vs. kept so Excel.removeBlankSheets can drop
    the whole sheet when none of its tables has a value.
    """
    st = ws.__dict__.setdefault("_bop_blank", {})
    if st.get("no_hab"):
        blocks = [(strip_hab_from_title(label), scrub_hab_df(df)) for label, df in blocks]
        blocks = [(l, d) for l, d in blocks if str(l).strip().lower() not in _HAB_LABELS]
    kept = [(label, df) for label, df in blocks if df_has_data(df)]
    st["extra_total"] = st.get("extra_total", 0) + len(blocks)
    st["extra_kept"] = st.get("extra_kept", 0) + len(kept)
    return kept


def set_small_market_context(footer_names=None) -> None:
    """Activate (list of company names) or clear (None) Small Market mode."""
    _SMALL_MARKET["footer_names"] = [str(n).strip() for n in footer_names] if footer_names else None


class Excel:
    """
    Factory class that creates an openpyxl Workbook and populates it with
    formatted worksheets, one per rate table.

    Typical usage (from AllProgramsPage.AllPrograms.buildAllProgramsPage):
        pages = ExcelSettingsBOP.Excel(state, "All Programs", nEff, rEff, companies)
        pages.generateWorksheet("SPR", title, df, False, True)
        ...
        wb = pages.getWB()
    """

    # Short section code for the {section} footer token, keyed by the exact
    # programName each page module passes in (see the ExcelSettingsBOP.Excel(...)
    # call in AllProgramsPage.py / AllPerilPage.py / HabPage.py / AutoServicePage.py /
    # ServicePage.py). A programName missing from this map falls back to
    # printing itself in full in the footer (that's what happened to Service
    # before it was added here).
    _SECTION_CODES = {
        "All Programs": "PG",
        "All Peril": "AP",
        "Habitational": "H",
        "Auto Service": "AS",
        "Service": "S",
        "Office": "O",
        "Wholesale": "W",
        "Food Service": "FS",
        "Retail": "R",
        "Optional Coverages": "OC",
        "Additional Rules": "AR",
        "Common Rules": "CR",
        "Rating Plans": "RP",
    }

    def __init__(self, state, programName, nEffective, rEffective, companyList) -> None:
        self.wb = Workbook()
        self.wb.active.title = "Index"

        self.state = state
        self.noHab = is_no_hab_state(state)   # state exception: no Habitational program
        self.programName = programName
        self.nEffective = nEffective
        self.rEffective = rEffective
        self.companyList = list(companyList) if companyList else []

        cfg = load_bop_config()
        self._cfg = cfg
        fmt = cfg.formatting

        self.fontName = fmt.get("FontName", "Arial")
        self.fontSize = fmt.get("FontSize", 10)
        self.headerFontName = fmt.get("HeaderFontName", self.fontName)
        self.headerFontSize = fmt.get("HeaderFontSize", self.fontSize)
        self.footerFontName = fmt.get("FooterFontName", self.fontName)
        self.footerFontSize = fmt.get("FooterFontSize", self.fontSize)
        self.leftMargin = fmt.get("LeftMargin", 0.25)
        self.rightMargin = fmt.get("RightMargin", 0.25)
        self.topMargin = fmt.get("TopMargin", 1.25)
        self.bottomMargin = fmt.get("BottomMargin", 0.95)
        self.headerMargin = fmt.get("HeaderMargin", 0.5)
        self.footerMargin = fmt.get("FooterMargin", 0.25)
        self.borderColor = fmt.get("BorderColor", "C1C1C1")
        self.currencyFormat = fmt.get("CurrencyFormat", "$#,##0.000")
        self.noDecimalFormat = fmt.get("NoDecimalFormat", "#,##0.000")
        self.zipCodeFormat = fmt.get("ZipCodeFormat", "####0")
        # Factor/value columns with no explicit Number Formats entry (most
        # tables' rate-factor columns) otherwise fall back to Excel's
        # "General" format. All table values display at 3 decimal places
        # except zip/pin codes, so this is the default for those columns.
        self.defaultValueFormat = "#,##0.000"
        self._defaultPrintTitleRows = fmt.get("PrintTitleRows", "1:3")

        self.font = Font(name=self.fontName, size=self.fontSize)
        self.fontBold = Font(name=self.fontName, size=self.fontSize, bold=True)
        self.fontItalic = Font(name=self.fontName, size=self.fontSize, italic=True)
        # Shared style objects — constructing a new Alignment per cell makes
        # openpyxl re-hash the style for every cell, which turns large sheets
        # (TRDEF has tens of thousands of rows) into a multi-minute format pass.
        self._centerAlign = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
        self.headerFont = f"{self.headerFontName},Bold"
        self.footerFont = f"{self.footerFontName},Bold"

        self._thinBorder = Border(
            left=Side(border_style="thin", color=self.borderColor),
            right=Side(border_style="thin", color=self.borderColor),
            top=Side(border_style="thin", color=self.borderColor),
            bottom=Side(border_style="thin", color=self.borderColor),
        )
        self._formatMap = {
            "Currency": self.currencyFormat,
            "NoDecimal": self.noDecimalFormat,
            "ZipCode": self.zipCodeFormat,
        }

    # ==========================================================================
    #  PRIVATE HELPERS
    # ==========================================================================

    def _apply_page_setup(self, ws) -> None:
        ws.page_setup.orientation = "portrait"
        ws.page_setup.blackAndWhite = False
        ws.page_setup.firstPageNumber = 1
        ws.page_setup.useFirstPageNumber = True
        ws.page_setup.fitToHeight = False
        ws.sheet_view.showGridLines = False
        ws.print_title_rows = self._defaultPrintTitleRows
        ws.page_margins.left = self.leftMargin
        ws.page_margins.right = self.rightMargin
        ws.page_margins.top = self.topMargin
        ws.page_margins.bottom = self._bottomMarginFor(len(self._footer_names()))
        ws.page_margins.header = self.headerMargin
        ws.page_margins.footer = self.footerMargin

    def _token_sub(self, template: str) -> str:
        return template.format(
            state=self.state,
            state_abb=self.state,
            n_effective=self.nEffective,
            r_effective=self.rEffective,
            companies=self._footer_companies_text(),
            program=self.programName,
            section=self._SECTION_CODES.get(self.programName, self.programName),
        )

    def _footer_names(self) -> list:
        # .strip() guards against stray leading/trailing whitespace on any
        # individual company name (e.g. NGIC showing indented relative to
        # the other companies in the footer) regardless of where it enters.
        if _SMALL_MARKET["footer_names"]:
            return _SMALL_MARKET["footer_names"]
        return [COMPANY_NAMES.get(c, c).strip() for c in self.companyList if c != "CW"]

    def _bottomMarginFor(self, n_names: int) -> float:
        # The configured bottom margin fits the usual <= 5 footer lines; more
        # companies than that (Small Market can list many) would otherwise
        # print over the table, so grow the margin to hold them all.
        needed = self.footerMargin + n_names * _FOOTER_LINE_INCHES + 0.05
        return max(self.bottomMargin, round(needed, 2))

    def _footer_companies_text(self) -> str:
        return self._companies_text(self._footer_names(), trailing_space=True)

    @staticmethod
    def _companies_text(names, trailing_space: bool) -> str:
        # "&" starts a control code in Excel header/footer text ("& C" would
        # switch sections), so a literal ampersand — e.g. "Allied Property &
        # Casualty" — must be doubled.
        end = " " if trailing_space else ""
        return "\n".join(f"{n.replace('&', '&&')}{end}" for n in names)

    @staticmethod
    def _stored_length(*sections: str) -> int:
        # What Excel measures: "&L"/"&C"/"&R" markers + each section's text.
        return 6 + sum(len(t) for t in sections)

    def _fit_footer_names(self, center_text: str) -> str:
        """
        Left-footer text listing the companies, shortened only as much as
        needed for the whole footer to stay within Excel's length limit.
        Full names exactly as before whenever they fit; otherwise drop the
        trailing spaces, then abbreviate step by step ("Insurance Company" ->
        "Ins Co", "Property & Casualty" -> "P&C", "Nationwide" -> "NW", ...).
        If even that can't fit every company, keep as many as fit plus a
        "+N more" line rather than ever writing a workbook Excel rejects.
        """
        prefix = len(f'&"{self.footerFont}"&{self.footerFontSize}')
        budget = _HF_STORED_LIMIT - 6 - prefix - (prefix + len(center_text))
        names = self._footer_names()

        def fits(text):
            return len(text) <= budget

        for level, trailing in ((0, True), (0, False), (1, False), (2, False), (3, False)):
            short = names
            for lv in range(1, level + 1):
                for old, new in _ABBREVIATION_LEVELS[lv]:
                    short = [n.replace(old, new) for n in short]
            text = self._companies_text(short, trailing)
            if fits(text):
                return text

        for keep in range(len(short) - 1, 0, -1):
            text = self._companies_text(short[:keep] + [_OVERFLOW_LINE.format(n=len(short) - keep)], False)
            if fits(text):
                print(f"[ExcelSettingsBOP] WARNING: {len(names)} footer company names exceed Excel's footer "
                      f"length limit; listing {keep} and '+{len(short) - keep} more'.")
                return text
        return ""

    @staticmethod
    def _set_hf_part(part, text: str, font: str, size: int) -> None:
        """
        Assign text/font/size to a header or footer section (ws.oddHeader.left,
        ws.oddFooter.center, etc.) without going through the .font/.size
        setters.

        openpyxl's own serialization (_HeaderFooterPart.__str__) renders the
        size code as "&{size} " — WITH a trailing space, to keep a text that
        starts with a digit from merging into the size number (e.g. so "&10"
        + "5th Ave" can't be misread as size 105). For multi-line text that
        space becomes the very first character rendered, so only the FIRST
        line is indented — e.g. NGIC (always first in the footer company
        list) sat visibly indented while every other company didn't. None of
        our header/footer text starts with a digit, so the codes are safe to
        embed directly in .text and skip .font/.size entirely, avoiding the
        stray space.
        """
        if not text:
            # An empty section's font code alone costs 15 of the 255
            # characters Excel allows the whole footer; leave it out.
            part.text = None
            return
        part.text = f'&"{font}"&{size}{text}'

    def _apply_standard_header(self, ws) -> None:
        hf = self._cfg.header_footer
        self._set_hf_part(ws.oddHeader.left, hf.get("HeaderLeftText", ""), self.headerFont, self.headerFontSize)
        self._set_hf_part(ws.oddHeader.center, self._token_sub(hf.get("HeaderCenterTemplate", "")), self.headerFont, self.headerFontSize)
        self._set_hf_part(ws.oddHeader.right, self._token_sub(hf.get("HeaderRightTemplate", "")), self.headerFont, self.headerFontSize)

    def _apply_default_footer(self, ws) -> None:
        hf = self._cfg.header_footer
        center = self._token_sub(hf.get("FooterCenterTemplate", ""))
        left_template = hf.get("FooterLeftTemplate", "")
        if left_template.strip() == "{companies}":
            left = self._fit_footer_names(center)
        else:
            left = self._token_sub(left_template)
        self._set_hf_part(ws.oddFooter.left, left, self.footerFont, self.footerFontSize)
        self._set_hf_part(ws.oddFooter.center, center, self.footerFont, self.footerFontSize)
        self._set_hf_part(ws.oddFooter.right, self._token_sub(hf.get("FooterRightTemplate", "")), self.footerFont, self.footerFontSize)

    def _apply_page_header_footer(self, ws) -> None:
        self._apply_page_setup(ws)
        self._apply_standard_header(ws)
        self._apply_default_footer(ws)

    def _write_df_block(self, ws, df, use_index: bool, use_header: bool, index_cell: str) -> None:
        for r in dataframe_to_rows(df, use_index, use_header):
            if use_index and len(list(r)) == 1:
                ws[index_cell] = list(r)[0]
                continue
            ws.append(r)

    def _resolve_col_range(self, range_str: str, max_col: int):
        if not range_str:
            return None
        start, _, end = range_str.partition(":")
        start_idx = column_index_from_string(start)
        end_idx = max_col if end.strip().upper() == "REST" else column_index_from_string(end)
        return start_idx, end_idx

    def _apply_sub_header(self, ws, table_code: str) -> int:
        """
        Inserts the merged sub-header row for tables that need one (per the
        "Sub Headers" tab), e.g. a row reading "Amount of Insurance" spanning
        several columns above the real column headers.

        Returns the inserted row number, or 0 if this table has no sub-header.
        """
        sh = self._cfg.sub_headers.get(table_code)
        if not sh:
            return 0

        row_idx = sh["insert_at_row"]
        ws.insert_rows(row_idx)
        max_col = ws.max_column

        for label_range, label_text in (
            (sh["label1_range"], sh["label1_text"]),
            (sh["label2_range"], sh["label2_text"]),
        ):
            rng = self._resolve_col_range(label_range, max_col)
            if not rng:
                continue
            start_idx, end_idx = rng
            if label_text:
                ws.cell(row=row_idx, column=start_idx).value = label_text
            if end_idx > start_idx:
                ws.merge_cells(start_row=row_idx, start_column=start_idx,
                                end_row=row_idx, end_column=end_idx)

        for col in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col)
            cell.border = self._thinBorder
            cell.font = self.fontBold
            cell.alignment = self._centerAlign

        if sh["print_title_rows"]:
            ws.print_title_rows = sh["print_title_rows"]

        return row_idx

    # ==========================================================================
    #  WORKSHEET GENERATORS
    # ==========================================================================

    def _retitle(self, title: str) -> str:
        # Small Market's merged base-rate sheet is built through the NGIC slot
        # (see BOPRatePages._merge_small_market), whose sheet title names NGIC.
        if _SMALL_MARKET["footer_names"] and isinstance(title, str):
            title = title.replace(_NGIC_TITLE_LABEL, _SM_TITLE_LABEL)
        if self.noHab:
            title = strip_hab_from_title(title)
        return title

    def isHabOnlySheet(self, title) -> bool:
        """True when this sheet is Hab-only and the state has no Hab: callers
        skip it before building (see bop_config.NO_HAB_STATES)."""
        return self.noHab and is_hab_only_title(title)

    def generateWorksheet(self, table_code, title, df, useIndex, useHeader, layout_key=None):
        """
        Single-table sheet: A1 = title, A2 = blank, row 3+ = data (headers + rows).

        layout_key: which key to look up in the Table Layout / Number Formats /
        Sub Headers config tabs, if different from table_code (the actual tab
        name). Used e.g. by BCEG, whose tab is always named "BCEG" but whose
        column-width profile depends on whether the state has multiple
        building-code groups (BCEG_MULTI vs BCEG_SINGLE).
        """
        ws = self.wb.create_sheet(title=table_code)
        ws["A1"] = self._retitle(title)
        ws["A2"] = ""
        if self.noHab:
            df = scrub_hab_df(df)
        self._write_df_block(ws, df, useIndex, useHeader, index_cell="A3")
        self.format_table(ws, layout_key or table_code)
        # No columns at all = a sheet whose body the page builds itself (its
        # own postFormat / _appendLabeledBlocks), so it can't be judged here.
        ws.__dict__.setdefault("_bop_blank", {})["no_hab"] = self.noHab
        ws.__dict__.setdefault("_bop_blank", {})["base"] = (
            "custom" if df is None or len(df.columns) == 0
            else "data" if df_has_data(df) else "empty")
        ws._bop_blank["cells"] = _filled_cells(ws)
        return ws

    def generateTerrWorksheet(self, table_code, title, df, useIndex, useHeader):
        """Territory-definitions sheet — same single-table layout as generateWorksheet."""
        return self.generateWorksheet(table_code, title, df, useIndex, useHeader)

    def generateMultiTableWorksheet(self, table_code, title, dataframes, useIndex, useHeader,
                                     reserved_header_rows=0, layout_key=None):
        """
        Multi-block sheet: A1 = title, A2 = blank, then each dataframe in
        `dataframes` written contiguously (its own header row + data rows,
        one block right after another, same as calling generateWorksheet's
        single-block write repeatedly on one sheet).

        Replaces Optional Coverages' source (root OptionalCoveragesPage.py)
        generateWorksheet2tables .. generateWorksheet17tables / generateLiquorLiability
        — a family of ~10 near-identical N-ary methods on the (lost, no
        longer present anywhere in this repo) root ExcelSettings.py — with
        one generic implementation.

        reserved_header_rows: blank rows left before each block for a
        postFormat callback to fill in with its own sub-title/sub-header
        (e.g. "Territory 21 Loss Costs" / "Contents Grade" on the Earthquake
        Class Rated page) — see OptionalCoveragesPage._formatEQClassRated.
        Returns (block_start_rows, ws): block_start_rows[i] is the first
        reserved row for block i (or its header row when
        reserved_header_rows is 0).
        """
        ws = self.wb.create_sheet(title=table_code)
        ws["A1"] = self._retitle(title)
        ws["A2"] = ""
        block_start_rows = []
        if self.noHab:
            # columns only: callers' formatters index blocks by original row counts
            dataframes = [scrub_hab_df(d, drop_rows=False) for d in dataframes]
        for df in dataframes:
            if reserved_header_rows and block_start_rows:
                ws.append([])  # blank spacer row between blocks
            for _ in range(reserved_header_rows):
                ws.append([])
            # ws.max_row doesn't count the empty reserved rows (append([])
            # only advances the write cursor), so use the cursor for them.
            block_start_rows.append(ws._current_row + 1 - reserved_header_rows if reserved_header_rows
                                     else ws.max_row + 1)
            self._write_df_block(ws, df, useIndex, useHeader, index_cell=f"A{ws.max_row + 1}")
        self.format_table(ws, layout_key or table_code)
        ws.__dict__.setdefault("_bop_blank", {})["no_hab"] = self.noHab
        ws.__dict__.setdefault("_bop_blank", {})["base"] = (
            "data" if any(df_has_data(d) for d in dataframes)
            else "empty" if any(len(d.columns) for d in dataframes) else "custom")
        ws._bop_blank["cells"] = _filled_cells(ws)
        return block_start_rows, ws

    # ==========================================================================
    #  GENERIC FORMATTER — replaces the ~30 near-duplicate format*() methods
    # ==========================================================================

    def format_table(self, ws, table_code: str) -> None:
        self._apply_page_header_footer(ws)

        sub_header_row = self._apply_sub_header(ws, table_code)
        header_row = (sub_header_row + 1) if sub_header_row else 3

        max_row, max_col = ws.max_row, ws.max_column

        # bestFit is a column-level property — set it once per column, not once
        # per cell (per-cell it dominated the build time of large sheets).
        for col in range(1, max_col + 1):
            ws.column_dimensions[get_column_letter(col)].bestFit = True

        # Title, spacer and column-header rows: a handful of cells, style
        # directly. (A sub-header row, if any, was styled by _apply_sub_header.)
        # The header row gets its own full box, matching the sub-header/data
        # rows, so tables with a sub-header (e.g. PD's "Amount of Insurance")
        # don't end up with the header text floating unboxed between the
        # sub-header's box above and the data's box below.
        for col in range(1, max_col + 1):
            ws.cell(row=1, column=col).font = self.fontBold
            ws.cell(row=2, column=col).font = self.fontItalic
            header_cell = ws.cell(row=header_row, column=col)
            header_cell.font = self.fontBold
            header_cell.alignment = self._centerAlign
            header_cell.border = self._thinBorder

        self._format_data_rows(ws, table_code, header_row + 1, max_row, max_col)

        for col_start, col_end, width_px in self._cfg.table_layout.get(table_code, []):
            end = ws.max_column if col_end == "REST" else col_end
            for col in range(col_start, end + 1):
                ws.column_dimensions[get_column_letter(col)].width = width_px / 7.0

        # Floor width for the long-form peril headers (see _WIDE_PERIL_HEADERS)
        # regardless of table code — a shared "REST" width sized for their
        # shorter neighbor perils leaves these wrapping to 3 lines.
        for col in range(1, max_col + 1):
            if ws.cell(row=header_row, column=col).value in _WIDE_PERIL_HEADERS:
                col_letter = get_column_letter(col)
                current_width_px = (ws.column_dimensions[col_letter].width or 0) * 7.0
                if current_width_px < _WIDE_PERIL_HEADER_MIN_PX:
                    ws.column_dimensions[col_letter].width = _WIDE_PERIL_HEADER_MIN_PX / 7.0

        for cell_ref, text in self._cfg.footnotes.get(table_code, []):
            c = ws[cell_ref]
            if c._style is not None:
                # Data cells share StyleArrays (see _format_data_rows); give
                # this cell a private copy before restyling it.
                c._style = copy(c._style)
            c.value = text
            c.font = self.font
            c.alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=False)
            c.border = Border()

    def _format_data_rows(self, ws, table_code: str, first_row: int, max_row: int, max_col: int) -> None:
        """
        Font / alignment / border / number-format for the data region.

        openpyxl rebuilds and re-hashes a cell's style on every .font /
        .alignment / .border / .number_format assignment, which made
        grid-level sheets (TRDEF: tens of thousands of rows) take minutes.
        Instead, build the handful of distinct styles once — through a scratch
        cell, so the ids land in the same workbook style registries openpyxl
        would use — then stamp the shared StyleArray straight onto each cell.

        The stamped arrays are SHARED between cells, so restyling a data cell
        afterwards through the normal descriptors would corrupt every cell
        sharing its array — replace cell._style with a copy first (as the
        footnote loop in format_table does).
        """
        if first_row > max_row:
            return

        # Per-column number format from config; for overlapping ranges the
        # last config row wins, matching the sequential application this
        # replaces.
        col_fmt = {}
        for col_start, col_end, row_start, fmt_name in self._cfg.number_formats.get(table_code, []):
            end = max_col if col_end == "REST" else col_end
            excel_fmt = self._formatMap.get(fmt_name, fmt_name)
            for col in range(col_start, min(end, max_col) + 1):
                col_fmt[col] = (row_start, excel_fmt)

        # Distinct styles, each as a (no-border, with-border) pair. The
        # scratch cell is inside the data region, so the stamping loop below
        # overwrites its style like any other cell — no cleanup needed.
        scratch = ws.cell(row=first_row, column=1)
        scratch.font = self.font
        scratch.alignment = self._centerAlign
        scratch.number_format = self.defaultValueFormat
        border_id = self.wb._borders.add(self._thinBorder)

        def _pair(style_array):
            plain = copy(style_array)
            bordered = copy(style_array)
            bordered.borderId = border_id
            return plain, bordered

        base_pair = _pair(scratch._style)
        fmt_pairs = {}
        for _, excel_fmt in col_fmt.values():
            if excel_fmt not in fmt_pairs:
                scratch.number_format = excel_fmt
                fmt_pairs[excel_fmt] = _pair(scratch._style)

        col_styles = []
        for col in range(1, max_col + 1):
            row_start, excel_fmt = col_fmt.get(col, (None, None))
            col_styles.append((row_start, fmt_pairs.get(excel_fmt), base_pair))

        for row_cells in ws.iter_rows(min_row=first_row, max_row=max_row, max_col=max_col):
            row = row_cells[0].row
            for cell, (row_start, fmt_pair, base) in zip(row_cells, col_styles):
                plain, bordered = fmt_pair if (row_start is not None and row >= row_start) else base
                cell._style = bordered if cell.value is not None else plain

    # ==========================================================================
    #  INDEX + GETTERS
    # ==========================================================================

    def removeBlankSheets(self) -> list:
        """
        Drop every sheet none of whose tables has a value, so the workbook
        (and so the PDF and Index) never shows a title + bare header row for
        a table a state simply doesn't have. A sheet is blank when:
          - its generated table(s) had no data ("empty"), or its body is
            hand-built ("custom") and every labeled block it tried to add
            was empty; AND
          - it was not given any content a page builds some other way.
        Sheets of unknown shape are always kept. Returns the removed names.
        """
        removed = []
        for ws in list(self.wb.worksheets):
            if ws.title == "Index":
                continue
            st = getattr(ws, "_bop_blank", None)
            if not st:
                continue
            base = st.get("base")
            total, kept = st.get("extra_total", 0), st.get("extra_kept", 0)
            if base == "empty":
                # kept == 0: no labeled block added data; the cell count
                # guard: nothing else was written into the sheet after it
                # was generated (a postFormat that fills in its own content).
                blank = kept == 0 and _filled_cells(ws) == st.get("cells")
            elif base == "custom":
                blank = total > 0 and kept == 0
            else:
                blank = False
            if blank:
                removed.append(ws.title)
                self.wb.remove(ws)
        if removed:
            print(f"[ExcelSettingsBOP] {self.state} {self.programName}: omitted {len(removed)} blank sheet(s): {', '.join(removed)}")
        self.removedSheets = removed
        return removed

    def createIndex(self) -> None:
        self.removeBlankSheets()
        ws_index = self.wb["Index"]
        ws_index["A1"] = f"{self.state} - INDEX"
        ws_index["A1"].font = self.fontBold

        sheet_names = self.wb.sheetnames
        for i in range(1, len(sheet_names)):
            name = sheet_names[i]
            row = i + 2
            cell = ws_index[f"A{row}"]
            target_title = self.wb[name]["A1"].value
            cell.value = target_title if target_title else name
            cell.hyperlink = f"#'{name}'!A1"
            cell.font = self.font
            cell.style = "Hyperlink"

        ws_index.column_dimensions["A"].width = 100

    def getWB(self):
        return self.wb

    def getFontName(self):
        return self.fontName

    def getFontSize(self):
        return self.fontSize
