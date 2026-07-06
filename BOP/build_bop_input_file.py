"""
build_bop_input_file.py
========================
One-time generator for "BOP/BOP Input File.xlsx" — the non-technical
person's config surface for the BOP (Business Owners Policy) LOB.

Every value written here is transcribed VERBATIM from the values that used
to be hardcoded in the root-level StatePageGenerator.py and AllProgramsPage.py
(pre-refactor). Running this script reproduces that same starting point as
an editable Excel workbook instead of buried Python literals.

Run this once to create the file:
    python BOP/build_bop_input_file.py

It is safe to re-run later to reset every tab back to these defaults —
but note that will overwrite any hand edits made directly in the workbook.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

OUT_PATH = Path(__file__).parent / "BOP Input File.xlsx"

HEADER_FONT = Font(bold=True)


def _write_table(ws, headers, rows):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = HEADER_FONT
    for row in rows:
        ws.append(row)
    for col_cells in ws.columns:
        width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(width + 2, 10), 60)


def build():
    wb = Workbook()
    wb.remove(wb.active)

    # =======================================================================
    # Formatting Defaults — fonts / margins / number formats / borders.
    # This is what makes header/sub-header/title/table/footer sizing
    # consistent across every BOP All Programs page: every sheet reads these
    # same values instead of each format method picking its own.
    # =======================================================================
    ws = wb.create_sheet("Formatting Defaults")
    _write_table(ws, ["Setting", "Value"], [
        ["FontName", "Arial"],
        ["FontSize", 10],
        ["HeaderFontName", "Arial"],
        ["HeaderFontSize", 10],
        ["FooterFontName", "Arial"],
        ["FooterFontSize", 10],
        ["LeftMargin", 0.25],
        ["RightMargin", 0.25],
        ["TopMargin", 1.25],
        ["BottomMargin", 0.95],
        ["HeaderMargin", 0.5],
        ["FooterMargin", 0.25],
        ["BorderColor", "C1C1C1"],
        ["CurrencyFormat", "$#,##0"],
        ["NoDecimalFormat", "#,##0"],
        ["ZipCodeFormat", "####0"],
        ["PrintTitleRows", "1:3"],
        ["PrintTitleRowsWithSubHeader", "1:4"],
    ])

    # =======================================================================
    # Header Footer Text — fixed boilerplate. {tokens} are substituted at
    # runtime with the real state/date/company values.
    # NOTE: this content is new (BOP never had a working ExcelSettings module
    # before, so there was nothing to transcribe) — edit these to the real
    # legal text your filings require.
    # =======================================================================
    ws = wb.create_sheet("Header Footer Text")
    _write_table(ws, ["Field", "Value"], [
        ["HeaderLeftText", "Commercial Lines Manual: Division Nine - Businessowners"],
        ["HeaderCenterTemplate", "\n\n{state} - Rate Pages"],
        ["HeaderRightTemplate", "Effective:\nNew: {n_effective}\nRenewal: {r_effective}"],
        ["FooterLeftTemplate", "{companies}"],
        ["FooterCenterTemplate", "{state_abb} - &[Tab] - &P "],
        ["FooterRightTemplate", ""],
    ])

    # =======================================================================
    # Table Layout — column widths per table code (pixels, converted to
    # inches the same way the old pixelsToInches() helper did: px / 7).
    # ColEnd may be "REST" meaning "this column through the last column".
    # Replaces the ~30 near-duplicate format*() methods' hardcoded widths.
    # =======================================================================
    ws = wb.create_sheet("Table Layout")
    _write_table(ws, ["TableCode", "ColStart", "ColEnd", "WidthPx"], [
        ["SPR", 1, 1, 82], ["SPR", 2, 2, 68], ["SPR", 3, 3, 47],
        ["PCBG", 1, 1, 131], ["PCBG", 2, "REST", 53],
        ["PCPP", 1, 1, 131], ["PCPP", 2, "REST", 53],
        ["PCBI", 1, 1, 131], ["PCBI", 2, "REST", 53],
        ["MVBG", 1, 1, 138], ["MVBG", 2, "REST", 53],
        ["MVPP", 1, 1, 138], ["MVPP", 2, "REST", 53],
        ["BV", 1, 1, 229], ["BV", 2, 2, 54],
        ["AIBI", 1, 1, 208], ["AIBI", 2, 2, 54],
        ["PD", 1, 3, 74], ["PD", 4, 4, 105], ["PD", 5, "REST", 53],
        ["PDH", 1, 3, 74], ["PDH", 4, 4, 105], ["PDH", 5, "REST", 53],
        ["WHOBG", 1, 2, 74], ["WHOBG", 3, 3, 105], ["WHOBG", 4, 4, 105], ["WHOBG", 5, "REST", 62],
        ["WHOBGH", 1, 2, 74], ["WHOBGH", 3, 3, 105], ["WHOBGH", 4, 4, 105], ["WHOBGH", 5, "REST", 62],
        ["WHOPP", 1, 2, 74], ["WHOPP", 3, 3, 105], ["WHOPP", 4, 4, 105], ["WHOPP", 5, "REST", 62],
        ["WHOPPH", 1, 2, 74], ["WHOPPH", 3, 3, 105], ["WHOPPH", 4, 4, 105], ["WHOPPH", 5, "REST", 62],
        ["WHBBG", 1, 1, 222], ["WHBBG", 2, 2, 159], ["WHBBG", 3, "REST", 70],
        ["WHBBGH", 1, 1, 222], ["WHBBGH", 2, 2, 159], ["WHBBGH", 3, "REST", 70],
        ["WHBPP", 1, 1, 222], ["WHBPP", 2, 2, 159], ["WHBPP", 3, "REST", 70],
        ["WHBPPH", 1, 1, 222], ["WHBPPH", 2, 2, 159], ["WHBPPH", 3, "REST", 70],
        ["WHPBG", 1, 1, 159], ["WHPBG", 2, 2, 105], ["WHPBG", 3, "REST", 70],
        ["WHPBGH", 1, 1, 159], ["WHPBGH", 2, 2, 105], ["WHPBGH", 3, "REST", 70],
        ["WHPPP", 1, 1, 159], ["WHPPP", 2, 2, 105], ["WHPPP", 3, "REST", 70],
        ["WHPPPH", 1, 1, 159], ["WHPPPH", 2, 2, 105], ["WHPPPH", 3, "REST", 70],
        ["BA", 1, 1, 82], ["BA", 2, 2, 166], ["BA", 3, 3, 96],
        ["CSFA", 1, "REST", 91],
        ["BABG", 1, 1, 145], ["BABG", 2, "REST", 53],
        ["BAPP", 1, 1, 145], ["BAPP", 2, "REST", 53],
        ["BABI", 1, 1, 145], ["BABI", 2, "REST", 53],
        ["AIBG", 1, 2, 82], ["AIBG", 3, "REST", 53],
        ["AIPP", 1, 2, 82], ["AIPP", 3, "REST", 53],
        ["BCEG_MULTI", 1, 1, 73], ["BCEG_MULTI", 2, 2, 66], ["BCEG_MULTI", 3, 3, 82], ["BCEG_MULTI", 4, "REST", 53],
        ["BCEG_SINGLE", 1, 1, 82], ["BCEG_SINGLE", 2, "REST", 53],
        ["TIB", 1, 1, 82], ["TIB", 2, 2, 68],
        ["EBL", 1, "REST", 180],
        ["EBD", 1, 1, 145], ["EBD", 2, 2, 68],
        ["MD", 1, 1, 180], ["MD", 2, 2, 68],
        ["TR", 1, 1, 100], ["TR", 2, 2, 150], ["TR", 3, 6, 80],
        # Pre-2.0 uses a DIFFERENT column-width profile than BP-2.0 for these
        # same two sheet codes (an extra column shifts the boundaries by one),
        # so AllProgramsPageCurrent.py passes layout_key="..._CURRENT" instead
        # of reusing the BP-2.0 rows above.
        ["WHOBG_CURRENT", 1, 3, 74], ["WHOBG_CURRENT", 4, 4, 105], ["WHOBG_CURRENT", 5, "REST", 62],
        ["WHOPP_CURRENT", 1, 3, 74], ["WHOPP_CURRENT", 4, 4, 105], ["WHOPP_CURRENT", 5, "REST", 62],
        ["WHPBG_CURRENT", 1, 1, 159], ["WHPBG_CURRENT", 2, "REST", 70],
        ["WHPPP_CURRENT", 1, 1, 159], ["WHPPP_CURRENT", 2, "REST", 70],
    ])

    # =======================================================================
    # Number Formats — currency / no-decimal overrides per table.
    # ColEnd may be "REST". RowStart = first data row the format applies to.
    # =======================================================================
    ws = wb.create_sheet("Number Formats")
    _write_table(ws, ["TableCode", "ColStart", "ColEnd", "RowStart", "Format"], [
        ["PD", 1, 4, 5, "Currency"],
        ["PDH", 1, 4, 5, "Currency"],
        ["WHOBG", 1, 4, 5, "Currency"],
        ["WHOBGH", 1, 4, 5, "Currency"],
        ["WHOPP", 1, 4, 5, "Currency"],
        ["WHOPPH", 1, 4, 5, "Currency"],
        ["WHBBG", 1, 2, 4, "Currency"],
        ["WHBBGH", 1, 2, 4, "Currency"],
        ["WHBPP", 1, 2, 4, "Currency"],
        ["WHBPPH", 1, 2, 4, "Currency"],
        ["BABG", 1, 1, 4, "NoDecimal"],
        ["BAPP", 1, 1, 4, "NoDecimal"],
        ["BABI", 1, 1, 4, "NoDecimal"],
        ["AIBG", 1, 2, 5, "NoDecimal"],
        ["AIPP", 1, 2, 5, "NoDecimal"],
        ["EBL", 1, 2, 5, "NoDecimal"],
        ["EBD", 1, 1, 4, "Currency"],
        ["MD", 1, 1, 4, "NoDecimal"],
        ["WHOBG_CURRENT", 1, 4, 5, "Currency"],
        ["WHOPP_CURRENT", 1, 4, 5, "Currency"],
    ])

    # =======================================================================
    # Sub Headers — the "insert a merged label row above the column headers"
    # pattern. Up to two labeled/merged column groups per table.
    # =======================================================================
    ws = wb.create_sheet("Sub Headers")
    _write_table(ws, ["TableCode", "InsertAtRow", "PrintTitleRows",
                       "Label1Range", "Label1Text", "Label2Range", "Label2Text"], [
        ["PD", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", ""],
        ["PDH", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", ""],
        ["WHOBG", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", "Wind-Hail Deductible"],
        ["WHOBGH", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", "Wind-Hail Deductible"],
        ["WHOPP", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", "Wind-Hail Deductible"],
        ["WHOPPH", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", "Wind-Hail Deductible"],
        ["AIBG", 3, "1:4", "A:B", "Building Limit", "C:REST", ""],
        ["AIPP", 3, "1:4", "A:B", "Building Limit", "C:REST", ""],
        ["BCEG_SINGLE", 3, "1:4", "B:REST", "Entire State", "", ""],
        ["EBL", 3, "1:4", "A:B", "Total Property Limit", "", ""],
        ["WHOBG_CURRENT", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", "Wind-Hail Deductible"],
        ["WHOPP_CURRENT", 3, "1:4", "B:D", "Amount of Insurance", "E:REST", "Wind-Hail Deductible"],
    ])

    # =======================================================================
    # Footnotes — one-off cell text unrelated to the column/subheader system.
    # =======================================================================
    ws = wb.create_sheet("Footnotes")
    _write_table(ws, ["TableCode", "Cell", "Text"], [
        ["AIBI", "A16", "For each additional 1%, add 0.005"],
    ])

    # =======================================================================
    # Page Break Rules — sheet-name-prefix -> rule. "*" is the default rule
    # applied to every sheet that no more specific prefix matches.
    # Seeded minimally (no BOP print samples yet to know what special rules
    # are needed) — add rows here as real print issues surface, no Python
    # required.
    # =======================================================================
    ws = wb.create_sheet("Page Break Rules")
    _write_table(ws, ["SheetPrefix", "Rule"], [
        ["Index", "index"],
        ["*", "fit_single_page"],
    ])

    # =======================================================================
    # Perils By State — replaces the 4 elif-state-in-(...) blocks.
    # NOTE: states not listed here (e.g. AK, HI, LA, OK) were not supported
    # by the original BOP script either — add a row for them here whenever
    # BOP expands to those states.
    # =======================================================================
    ws = wb.create_sheet("Perils By State")
    full_18 = "allother1,cat1,cat2,cat3,cat4,fire1,fire2,fire3,fire4,liability1,liability2,liability3,liability4,theft1,water1,water2,weather1,weather2"
    no_cat3 = "allother1,cat1,cat2,cat4,fire1,fire2,fire3,fire4,liability1,liability2,liability3,liability4,theft1,water1,water2,weather1,weather2"
    no_fire2 = "allother1,cat1,cat2,cat3,cat4,fire1,fire3,fire4,liability1,liability2,liability3,liability4,theft1,water1,water2,weather1,weather2"
    no_cat3_no_fire2 = "allother1,cat1,cat2,cat4,fire1,fire3,fire4,liability1,liability2,liability3,liability4,theft1,water1,water2,weather1,weather2"
    rows = [["TX", full_18]]
    for s in ["AZ", "CA", "CO", "ID", "MT", "NM", "NV", "OR", "UT", "WA", "WY"]:
        rows.append([s, no_cat3])
    for s in ["AL", "AR", "CT", "DC", "DE", "FL", "GA", "IL", "IN", "KY", "MA", "MD", "ME",
              "MO", "MS", "NC", "NH", "NJ", "NY", "OH", "PA", "RI", "SC", "TN", "VA", "VT", "WV"]:
        rows.append([s, no_fire2])
    for s in ["IA", "KS", "MI", "MN", "ND", "NE", "SD", "WI"]:
        rows.append([s, no_cat3_no_fire2])
    _write_table(ws, ["State", "Perils"], rows)

    # =======================================================================
    # Peril Conversions — internal peril code -> display name.
    # =======================================================================
    ws = wb.create_sheet("Peril Conversions")
    _write_table(ws, ["PerilCode", "DisplayName"], [
        ["allother1", "NW-Other"], ["allperil", "AllPeril"], ["cat1", "ST"], ["cat2", "WS"],
        ["cat3", "HU"], ["cat4", "L-Products"], ["fire1", "NW-Fire"], ["fire2", "WF"],
        ["fire3", "FFEQ"], ["fire4", "NC-BINC"], ["liability1", "L-SlipFall"],
        ["liability2", "L-Violence"], ["liability3", "L-OtherMed"], ["liability4", "L-OtherPrem"],
        ["theft1", "NW-Theft"], ["water1", "NW-Water"], ["water2", "NC-Water"],
        ["weather1", "NC-Other"], ["weather2", "NC-Wind"],
    ])

    # =======================================================================
    # Protection Class Conversions — strips excess leading zeros.
    # =======================================================================
    ws = wb.create_sheet("Protection Class Conversions")
    _write_table(ws, ["Code", "DisplayValue"], [
        ["000001", "1"], ["000002", "2"], ["000003", "3"], ["000004", "4"], ["000005", "5"],
        ["000006", "6"], ["000007", "7"], ["000008", "8"], ["000009", "9"], ["000010", "10"],
        ["00001X", "1X"], ["00002X", "2X"], ["00003X", "3X"], ["00004X", "4X"], ["00005X", "5X"],
        ["00006X", "6X"], ["00007X", "7X"], ["00008X", "8X"],
        ["00001Y", "1Y"], ["00002Y", "2Y"], ["00003Y", "3Y"], ["00004Y", "4Y"], ["00005Y", "5Y"],
        ["00006Y", "6Y"], ["00007Y", "7Y"], ["00008Y", "8Y"],
        ["00001W", "1W"], ["00002W", "2W"], ["00003W", "3W"], ["00004W", "4W"], ["00005W", "5W"],
        ["00006W", "6W"], ["00007W", "7W"], ["00008W", "8W"],
        ["00008B", "8B"], ["00009E", "9E"], ["00009S", "9S"], ["00010W", "10W"],
    ])

    # =======================================================================
    # Building Codes By State — states with more than 1 BCEG group.
    # Codes is a comma-separated list of raw territory codes for that group.
    # =======================================================================
    ws = wb.create_sheet("Building Codes By State")
    _write_table(ws, ["State", "Group", "Codes"], [
        ["AL", "A", "001"], ["AL", "B", "004"], ["AL", "C", "005"], ["AL", "D", "006"],
        ["FL", "A", "011,012"], ["FL", "B", "010,015"], ["FL", "C", "002,007,008,014,016,017"], ["FL", "D", "009,013"],
        ["GA", "A", "002"], ["GA", "B", "004"], ["GA", "C", "005"], ["GA", "D", "006"],
        ["MS", "A", "002"], ["MS", "B", "003"], ["MS", "C", "004"],
        ["NC", "A", "003"], ["NC", "B", "004"], ["NC", "C", "005"], ["NC", "D", "006"],
        ["NE", "A", "701"], ["NE", "B", "703"], ["NE", "C", "704"],
        ["SC", "A", "002"], ["SC", "B", "003"], ["SC", "C", "004"],
        ["TX", "A", "004,005,006,007,008,009,015,016"], ["TX", "B", "010,011,012,013,014"],
        ["VA", "A", "001,005,006,007,008,009,012,013"], ["VA", "B", "010,011"],
        ["WY", "A", "702"], ["WY", "B", "703"],
    ])

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    build()
