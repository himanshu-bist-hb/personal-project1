# This module builds and formats the Common Rules State Page workbook (BOP).
#
# Ported from the root-level CommonRulesPage.py, same porting pattern as
# [[bop_wholesale_port]]/[[bop_class_modifier_port]]: the [[bop_nesting_order]]
# fix to buildDataFrame (lower-level company -> NGIC -> CW, not NGIC-first),
# tables built via ExcelSettingsBOP.Excel.generateWorksheet + a postFormat
# callback, same sheetSpecs convention every other BOP program page uses.
#
# Like Optional Coverages/Rating Plans/Class Modifier, Common Rules has no
# "2.0"/"pre2.0" split at all -- there is deliberately no
# CommonRulesPageCurrent.py. BOP/BOPRatePages.py's dispatch builds this one
# class regardless of `version`.
#
# The root source's own formatFixedExpense() set ws.column_dimensions['A'].width
# twice (once to 70px, then unconditionally again to 100px on every data
# row) -- a redundant repeated assignment, not two different states -- so
# it's set once here.

import pandas as pd

from . import ExcelSettingsBOP


class CommonRules:
    def __init__(self, state, rateTables, perils, perilsConversions, nEffective, rEffective) -> None:
        self.state = state
        self.rateTables = rateTables
        self.perils = perils
        self.perilsConversions = perilsConversions
        self.nEffective = nEffective  # New business effective date
        self.rEffective = rEffective  # Renewal business effective date

    # Builds a dataframe for the given table code. Hierarchy matches every
    # other BOP program page: lower-level company (NACO/NAFF/NICOF) first,
    # then NGIC (state-level default), then CW as the country-wide fallback.
    # See [[bop_nesting_order]] — the root file checked NGIC first.
    def buildDataFrame(self, tableCode):
        if 'NACO' in self.rateTables.keys():
            if tableCode in self.rateTables['NACO'].keys():
                return pd.DataFrame(data=self.rateTables['NACO'][tableCode][1:], index=None, columns=self.rateTables['NACO'][tableCode][0])
        if 'NAFF' in self.rateTables.keys():
            if tableCode in self.rateTables['NAFF'].keys():
                return pd.DataFrame(data=self.rateTables['NAFF'][tableCode][1:], index=None, columns=self.rateTables['NAFF'][tableCode][0])
        if 'NICOF' in self.rateTables.keys():
            if tableCode in self.rateTables['NICOF'].keys():
                return pd.DataFrame(data=self.rateTables['NICOF'][tableCode][1:], index=None, columns=self.rateTables['NICOF'][tableCode][0])
        if tableCode in self.rateTables['NGIC'].keys():
            return pd.DataFrame(data=self.rateTables['NGIC'][tableCode][1:], index=None, columns=self.rateTables['NGIC'][tableCode][0])
        return pd.DataFrame(data=self.rateTables['CW'][tableCode][1:], index=None, columns=self.rateTables['CW'][tableCode][0])

    # Builds the Expense Constant table.
    # Returns a dataframe
    def buildExpenseConstant(self):
        expenseConstant = self.buildDataFrame("BP7PerilCWSpecificCovRateFactors")
        return expenseConstant.query('CoverageSpecific == "Expense Constant"').filter(items=['Value']).rename(columns={'Value': 'Rate'})

    # Converts the given pixels to inches (same 7px/char ratio ExcelSettingsBOP
    # uses internally for its own Table Layout column widths).
    def pixelsToInches(self, px):
        return px / 7.0

    # Applies the EC worksheet's column width + currency format.
    def formatFixedExpense(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(100)
        for row in range(4, ws.max_row + 1):
            ws['A' + str(row)].number_format = '$#,##0'

    # Sets up the Common Rules Excel file and creates a worksheet for the
    # Expense Constant table. progress_callback (optional) is called with a
    # short message before the sheet is built.
    # Returns the Excel workbook
    def buildCommonRulesPage(self, progress_callback=None):
        companies = [c for c in self.rateTables.keys() if c != 'CW']

        CommonRules = ExcelSettingsBOP.Excel(state=self.state, programName='Common Rules', nEffective=self.nEffective, rEffective=self.rEffective, companyList=companies)

        sheetSpecs = [
            ('EC', 'CR Table 8. Expense Constant', self.buildExpenseConstant, False, True, None, self.formatFixedExpense),
        ]

        total = len(sheetSpecs)
        for i, (tableCode, title, build, useIndex, useHeader, layoutKey, postFormat) in enumerate(sheetSpecs, start=1):
            if progress_callback:
                progress_callback(f"Building sheet {i}/{total}: {tableCode}...")
            print(f"  [{i}/{total}] Building sheet: {tableCode}")
            ws = CommonRules.generateWorksheet(tableCode, title, build(), useIndex, useHeader, layout_key=layoutKey)
            if postFormat:
                postFormat(ws)

        if progress_callback:
            progress_callback("Building Index sheet...")
        print(f"  [{total}/{total}] Building sheet: Index")
        CommonRules.createIndex()
        return CommonRules.getWB()
