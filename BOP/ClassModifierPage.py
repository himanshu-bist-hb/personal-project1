# This module builds and formats the Class Modifier State Page workbook (BOP).
#
# Ported from the root-level ClassModifierPage.py, same porting pattern as
# [[bop_wholesale_port]]/[[bop_rating_plans_port]]: the [[bop_nesting_order]]
# fix to buildDataFrame (lower-level company -> NGIC -> CW, not NGIC-first),
# tables built via ExcelSettingsBOP.Excel.generateWorksheet + a postFormat
# callback, same sheetSpecs convention every other BOP program page uses.
#
# Like Optional Coverages and Rating Plans, Class Modifier has no
# "2.0"/"pre2.0" split at all -- there is deliberately no
# ClassModifierPageCurrent.py. BOP/BOPRatePages.py's dispatch builds this
# one class regardless of `version`.
#
# The one real duplication in the root source is collapsed here instead of
# being ported as-is: buildPropertyLiabClassModifiers() repeated the exact
# same "exclude the 5 liability-only peril codes" query 3 times (once each
# for Building/BPP/Business Income), and the Liability branch re-listed
# those same 5 codes again (plus AllPeril) as an OR chain. Both are now
# derived from one class-level LIABILITY_ONLY_PERILS tuple.

import pandas as pd
from openpyxl.utils import get_column_letter

from . import ExcelSettingsBOP


class ClassModifier:
    # Peril codes that only ever appear on the Liability class-modifier
    # table, never on Building/BPP/Business Income. Used both to exclude
    # them from the property-side tables and to build the Liability table's
    # inclusion filter (which is just this set plus "AllPeril").
    LIABILITY_ONLY_PERILS = ('L-OtherMed', 'L-OtherPrem', 'L-Products', 'L-SlipFall', 'L-Violence')

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

    # Builds the property & liability class modifiers table for the given
    # coverage ('building', 'bpp', 'business income', 'liability' or 'eb').
    # Returns a dataframe
    def buildPropertyLiabClassModifiers(self, coverage):
        coverage = coverage.casefold()

        if coverage == 'eb':
            ebClassModifier = self.buildDataFrame("BP7_EBClassModifier")
            return ebClassModifier.rename(columns={'Classcode': 'Code', 'EBClassModifier': 'Modifier'})

        propertyLiabClassModifiers = self.buildDataFrame("BP7_Peril_Class_Codes")
        filtered = propertyLiabClassModifiers.query(f'`Peril TypeCode` in {self.perils}'). \
                replace({'Peril TypeCode': self.perilsConversions}).rename(columns={'BuildingClassCode': 'Code'})

        if coverage in ('building', 'bpp', 'business income'):
            filtered = filtered[~filtered['Peril TypeCode'].isin(self.LIABILITY_ONLY_PERILS)]
            value_col = {'building': 'BuildingClassFactor', 'bpp': 'BPPClassFactor', 'business income': 'BIClassFactor'}[coverage]
            return filtered.pivot(index='Code', columns='Peril TypeCode', values=value_col).reset_index('Code')
        elif coverage == 'liability':
            filtered = filtered[filtered['Peril TypeCode'].isin(('AllPeril',) + self.LIABILITY_ONLY_PERILS)]
            return filtered.pivot(index='Code', columns='Peril TypeCode', values='GLClassFactor').reset_index('Code')

    # Converts the given pixels to inches (same 7px/char ratio ExcelSettingsBOP
    # uses internally for its own Table Layout column widths).
    def pixelsToInches(self, px):
        return px / 7.0

    # Applies the shared Code-column number format + peril-column widths
    # every Property & Liability Class Modifier worksheet needs. The EB
    # worksheet has no peril columns, so it gets a single fixed-width column
    # instead of one per peril.
    def formatPropertyLiabClass(self, ws):
        for row in range(4, ws.max_row + 1):
            ws['A' + str(row)].number_format = '####0'
        if ws.title == 'CLEB':
            ws.column_dimensions['B'].width = self.pixelsToInches(80)
        else:
            for col in range(2, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(col)].width = self.pixelsToInches(53)

    # Sets up the Class Modifier Excel file and creates a separate worksheet
    # for each of the given dataframes, then applies each table's shared
    # postFormat. progress_callback (optional) is called with a short
    # message before each sheet is built.
    # Returns the Excel workbook
    def buildClassModifierPage(self, progress_callback=None):
        companies = [c for c in self.rateTables.keys() if c != 'CW']

        ClassModifier = ExcelSettingsBOP.Excel(state=self.state, programName='Class', nEffective=self.nEffective, rEffective=self.rEffective, companyList=companies)

        sheetSpecs = [
            ('CLBG', 'Table 3.C. Property & Liability Class Modifiers - Building', lambda: self.buildPropertyLiabClassModifiers('Building'), False, True, None, self.formatPropertyLiabClass),
            ('CLPP', 'Table 3.C. Property & Liability Class Modifiers - BPP', lambda: self.buildPropertyLiabClassModifiers('BPP'), False, True, None, self.formatPropertyLiabClass),
            ('CLBI', 'Table 3.C. Property & Liability Class Modifiers - Bus Inc', lambda: self.buildPropertyLiabClassModifiers('Business Income'), False, True, None, self.formatPropertyLiabClass),
            ('CLGL', 'Table 3.C. Property & Liability Class Modifiers - Liability', lambda: self.buildPropertyLiabClassModifiers('Liability'), False, True, None, self.formatPropertyLiabClass),
            ('CLEB', 'Table 3.C. Property & Liability Class Modifiers - EB', lambda: self.buildPropertyLiabClassModifiers('EB'), False, True, None, self.formatPropertyLiabClass),
        ]

        total = len(sheetSpecs)
        for i, (tableCode, title, build, useIndex, useHeader, layoutKey, postFormat) in enumerate(sheetSpecs, start=1):
            if progress_callback:
                progress_callback(f"Building sheet {i}/{total}: {tableCode}...")
            print(f"  [{i}/{total}] Building sheet: {tableCode}")
            ws = ClassModifier.generateWorksheet(tableCode, title, build(), useIndex, useHeader, layout_key=layoutKey)
            if postFormat:
                postFormat(ws)

        if progress_callback:
            progress_callback("Building Index sheet...")
        print(f"  [{total}/{total}] Building sheet: Index")
        ClassModifier.createIndex()
        return ClassModifier.getWB()
