# This module builds and formats the Rating Plans State Page workbook (BOP).
#
# Ported from the root-level RatingPlansPage.py, same porting pattern as
# [[bop_wholesale_port]]/[[bop_optional_coverages_port]]: the [[bop_nesting_order]]
# fix to buildDataFrame (lower-level company -> NGIC -> CW, not NGIC-first),
# tables built via ExcelSettingsBOP.Excel.generateWorksheet + a postFormat
# callback for the bespoke merged-header layouts, same sheetSpecs convention
# every other BOP program page already uses.
#
# Like Optional Coverages, Rating Plans has no "2.0"/"pre2.0" split at all —
# there is deliberately no RatingPlansPageCurrent.py. BOP/BOPRatePages.py's
# dispatch builds this one class regardless of `version`.
#
# Two real duplications in the root source were collapsed here instead of
# being ported as-is (the user's explicit ask — one class, no copy-pasted
# logic per near-identical case):
#   - buildMultiBuildingCredit() built the same table 4 times (All Other /
#     Habitational x non-NRP / NRP) via 4 copy-pasted blocks -> now one
#     parametrized _buildMultiBuildingBlock() called 4 times and merged.
#   - buildTieringFactor() built the same table 3 times (Hab / AS-FS-W /
#     O-R-S grade bands) via 3 copy-pasted blocks -> now one parametrized
#     _buildTieringBlock() called 3 times and concatenated.
#   - buildLPDPFactors() repeated the same 17-entry premium-range label dict
#     twice (once per NRP/non-NRP table, differing only in the $0-$5,000
#     bucket's source key) -> now one shared PREMIUM_RANGE_LABELS dict.
#
# The old desktop tool's RatingPlansApplies checkbox and the IRPM
# Credit/Debit Tkinter DoubleVars are dropped from the constructor (program
# selection already gates whether this page builds at all, same reasoning
# as dropping OptionalCoveragesApplies in the Optional Coverages port).
# IRPM Credit/Debit are plain floats defaulting to 0.0 — app.py's IRPM
# inputs are still disabled placeholders ("coming soon, has no effect yet"),
# so the State IRPM Modification Plan table (RPMP) renders as 0%/0% until
# that UI is wired up; that's a separate, not-yet-requested task.

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter

from . import ExcelSettingsBOP


class RatingPlans:
    # Shared by buildLPDPFactors' two tables — identical except for the
    # $0-$5,000 bucket's source key (0 for the non-NRP table, 1 for NRP).
    PREMIUM_RANGE_LABELS = {
        5001: "$5,001 to $6,000", 6001: "$6,001 to $8,000", 8001: "$8,001 to $10,000",
        10001: "$10,001 to $15,000", 15001: "$15,001 to $20,000", 20001: "$20,001 to $25,000",
        25001: "$25,001 to $30,000", 30001: "$30,001 to $35,000", 35001: "$35,001 to $40,000",
        40001: "$40,001 to $45,000", 45001: "$45,001 to $50,000", 50001: "$50,001 to $75,000",
        75001: "$75,001 to $100,000", 100001: "$100,001 to $150,000", 150001: "$150,001 to $200,000",
        200001: "$200,001 to $250,000", 250001: "$250,001 and greater",
    }

    def __init__(self, state, rateTables, perils, perilsConversions, nEffective, rEffective,
                 irpmCredit=0.0, irpmDebit=0.0) -> None:
        self.state = state
        self.rateTables = rateTables
        self.perils = perils
        self.perilsConversions = perilsConversions
        self.nEffective = nEffective  # New business effective date
        self.rEffective = rEffective  # Renewal business effective date
        self.irpmCredit = irpmCredit
        self.irpmDebit = irpmDebit

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

    # One "Total # of Buildings" x Multi_Building_Credit_Factor block, filtered
    # to the given class code and renamed to the given output column. Called
    # once per (table, class code) combination by buildMultiBuildingCredit.
    def _buildMultiBuildingBlock(self, tableCode, classCodeMin, columnName):
        block = self.buildDataFrame(tableCode).query(f'Class_Code_Min == {classCodeMin} & `Peril TypeCode` == "allperil"'). \
                fillna({'Building_No_Max': 0}).astype({'Building_No_Min': 'int64', 'Building_No_Max': 'int64'}). \
                astype({'Building_No_Min': 'string', 'Building_No_Max': 'string'})
        block['Total # of Buildings'] = np.where(block['Building_No_Max'] == '0',
                                                  block['Building_No_Min'] + '+',
                                                  block['Building_No_Min'] + ' - ' + block['Building_No_Max'])
        return block.filter(items=['Total # of Buildings', 'Multi_Building_Credit_Factor']).rename(columns={'Multi_Building_Credit_Factor': columnName})

    # Builds the Multi Building Credit Factor table (All Other / Habitational,
    # each with a non-NRP and an NRP column) by merging 4 filtered blocks.
    # Returns a dataframe
    def buildMultiBuildingCredit(self):
        allOther    = self._buildMultiBuildingBlock('BP7_Peril_Multi_Building_Credit', 20000, 'All Other')
        hab         = self._buildMultiBuildingBlock('BP7_Peril_Multi_Building_Credit', 10000, 'Habitational')
        allOtherNRP = self._buildMultiBuildingBlock('BP7 Peril Multi Building Credit NRP', 20000, 'All Other (NRP)')
        habNRP      = self._buildMultiBuildingBlock('BP7 Peril Multi Building Credit NRP', 10000, 'Habitational (NRP)')
        nonNRP = pd.merge(allOther, hab, on='Total # of Buildings', how='inner')
        nrp = pd.merge(allOtherNRP, habNRP, on='Total # of Buildings', how='inner')
        return pd.merge(nonNRP, nrp, on='Total # of Buildings', how='inner')

    # One Peril/Grade/Factor block for the given grade band, tagged with the
    # given program label. Called once per program band by buildTieringFactor.
    def _buildTieringBlock(self, grades, programLabel):
        block = self.buildDataFrame("BP7_Peril_Tiering_Factor").query(f'`Peril TypeCode` in {self.perils}'). \
                replace({'Peril TypeCode': self.perilsConversions}).rename(columns={'TierFactor': 'Factor'})
        block = block[block['Grade'].isin(grades)]
        block['Program'] = programLabel
        return block

    # Builds the Risk Tier Rating Plan table across all 3 program grade bands.
    # Returns a dataframe
    def buildTieringFactor(self):
        hab    = self._buildTieringBlock(["1", "2", "3", "4", "5", "6", "7", "8", "9"], 'H')
        asfsw  = self._buildTieringBlock(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"], 'AS/FS/W')
        ors    = self._buildTieringBlock(["K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"], 'O/R/S')
        tieringFactor = pd.concat([hab, asfsw, ors])
        return tieringFactor.sort_values(by='Program', ascending=False).sort_values(by=['Peril TypeCode', 'Grade']). \
                rename(columns={'Peril TypeCode': 'Peril'}).filter(items=['Peril', 'Program', 'Grade', 'Factor'])

    # Builds the State Individual Risk Premium Modification Eligibility
    # Threshold table.
    # Returns a dataframe
    def buildIRPMThreshold(self):
        irpmThreshold = self.buildDataFrame("BP7_IRPM_Eligibility_Threshold")
        return irpmThreshold.query('ProgramCode == "Auto Service"').filter(items=['IRPMEligibleAmount']).rename(columns={'IRPMEligibleAmount': 'Amount'})

    # Builds the State Individual Risk Premium Modification Plan table from
    # the user-entered IRPM credit/debit percentages.
    # Returns a dataframe
    def buildIRPMModPlan(self):
        return pd.DataFrame([[self.irpmCredit, self.irpmDebit]], columns=['Credit', 'Debit'])

    # Builds the Lifetime Expense Allocation Factor table.
    # Returns a dataframe
    def buildLEAFfactors(self):
        leafFactors = self.buildDataFrame("BP7_Peril_Retention_Factor")
        leafFactors = leafFactors[leafFactors['RetentionGrade'].isin(["A", "B", "C", "D", "E", "F"])]
        return leafFactors. \
                filter(items=['Peril TypeCode', 'RetentionGrade', 'RetentionFactor']).replace({'Peril TypeCode': self.perilsConversions}). \
                rename(columns={'Peril TypeCode': 'Peril', 'RetentionGrade': 'Grade', 'RetentionFactor': 'Factor'}). \
                sort_values(by=['Peril', 'Grade'])

    # Builds the Large Premium Discount Plan factor table (All Other +
    # National Retail Program factor ranges, side by side by Annual Premium).
    # Returns a dataframe
    def buildLPDPFactors(self):
        nrp = self.buildDataFrame("BP7 LPDP NRP Factor_v1_Ext").query('ProgramCode == "Auto Service"'). \
                replace({'PremiumRange': {1: "$0 to $5,000", **self.PREMIUM_RANGE_LABELS}})
        filteredNRP = nrp.filter(items=['PremiumRange', 'LPDPFactor']).rename(columns={'PremiumRange': 'Annual Premium', 'LPDPFactor': 'National Retail Program Factor Range'})
        nonNRP = self.buildDataFrame("BP7 LPDP Factor_v2_Ext").query('ProgramCode == "Auto Service"'). \
                replace({'PremiumRange': {0: "$0 to $5,000", **self.PREMIUM_RANGE_LABELS}})
        filteredNonNRP = nonNRP.filter(items=['PremiumRange', 'LPDPFactor']).rename(columns={'PremiumRange': 'Annual Premium', 'LPDPFactor': 'All Other Factor Range'})
        return pd.merge(filteredNonNRP, filteredNRP, on='Annual Premium', how='inner')

    # Converts the given pixels to inches (same 7px/char ratio ExcelSettingsBOP
    # uses internally for its own Table Layout column widths).
    def pixelsToInches(self, px):
        return px / 7.0

    _thinBorder = Border(left=Side(border_style='thin', color='C1C1C1'), right=Side(border_style='thin', color='C1C1C1'),
                          top=Side(border_style='thin', color='C1C1C1'), bottom=Side(border_style='thin', color='C1C1C1'))

    # Applies the MBCP worksheet's 2-row merged group header ("All Other" /
    # "National Retail Program" spanning their non-NRP/NRP column pairs).
    def formatMultiBuildingCredit(self, ws):
        ws.insert_rows(3)
        ws['B2'] = 'MBCP Factor'
        ws['B3'] = 'All Other'
        ws['D3'] = 'National Retail Program'
        for cell in list(ws['2:2']) + list(ws['3:3']):
            cell.border = self._thinBorder
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='bottom', wrap_text=True)
        ws.merge_cells('B2:E2')
        ws.merge_cells('B3:C3')
        ws.merge_cells('D3:E3')
        ws.print_title_rows = '1:4'
        ws.column_dimensions['A'].width = self.pixelsToInches(125)
        for col in range(2, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = self.pixelsToInches(80)
        ws.merge_cells('A2:A4')
        ws['A2'] = 'Total # of Buildings'

    # Applies the RTRP worksheet's column widths + 4-decimal Factor format.
    def formatTieringFactor(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(138)
        for col in range(2, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = self.pixelsToInches(80)
        for row in range(4, ws.max_row + 1):
            ws['D' + str(row)].number_format = '#,##0.0000'

    # Applies the RPMET worksheet's column width + currency format.
    def formatIRPMThreshold(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(70)
        for row in range(4, ws.max_row + 1):
            ws['A' + str(row)].number_format = '$#,##0'

    # Applies the RPMP worksheet's "Total Modification" label row + percent format.
    def formatIRPMModPlan(self, ws):
        ws.insert_rows(3)
        ws.insert_rows(4)
        ws['A3'] = 'Total Modification:'
        for cell in ws['3:3']:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='left', vertical='bottom', wrap_text=False)
        ws.column_dimensions['A'].width = self.pixelsToInches(80)
        ws.column_dimensions['B'].width = self.pixelsToInches(80)
        for row in range(6, ws.max_row + 1):
            ws['A' + str(row)].number_format = '#,##0%'
            ws['B' + str(row)].number_format = '#,##0%'

    # Applies the LPDP worksheet's column widths.
    def formatLPDPFactors(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(200)
        ws.column_dimensions['B'].width = self.pixelsToInches(95)
        ws.column_dimensions['C'].width = self.pixelsToInches(150)

    # Sets up the Rating Plans Excel file and creates a separate worksheet
    # for each of the given dataframes, then applies each table's bespoke
    # postFormat. progress_callback (optional) is called with a short
    # message before each sheet is built.
    # Returns the Excel workbook
    def buildRatingPlansPage(self, progress_callback=None):
        companies = [c for c in self.rateTables.keys() if c != 'CW']

        RatingPlans = ExcelSettingsBOP.Excel(state=self.state, programName='Rating Plans', nEffective=self.nEffective, rEffective=self.rEffective, companyList=companies)

        sheetSpecs = [
            ('MBCP',  'RP Table 90.B. Multiple Building Credit Plan', self.buildMultiBuildingCredit, False, True, None, self.formatMultiBuildingCredit),
            ('RTRP',  'RP Table 91.C. Risk Tier Rating Plan', self.buildTieringFactor, False, True, None, self.formatTieringFactor),
            ('RPMET', 'RP Table 92.A. State Individual Risk Premium Modification Eligibility Threshold', self.buildIRPMThreshold, False, True, None, self.formatIRPMThreshold),
            ('RPMP',  'RP Table 92.C. State Individual Risk Premium Modification Plan', self.buildIRPMModPlan, False, True, None, self.formatIRPMModPlan),
            ('LEAF',  'RP Table 94.C.1. Lifetime Expense Allocation Factor', self.buildLEAFfactors, False, True, None, None),
            ('LPDP',  'RP Table 95.C. Large Premium Discount Plan', self.buildLPDPFactors, False, True, None, self.formatLPDPFactors),
        ]

        total = len(sheetSpecs)
        for i, (tableCode, title, build, useIndex, useHeader, layoutKey, postFormat) in enumerate(sheetSpecs, start=1):
            if progress_callback:
                progress_callback(f"Building sheet {i}/{total}: {tableCode}...")
            print(f"  [{i}/{total}] Building sheet: {tableCode}")
            ws = RatingPlans.generateWorksheet(tableCode, title, build(), useIndex, useHeader, layout_key=layoutKey)
            if postFormat:
                postFormat(ws)

        if progress_callback:
            progress_callback("Building Index sheet...")
        print(f"  [{total}/{total}] Building sheet: Index")
        RatingPlans.createIndex()
        return RatingPlans.getWB()
