# This module builds and formats the Additional Rules State Page workbook (BOP).
#
# Ported from the root-level AdditionalRulesPage.py, same porting pattern as
# [[bop_wholesale_port]]/[[bop_common_rules_port]]: the [[bop_nesting_order]]
# fix to buildDataFrame (lower-level company -> NGIC -> CW, not NGIC-first),
# tables built via ExcelSettingsBOP.Excel.generateWorksheet + a postFormat
# callback, same sheetSpecs convention every other BOP program page uses.
#
# Like Optional Coverages/Rating Plans/Class Modifier/Common Rules,
# Additional Rules has no "2.0"/"pre2.0" split at all -- there is
# deliberately no AdditionalRulesPageCurrent.py. BOP/BOPRatePages.py's
# dispatch builds this one class regardless of `version`.
#
# Unlike the other Additional Rules-adjacent pages, which sheets this page
# builds genuinely depends on the state (MD/OH/AL/IN,KY/WA each get their
# own state-specific rule tables; everyone else gets Rate Capping +
# Distribution Factor) -- that's real business logic, not duplication, and
# is kept as a single per-state sheetSpecs list. What WAS duplication in the
# source, and is collapsed here:
#   - buildSGLIL/buildSGEIL were the same "Increased Limits" query against
#     two different tables -> one _buildStopGapIncreasedLimits(tableCode).
#   - buildSGLPD/buildSGEPD were the same "Premium Determination" query
#     against those same two tables -> one _buildStopGapPremiumDetermination.
#   - Both of the above repeated the same 3-entry "Increased Limits" label
#     dict -> one class-level STOP_GAP_LIMIT_LABELS.
#   - buildMSB/buildMSC differed only by county ("ATHENS"/"ERIE") -> one
#     _buildMineSubsidenceCovRate(county).
#   - buildIBHSHurricane/buildIBHSWindHail differed only by Certificate Type
#     -> one _buildIBHSCertificate(certificateType).
#   - formatMSB/formatMSC, formatSGlil/formatSGeil and formatSGlpd/formatSGepd
#     were byte-for-byte identical method bodies -> one formatter each.
#   - All 8 of the above format methods (plus formatMineSubsidence/formatUPS)
#     repeated the same "loop every column/row and set currency format past
#     some column threshold" nested loop -> one _applyCurrencyFormat(ws, min_col).
#   - buildWHExclusion called `.replace({'Program': self.classCodes})` twice
#     in the same chain (the second call was a no-op) -> called once.
#   - buildDistributionFactors typed out a 16-entry "DG00"->"00" .. "DG15"->
#     "15" dict by hand -> derived with a comprehension.
#
# AL's IBHS sheet (3 tables stacked on one worksheet) is built via
# ExcelSettingsBOP.generateMultiTableWorksheet — the generic replacement for
# the source's lost generateWorksheet3tables (see [[bop_optional_coverages_port]]).
# formatIBHSZones pokes absolute row numbers that assume the exact
# contiguous block layout that helper produces; this could not be checked
# against a real AL ratebook/PDF, same caveat as Optional Coverages' own
# absolute-row formatters.

import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter

from . import ExcelSettingsBOP


class AdditionalRules:
    # Shared by both Stop Gap builders (Increased Limits + Premium
    # Determination), each of which filters BP7[_Extended]_Stop_Gap_Base_Rate
    # by "Limit" and relabels it for display.
    STOP_GAP_LIMIT_LABELS = {
        '100000/100000/500000': '$100,000/100,000/500,000',
        '500000/500000/500000': '$500,000/500,000/500,000',
        '1000000/1000000/1000000': '$1,000,000/1,000,000/1,000,000',
    }

    _thinBorder = Border(left=Side(border_style='thin', color='C1C1C1'), right=Side(border_style='thin', color='C1C1C1'),
                          top=Side(border_style='thin', color='C1C1C1'), bottom=Side(border_style='thin', color='C1C1C1'))

    def __init__(self, state, rateTables, perils, perilsConversions, classCodes, nEffective, rEffective) -> None:
        self.state = state
        self.rateTables = rateTables
        self.perils = perils
        self.perilsConversions = perilsConversions
        self.classCodes = classCodes
        self.nEffective = nEffective  # New business effective date
        self.rEffective = rEffective  # Renewal business effective date

        self.currencyFormat = '$#,##0'
        self.percentageFormat = '#,##0%'

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

    # ── Table builders ──────────────────────────────────────────────────────

    # Builds the Transition/Rate Capping factor table.
    # Returns a dataframe
    def buildRateCappingFactors(self):
        rateCapping = self.buildDataFrame("BP7_RateCappingPremiumRange2")
        return rateCapping.query('RateCapType == "Migration" & FiledRatesReachedIndicator == 0 & YrsOnPCMin == 1'). \
                filter(items=['MinimumRange', 'MaximumRange']).rename(columns={'MinimumRange': 'Lower Bound', 'MaximumRange': 'Upper Bound'})

    # Builds the Distribution Factor table.
    # Returns a dataframe
    def buildDistributionFactors(self):
        distributionGroupLabels = {f'DG{i:02d}': f'{i:02d}' for i in range(16)}  # DG00->00 .. DG15->15
        distributionFactors = self.buildDataFrame("BP7Distribution Factor_Ext")
        return distributionFactors.query('DistributionGroup != "DG99"').rename(columns={'DistributionGroup': 'Distribution Group'}). \
                replace({'Distribution Group': distributionGroupLabels})

    # Builds the Windstorm or Hail Exclusion factor table.
    # Returns a dataframe
    def buildWHExclusion(self):
        whExclusion = self.buildDataFrame("BP7_WHExclusionFactor")
        return whExclusion.rename(columns={'BuildingWHExclusionFactor': 'Building', 'BPPWHExclusionFactor': 'BPP', 'ClassCode_Min': 'Program'}). \
                replace({'Program': self.classCodes}).filter(items=['Program', 'Building', 'BPP']). \
                replace({'Program': {'Hab': 'Habitational', 'Food': 'Food Service', 'Auto': 'Auto Service', 'Service': 'Process/Service'}}).sort_values('Program')

    # Builds the Liability for Hazards of Lead factor table.
    # Returns a dataframe
    def buildLiabilityHazards(self):
        liabilityHazards = self.buildDataFrame("BP7_LiabForHazdOfLead_Factor")
        liabilityHazards = liabilityHazards.fillna({'YearBuilt': 'Buildings built 1979 and later:'}). \
                replace({'YearBuilt': {1979: 'For buildings built prior to 1979:'}}).replace({'HazardOfLeadFactor': {1: 'No Charge'}})
        return liabilityHazards.rename(columns={'YearBuilt': 'Year Built', 'HazardOfLeadFactor': 'General Liability'}).sort_values('Year Built', ascending=False)

    # Builds the IBHS Wind Mitigation Zones table (AL only).
    # Returns a dataframe
    def buildIBHSzones(self):
        ibhsZones = self.buildDataFrame("BP7_Wind_Mitigation_Zone")
        ibhsZones = ibhsZones.rename(columns={'County Name': 'Counties'}). \
                replace({'IBHS Zone': {'Central': 'Central Zone', 'Coastal': 'Coastal Zone', 'Northern': 'Northern Zone'}}).filter(items=['Counties', 'IBHS Zone'])
        ibhsZones = ibhsZones.groupby('IBHS Zone')['Counties'].apply(', '.join).reset_index()
        ibhsZones.loc[ibhsZones['IBHS Zone'] == 'Northern Zone', 'Counties'] = 'All Other Counties'
        return ibhsZones

    # Builds one IBHS Certificate Discount table (Hurricane or High Wind and
    # Hail — the two share the exact same shape, differing only by
    # Certificate Type). Called twice by buildAdditionalRulesPage for AL.
    # Returns a dataframe
    def _buildIBHSCertificate(self, certificateType):
        wm = self.buildDataFrame("BP7_Wind_Mitigation")
        wm = wm[(wm['Sub-Decking'] != "No") & (wm['Sub-Decking'] != "Yes") & (wm['Roof Age Max'] != 999) & (wm['Certificate Type'] == certificateType)]
        wm = wm.pivot(index='IBHS Zone', columns='Certificate Level', values='Factor').reset_index('IBHS Zone'). \
                rename(columns={'Fortified Bronze': 'Bronze', 'Fortified Silver': 'Silver', 'Fortified Gold': 'Gold'})
        wm = wm[['IBHS Zone', '2006+ IBC', 'Bronze', 'Silver', 'Gold']]
        for col in ('2006+ IBC', 'Bronze', 'Silver', 'Gold'):
            wm[col] = 1 - wm[col]
        return wm

    # Builds the Mine Subsidence Insurance table (IN/KY only) — Dwelling and
    # Non-Dwelling structure charges side by side by Amount of Insurance.
    # Returns a dataframe
    def buildMineSubsidence(self):
        mineDwelling = self.buildDataFrame("BP7_Bldg_Mine_Subsidence_Charge_Dwelling").astype({'Amt_Insurance_Min': 'int64', 'Amt_Insurance_Max': 'int64'})
        mineDwelling['Amt_Insurance_Min'] = "$" + mineDwelling['Amt_Insurance_Min'].apply('{:,.0f}'.format)
        mineDwelling['Amt_Insurance_Max'] = "$" + mineDwelling['Amt_Insurance_Max'].apply('{:,.0f}'.format)
        mineDwelling["Amount of Insurance D"] = mineDwelling["Amt_Insurance_Min"] + ' - ' + mineDwelling["Amt_Insurance_Max"]
        mineDwelling = mineDwelling.replace({'Amount of Insurance D': {'$1 - $25,000': 'Up to $25,000'}}).filter(items=['Amount of Insurance D', 'MineSubsidenceChargeDwelling'])
        mineDwelling = mineDwelling.rename(columns={'MineSubsidenceChargeDwelling': 'Dwelling Structure'}).sort_values('Dwelling Structure')

        mineNonDwelling = self.buildDataFrame("BP7_Bldg_Mine_Subsidence_Charge_NonDwelling").astype({'Amt_Insurance_Min': 'int64', 'Amt_Insurance_Max': 'int64'})
        mineNonDwelling['Amt_Insurance_Min'] = "$" + mineNonDwelling['Amt_Insurance_Min'].apply('{:,.0f}'.format)
        mineNonDwelling['Amt_Insurance_Max'] = "$" + mineNonDwelling['Amt_Insurance_Max'].apply('{:,.0f}'.format)
        mineNonDwelling["Amount of Insurance ND"] = mineNonDwelling["Amt_Insurance_Min"] + ' - ' + mineNonDwelling["Amt_Insurance_Max"]
        mineNonDwelling = mineNonDwelling.replace({'Amount of Insurance ND': {'$1 - $25,000': 'Up to $25,000'}}).filter(items=['Amount of Insurance ND', 'MineSubsidenceChargeNonDwelling'])
        mineNonDwelling = mineNonDwelling.rename(columns={'MineSubsidenceChargeNonDwelling': 'Non-Dwelling Structure'}).sort_values('Non-Dwelling Structure')

        mineSubsidence = mineDwelling.join(mineNonDwelling)
        mineSubsidence = mineSubsidence.fillna({'Amount of Insurance D': 'N/A', 'Dwelling Structure': 'N/A', 'Amount of Insurance ND': 'N/A', 'Non-Dwelling Structure': 'N/A'})
        return mineSubsidence.replace({'Amount of Insurance ND': {'N/A': ' '}}).replace({'Non-Dwelling Structure': {'N/A': ' '}}). \
                rename(columns={'Amount of Insurance D': 'Amount of Insurance', 'Amount of Insurance ND': 'Amount of Insurance'})

    # Builds one Mine Subsidence Coverage Rate table for the given county
    # (OH only — "ATHENS" -> MSB, "ERIE" -> MSC; both share the same shape).
    # Returns a dataframe
    def _buildMineSubsidenceCovRate(self, county):
        rate = self.buildDataFrame("BP7_Bldg_MineSubsidenceCov_Rate")
        return rate[rate['County'] == county].replace({'County': {county: 'Charge per property location:'}})

    # Builds one Stop Gap "Increased Limits" table (OH's SGLIL/SGEIL, WA's
    # SGLIL — differ only by which base-rate table they read from).
    # Returns a dataframe
    def _buildStopGapIncreasedLimits(self, tableCode):
        sg = self.buildDataFrame(tableCode)
        sg = sg[sg['Limit'].isin(['500000/500000/500000', '1000000/1000000/1000000'])]
        return sg.rename(columns={'Limit': 'Increased Limits', 'MinimumPremium': 'Minimum Premium'}).replace({'Increased Limits': self.STOP_GAP_LIMIT_LABELS})

    # Builds one Stop Gap "Premium Determination" table (OH's SGLPD/SGEPD,
    # WA's SGLPD — differ only by which base-rate table they read from).
    # Returns a dataframe
    def _buildStopGapPremiumDetermination(self, tableCode):
        sg = self.buildDataFrame(tableCode)
        sg = sg[sg['Limit'] == '100000/100000/500000']
        return sg.rename(columns={'Limit': 'Increased Limits', 'MinimumPremium': 'Minimum Premium'}).replace({'Increased Limits': self.STOP_GAP_LIMIT_LABELS}). \
                filter(items=['Rate', 'Minimum Premium']).rename(columns={'Rate': 'Rate (per $100 of payroll)'})

    # Builds the Underground Petroleum Storage Tank Deductible Coverage table (OH only).
    # Returns a dataframe
    def buildUPS(self):
        ups = self.buildDataFrame("BP7_Pol_UGPetroleumStorageTank_BaseRate")
        return ups.rename(columns={'NoOfTanks': 'Number of Tanks', 'LimitPerTank': 'Limit Per Tank', 'AggregateLimit': 'Aggregate Limit', 'UGPetroleumStorageTankBaseRate': 'Premium'})

    # ── Formatting ──────────────────────────────────────────────────────────

    # Converts the given pixels to inches (same 7px/char ratio ExcelSettingsBOP
    # uses internally for its own Table Layout column widths).
    def pixelsToInches(self, px):
        return px / 7.0

    # Applies currency formatting to every data row (row 4+) of every column
    # from min_col onward. Replaces 8 copy-pasted "for col: for row:" loops
    # in the source's individual format*() methods.
    def _applyCurrencyFormat(self, ws, min_col, start_row=4):
        for col in range(min_col, ws.max_column + 1):
            char = get_column_letter(col)
            for row in range(start_row, ws.max_row + 1):
                ws[f'{char}{row}'].number_format = self.currencyFormat

    def formatRateCappingFactors(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(80)
        ws.column_dimensions['B'].width = self.pixelsToInches(80)
        for cell in ws['4:4']:
            cell.number_format = self.percentageFormat

    def formatDistributionFactors(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(125)

    def formatWHExclusion(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(125)

    def formatLiabilityHazards(self, ws):
        ws.column_dimensions['A'].width = self.pixelsToInches(190)
        ws.column_dimensions['B'].width = self.pixelsToInches(150)

    # Applies the IBHS worksheet's group-title bands over its 3 stacked
    # blocks (Zones / Hurricane discounts / Wind & Hail discounts). Pokes
    # absolute row numbers that assume generateMultiTableWorksheet wrote the
    # 3 blocks contiguously with no gaps — see the module-level caveat above.
    def formatIBHSZones(self, ws, boldFont):
        ws.row_dimensions[4].height = self.pixelsToInches(450)
        ws.row_dimensions[5].height = self.pixelsToInches(250)
        for row_range in (ws['4:4'], ws['5:5'], ws['6:6']):
            for cell in row_range:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.merge_cells('B3:E3')
        ws.merge_cells('B4:E4')
        ws.merge_cells('B5:E5')
        ws.merge_cells('B6:E6')

        for col in range(2, ws.max_column + 1):
            char = get_column_letter(col)
            for row in range(4, ws.max_row + 1):
                ws[f'{char}{row}'].number_format = self.percentageFormat

        for row_range in (ws['8:8'], ws['13:13']):
            for cell in row_range:
                cell.font = boldFont
                cell.alignment = Alignment(horizontal='center', vertical='bottom', wrap_text=True)

        for row_range in (ws['7:7'], ws['12:12']):
            for cell in row_range:
                cell.border = None

        ws.insert_rows(7)
        ws['B8'] = 'Commercial Hurricane Premium Discounts*'
        ws.insert_rows(13)
        ws.insert_rows(14)
        ws.insert_rows(15)
        ws['A13'] = '* - Adjustments: Metal Roof > 10 years old or metal roof with no sub-decking, or both; all non-metal roofs > 5 years old:'
        ws['A14'] = '10 point reduction from above discounts all zones'
        ws['B16'] = 'Commercial Other Wind & Hail Premium Discounts**'
        ws['A21'] = '** - Adjustments: Metal Roof > 10 years old, All other roofs > 5 years old:'
        ws['A22'] = '10 point reduction from above discounts all zones'

        for row_num in (8, 16):
            for cell in ws[f'{row_num}:{row_num}']:
                cell.border = self._thinBorder
                cell.font = boldFont
                cell.alignment = Alignment(horizontal='center', vertical='bottom', wrap_text=True)
            ws.merge_cells(f'B{row_num}:E{row_num}')

        for row_num in (14, 22):
            for cell in ws[f'{row_num}:{row_num}']:
                cell.font = boldFont

    def formatMineSubsidence(self, ws):
        self._applyCurrencyFormat(ws, min_col=1)
        ws.column_dimensions['A'].width = self.pixelsToInches(175)
        ws.column_dimensions['C'].width = self.pixelsToInches(175)

    # Shared by MSB (OH "ATHENS") and MSC (OH "ERIE") — identical layout,
    # only the underlying county-filtered data differs.
    def _formatMineSubsidenceCovRate(self, ws):
        self._applyCurrencyFormat(ws, min_col=2)
        ws.column_dimensions['A'].width = self.pixelsToInches(175)
        ws.column_dimensions['C'].width = self.pixelsToInches(175)
        ws.delete_rows(3)
        for cell in ws['3:3']:
            cell.border = None

    # Shared by SGLIL/SGEIL (OH) and SGLIL (WA).
    def _formatStopGapIncreasedLimits(self, ws):
        self._applyCurrencyFormat(ws, min_col=3)
        ws.column_dimensions['A'].width = self.pixelsToInches(200)
        ws.column_dimensions['C'].width = self.pixelsToInches(150)

    # Shared by SGLPD/SGEPD (OH) and SGLPD (WA).
    def _formatStopGapPremiumDetermination(self, ws):
        self._applyCurrencyFormat(ws, min_col=2)
        ws.column_dimensions['A'].width = self.pixelsToInches(175)
        ws.column_dimensions['B'].width = self.pixelsToInches(150)

    def formatUPS(self, ws):
        self._applyCurrencyFormat(ws, min_col=2)
        ws.column_dimensions['A'].width = self.pixelsToInches(175)
        ws.column_dimensions['B'].width = self.pixelsToInches(150)
        ws.column_dimensions['C'].width = self.pixelsToInches(150)
        ws.column_dimensions['D'].width = self.pixelsToInches(150)
        ws['A10'] = "7 or more"
        ws['B10'] = "Ineligible for coverage"
        ws.merge_cells('B10:D10')
        for cell in ws['10:10']:
            cell.border = self._thinBorder
            cell.alignment = Alignment(horizontal='center', vertical='bottom', wrap_text=True)

    # ── Page assembly ───────────────────────────────────────────────────────

    # Which sheets get built is genuinely state-specific business logic
    # (each state files its own Additional Rules tables) — not duplication,
    # so it stays as one per-state sheetSpecs list.
    def _sheetSpecs(self):
        if self.state == "MD":
            return [
                ('LHL', 'AR Table 1.C.1-2. Liability for Hazards of Lead', self.buildLiabilityHazards, False, True, None, self.formatLiabilityHazards),
                ('WHE', 'AR Table A.3. Windstorm or Hail Exclusion (Coverages Not Included in By-peril Rating)', self.buildWHExclusion, False, True, None, self.formatWHExclusion),
                ('RC', 'AR-4 Table A.4. Transition Capping Program', self.buildRateCappingFactors, False, True, None, self.formatRateCappingFactors),
            ]
        if self.state == "OH":
            return [
                ('MSB', 'AR-1 Table B.2. Mine Subsidence', lambda: self._buildMineSubsidenceCovRate('ATHENS'), False, True, None, self._formatMineSubsidenceCovRate),
                ('MSC', 'AR-1 Table C.2. Mine Subsidence', lambda: self._buildMineSubsidenceCovRate('ERIE'), False, True, None, self._formatMineSubsidenceCovRate),
                ('SGLIL', 'AR-2 Table C.2 Stop Gap - Employers Liability Coverage Increased Limits', lambda: self._buildStopGapIncreasedLimits('BP7_Stop_Gap_Base_Rate'), False, True, None, self._formatStopGapIncreasedLimits),
                ('SGLPD', 'AR-2 Table D. Stop Gap - Employers Liability Coverage Premium Determination', lambda: self._buildStopGapPremiumDetermination('BP7_Stop_Gap_Base_Rate'), False, True, None, self._formatStopGapPremiumDetermination),
                ('SGEIL', 'AR-2 Table E.4.C Stop Gap - Extended Coverage Endorsement Increased Limits', lambda: self._buildStopGapIncreasedLimits('BP7_Extended_Stop_Gap_Base_Rate'), False, True, None, self._formatStopGapIncreasedLimits),
                ('SGEPD', 'AR-2 Table E.4 Stop Gap - Extended Coverage Endorsement Premium Determination', lambda: self._buildStopGapPremiumDetermination('BP7_Extended_Stop_Gap_Base_Rate'), False, True, None, self._formatStopGapPremiumDetermination),
                ('UPS', 'AR-3 Table F. Underground Petroleum Storage Tank Deductible Coverage', self.buildUPS, False, True, None, self.formatUPS),
                ('RC', 'AR-4 Table B. Transition Capping Program', self.buildRateCappingFactors, False, True, None, self.formatRateCappingFactors),
            ]
        if self.state == "AL":
            return [
                ('WHE', 'AR Table 1.3 Windstorm or Hail Exclusion (Coverages Not Included in By-peril Rating)', self.buildWHExclusion, False, True, None, self.formatWHExclusion),
            ]
        if self.state in ("IN", "KY"):
            return [
                ('MS', 'AR Table 1. Mine Subsidence Insurance', self.buildMineSubsidence, False, True, None, self.formatMineSubsidence),
            ]
        if self.state == "WA":
            return [
                ('SGLIL', 'AR-1 Table C.2 Stop Gap - Employers Liability Coverage Increased Limits', lambda: self._buildStopGapIncreasedLimits('BP7_Stop_Gap_Base_Rate'), False, True, None, self._formatStopGapIncreasedLimits),
                ('SGLPD', 'AR-1 Table D. Stop Gap - Employers Liability Coverage Premium Determination', lambda: self._buildStopGapPremiumDetermination('BP7_Stop_Gap_Base_Rate'), False, True, None, self._formatStopGapPremiumDetermination),
            ]
        return [
            ('RC', 'AR Table 98.C. Rate Capping', self.buildRateCappingFactors, False, True, None, self.formatRateCappingFactors),
            ('DF', 'AR - 2 Distribution Factor', self.buildDistributionFactors, False, True, None, self.formatDistributionFactors),
        ]

    # Sets up the Additional Rules Excel file and creates a worksheet for
    # each of this state's tables, then applies each table's postFormat.
    # progress_callback (optional) is called with a short message before
    # each sheet is built.
    # Returns the Excel workbook
    def buildAdditionalRulesPage(self, progress_callback=None):
        companies = [c for c in self.rateTables.keys() if c != 'CW']

        AdditionalRules = ExcelSettingsBOP.Excel(state=self.state, programName='Additional Rules', nEffective=self.nEffective, rEffective=self.rEffective, companyList=companies)

        sheetSpecs = self._sheetSpecs()
        total = len(sheetSpecs) + (1 if self.state == "AL" else 0)  # AL also gets the 3-block IBHS sheet

        built = 0
        for tableCode, title, build, useIndex, useHeader, layoutKey, postFormat in sheetSpecs:
            built += 1
            if progress_callback:
                progress_callback(f"Building sheet {built}/{total}: {tableCode}...")
            print(f"  [{built}/{total}] Building sheet: {tableCode}")
            ws = AdditionalRules.generateWorksheet(tableCode, title, build(), useIndex, useHeader, layout_key=layoutKey)
            if postFormat:
                postFormat(ws)

        if self.state == "AL":
            built += 1
            if progress_callback:
                progress_callback(f"Building sheet {built}/{total}: IBHS...")
            print(f"  [{built}/{total}] Building sheet: IBHS")
            _, ibhsWs = AdditionalRules.generateMultiTableWorksheet(
                'IBHS', 'AR Table 2. IBHS Certificate Discounts',
                [self.buildIBHSzones(), self._buildIBHSCertificate('IBHS Hurricane'), self._buildIBHSCertificate('IBHS High Wind and Hail')],
                False, True,
            )
            self.formatIBHSZones(ibhsWs, AdditionalRules.fontBold)

        if progress_callback:
            progress_callback("Building Index sheet...")
        print(f"  [{total}/{total}] Building sheet: Index")
        AdditionalRules.createIndex()
        return AdditionalRules.getWB()
