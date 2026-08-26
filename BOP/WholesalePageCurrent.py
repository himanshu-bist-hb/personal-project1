# This module builds and formats the Wholesale State Page workbook (pre-2.0).
#
# Same idea as WholesalePage.py (BP-2.0) — see that file's module docstring
# — but transcribed from the root-level WholesalePageCurrent.py, which
# predates the Territory Definitions workbook (State Territory Multiplier
# tables are built straight from the ratebook's own
# BP7_Peril_TerritorialFactor table, same TRBG/TRPP/TRLB codes + PROGRAM_TR
# layout key already used by Hab/Auto Service/Retail's pre-2.0 pages) and
# does not have the cat4/L-Products exclusions BP-2.0 added to
# buildBaseRates/buildPDDeductibleAmount.

import numpy as np
import pandas as pd

from . import ExcelSettingsBOP


class Wholesale:
    # Tables buildBaseRates() needs present for a given company before it can
    # be built without a KeyError.
    _BASE_RATE_TABLES = (
        'BP7_Peril_Building_Base_Rates',
        'BP7_Peril_BPP_Base_Rates',
        'BP7_Peril_Liability_Base_Rates',
    )

    def __init__(self, state, rateTables, perils, perilsConversions, nEffective, rEffective) -> None:
        self.state = state
        self.rateTables = rateTables
        # Individual program pages (unlike All Programs) also show the
        # "AllPeril" row/column — see [[bop_auto_service_allperil_row_fix]].
        self.perils = list(perils) + ['allperil']
        self.perilsConversions = perilsConversions
        self.nEffective = nEffective
        self.rEffective = rEffective

        self.wholesaleProgramCode = 80000

    # Builds a dataframe for the given table code
    # The hierarchy matches Business Auto: lower-level company (NACO/NAFF/NICOF)
    # first, then NGIC (state-level default), then CW as the country-wide
    # fallback. See [[bop-nesting-order]] — the root-level
    # WholesalePageCurrent.py checked NGIC first, which is backwards.
    # Returns the dataframe that was built
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

    # Builds the base rates table for the given company
    # Returns a dataframe
    def buildBaseRates(self, company):
        buildingBaseRates = pd.DataFrame(data=self.rateTables[company]['BP7_Peril_Building_Base_Rates'][1:], index=None, columns=self.rateTables[company]['BP7_Peril_Building_Base_Rates'][0])
        bppBaseRates = pd.DataFrame(data=self.rateTables[company]['BP7_Peril_BPP_Base_Rates'][1:], index=None, columns=self.rateTables[company]['BP7_Peril_BPP_Base_Rates'][0])
        liabilityBaseRates = pd.DataFrame(data=self.rateTables[company]['BP7_Peril_Liability_Base_Rates'][1:], index=None, columns=self.rateTables[company]['BP7_Peril_Liability_Base_Rates'][0])
        filteredBuilingBaseRates = buildingBaseRates.query(f'Class_Code_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'BuildingBaseRate'])
        filteredBPPBaseRates = bppBaseRates.query(f'Class_Code_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'BPPBaseRate'])
        filteredLiabilityBaseRates = liabilityBaseRates.query(f'ClassCode_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils} & OccupanyType != "tenant"'). \
                pivot(index='Peril TypeCode', columns='OccupanyType', values='LiabilityFactor').reset_index().rename_axis(None, axis=1)
        baseRates = pd.merge(filteredBuilingBaseRates, filteredBPPBaseRates, how='inner', on='Peril TypeCode')
        finalBaseRates = pd.merge(baseRates, filteredLiabilityBaseRates, how='outer', on='Peril TypeCode')
        return finalBaseRates.replace({'Peril TypeCode': self.perilsConversions}).rename(columns={"Peril TypeCode": "Peril", "BuildingBaseRate": "Building",
                "BPPBaseRate": "BPP", "buildingOwnerLessorsrisk": "Liability Lessor's Risk", "buildingOwnerOccupant": "Liability Occupant"}).sort_values(by='Peril')

    # Builds the territory multiplier table for the given coverage (either building, bpp, or liability)
    # Returns a dataframe
    def buildTerritoryMultiplier(self, coverage):
        territorialFactor = self.buildDataFrame("BP7_Peril_TerritorialFactor")
        filteredTerritorialFactor = territorialFactor.query(f'Class_Code_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils}').replace({'Peril TypeCode': self.perilsConversions}).rename(columns={'TerritoryCode': 'Territory'})
        if coverage.casefold() == 'building':
            return filteredTerritorialFactor.pivot(index='Territory', columns='Peril TypeCode', values='BldgTerritoryFactor').reset_index('Territory'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})
        elif coverage.casefold() == 'bpp':
            return filteredTerritorialFactor.pivot(index='Territory', columns='Peril TypeCode', values='BPPTerritoryFactor').reset_index('Territory'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})
        elif coverage.casefold() == 'liability':
            return filteredTerritorialFactor.pivot(index='Territory', columns='Peril TypeCode', values='LiabilityTerritoryFactor').reset_index('Territory'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})

    # Builds the construction type table for the given coverage (either building or bpp)
    # Returns a dataframe
    def buildConstructionType(self, coverage):
        constructionType = self.buildDataFrame("BP7 Peril Construction_Type")
        filteredConstructionType = constructionType.query(f'Class_Code_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils}').replace({'Peril TypeCode': self.perilsConversions}). \
                rename(columns={'ConstructionClassDisplay Name': 'Construction'})
        if coverage.casefold() == 'building':
            return filteredConstructionType.pivot(index='Construction', columns='Peril TypeCode', values='BldgConstructionClassFactor').reset_index('Construction'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})
        elif coverage.casefold() == 'bpp':
            return filteredConstructionType.pivot(index='Construction', columns='Peril TypeCode', values='BPPConstructionClassFactor').reset_index('Construction'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})

    # Builds the theft options table
    # Returns a dataframe
    def buildTheftOptions(self):
        theftOptions = self.buildDataFrame("BP7_Peril_BPP_Theft_Options_Factor")
        filteredTheftOptions = theftOptions.query(f'Class_Code_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils} & `Theft Option` != "Full Theft"'). \
                replace({'Peril TypeCode': self.perilsConversions})
        filteredTheftOptions = filteredTheftOptions.drop(filteredTheftOptions[filteredTheftOptions['Peril TypeCode'].isin(['L-Products', 'NC-BINC', 'WF'])].index)
        return filteredTheftOptions.pivot(index='Peril TypeCode', columns='Theft Option', values='BPP Theft Options Factor').reset_index('Peril TypeCode'). \
                rename(columns={'Peril TypeCode': 'Peril', 'Excluded Theft': 'Excluded', 'Limited Theft': 'Limited'})

    # Builds the year built modifier table for the given coverage (either building or bpp)
    # Returns a dataframe
    def buildYearBuiltModifier(self, coverage):
        yearBuiltModifier = pd.DataFrame()
        if coverage.casefold() == 'building':
            yearBuiltModifier = self.buildDataFrame("BP7 Peril_Building_Year_Built_Modifier")
        elif coverage.casefold() == 'bpp':
            yearBuiltModifier = self.buildDataFrame("BP7 Peril_BPP_Year_Built_Modifier")
        filteredYearBuiltModifier = yearBuiltModifier.query(f'Class_Code_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils}').replace({'Peril TypeCode': self.perilsConversions}). \
                fillna({'Year_Built_Max': 0}).astype({'Year_Built_Min': 'int64', 'Year_Built_Max': 'int64'}).astype({'Year_Built_Min': 'string', 'Year_Built_Max': 'string'}) # Converting to int first to get rid of decimal places
        filteredYearBuiltModifier['Year Built Range'] = np.where(filteredYearBuiltModifier['Year_Built_Max'] == '0',
                                                                 filteredYearBuiltModifier['Year_Built_Min'] + '+',
                                                                 filteredYearBuiltModifier['Year_Built_Min'] + ' - ' + filteredYearBuiltModifier['Year_Built_Max'])
        if coverage.casefold() == 'building':
            return filteredYearBuiltModifier.pivot(index='Year Built Range', columns='Peril TypeCode', values='Bldg_Year_Built_Factor').reset_index('Year Built Range'). \
                    drop(columns=['L-Products', 'WF', 'NC-BINC'], errors='ignore')
        elif coverage.casefold() == 'bpp':
            return filteredYearBuiltModifier.pivot(index='Year Built Range', columns='Peril TypeCode', values='BPP_Year_Built_Factor').reset_index('Year Built Range'). \
                    drop(columns=['L-Products', 'WF', 'NC-BINC'], errors='ignore')

    # Builds the equipment breakdown base rate table
    # Returns a dataframe
    def buildEBBaseRate(self):
        ebBaseRate = self.buildDataFrame("BP7_EBBaseRate")
        return ebBaseRate.query(f'Class_Code_Min == {self.wholesaleProgramCode}').rename(columns={'BaseRate': 'Rate'}).filter(items=['Rate'])

    # Builds the property damage liability deductible factor table
    # Returns a dataframe
    def buildPDDeductibleAmount(self):
        pdDeductibleAmount = self.buildDataFrame("BP7_Peril_Property_Damage_Liability_Factor")
        return pdDeductibleAmount.query(f'ClassCode_Min == {self.wholesaleProgramCode} & `Peril TypeCode` == "liability1"').rename(columns={'PDDeductibleAmount': 'P.D. Deductible Amount', 'PDDeductibleFactor': 'Factor'}). \
                replace({'P.D. Deductible Amount': {'NoDeductible': '0'}}).astype({'P.D. Deductible Amount': 'int64'}).sort_values(by=['P.D. Deductible Amount']). \
                replace({'P.D. Deductible Amount': {0: 'No Deductible'}}).filter(items=['P.D. Deductible Amount', 'Factor'])

    # Builds the liability limit factor table
    # Returns a dataframe
    def buildLiabilityLimitFactor(self):
        liabilityLimitFactor = self.buildDataFrame("BP7_Peril_ILF_Factor")
        return liabilityLimitFactor.query(f'ClassCode_Min == {self.wholesaleProgramCode} & `Peril TypeCode` == "liability1"').filter(items=['LiabilityLimit', 'LiabilityFactor']). \
                rename(columns={'LiabilityLimit': 'Liability Limit of Insurance', 'LiabilityFactor': 'Factor'}).astype({'Liability Limit of Insurance': 'int32'})

    # Builds the liquified petroleum gas (LPG) exposures table
    # Returns a dataframe
    def buildLPGExposure(self):
        lpgExposure = self.buildDataFrame("BP7_LPG_Premium")
        return lpgExposure.rename(columns={'LiabilityLimitOfInsurance': 'Liability Limit of Insurance', 'LPGPremium': 'Premium (each premises)'})

    # Builds the liability size of risk modifier table
    # Returns a dataframe
    def buildLiabilitySizeRisk(self):
        liabilitySizeRisk = self.buildDataFrame("BP7_Peril_Liability_Factor_Receipts_Limit")
        filteredLiabilitySizeRisk = liabilitySizeRisk.query(f'ClassCode_Min == {self.wholesaleProgramCode} & `Peril TypeCode` in {self.perils} & AutoServType == "N/A" & FoodServType == "N/A"').rename(columns={'ReceiptMin': 'Min', 'ReceiptMax': 'Max', 'LiabilityFactor': 'Factor'}).replace({'Peril TypeCode': self.perilsConversions})
        return filteredLiabilitySizeRisk.pivot(index=['Min', 'Max'], columns='Peril TypeCode', values='Factor').reset_index(['Min', 'Max']).fillna({'Max': 'and over'})

    # Builds the endorsement charge table
    # Returns a dataframe
    def buildEndorsementCharge(self):
        endorsementCharge = self.buildDataFrame("BP7_PlusEndorsementCharge")
        return endorsementCharge.query(f'ClassCodeMIn == {self.wholesaleProgramCode}').filter(items=['PlusEndorsementCharge']).rename(columns={'PlusEndorsementCharge': 'Base premium for each Wholesale premises'})

    # Sets up the Wholesale Excel file and creates a separate worksheet for
    # each of the given dataframes. progress_callback (optional) is called
    # with a short message before each sheet is built.
    # Returns the Excel workbook
    def buildWholesalePage(self, progress_callback=None):
        companies = [c for c in self.rateTables.keys() if c != 'CW']

        Wholesale = ExcelSettingsBOP.Excel(state=self.state, programName='Wholesale', nEffective=self.nEffective, rEffective=self.rEffective, companyList=companies)

        sheetSpecs = []
        # A company can be present in rateTables (its ratebook was uploaded)
        # without having filed its own base-rate tables — a deviation
        # ratebook may only override a handful of tables. Check for the
        # specific tables buildBaseRates() needs, not just company
        # membership, or it KeyErrors on that company's missing table.
        for company, tab, label in (('NACO', 'BRNACO', 'NW Assurance'), ('NAFF', 'BRNAFF', 'NW Affinity'),
                                     ('NGIC', 'BRNGIC', 'NW General Insurance Company'), ('NICOF', 'BRNICOF', 'NICOF')):
            if company in self.rateTables and all(t in self.rateTables[company] for t in self._BASE_RATE_TABLES):
                sheetSpecs.append((tab, f'W Table 3.B.1. {label} State Base Rates', lambda c=company: self.buildBaseRates(c), False, True, 'AS_BR', None))

        sheetSpecs += [
            ('TRBG', 'W Table 3.C.1.a. State Territory Multiplier - Building', lambda: self.buildTerritoryMultiplier('Building'), False, True, 'PROGRAM_TR', None),
            ('TRPP', 'W Table 3.C.1.a. State Territory Multiplier - BPP', lambda: self.buildTerritoryMultiplier('BPP'), False, True, 'PROGRAM_TR', None),
            ('TRLB', 'W Table 3.C.1.a. State Territory Multiplier - Liability', lambda: self.buildTerritoryMultiplier('Liability'), False, True, 'PROGRAM_TR', None),
            ('CBG', 'W Table 3.C.2.c. Construction Factor - Building', lambda: self.buildConstructionType('Building'), False, True, None, None),
            ('CPP', 'W Table 3.C.2.c. Construction Factor - BPP', lambda: self.buildConstructionType('BPP'), False, True, None, None),
            ('ET', 'W Table 3.C.2.m. Exclude Theft Factor', self.buildTheftOptions, False, True, None, None),
            ('YBBG', 'W Table 3.C.2.p. Year Built Modifier - Building', lambda: self.buildYearBuiltModifier('Building'), False, True, None, None),
            ('YBPP', 'W Table 3.C.2.p. Year Built Modifier - BPP', lambda: self.buildYearBuiltModifier('BPP'), False, True, None, None),
            ('EBB', 'W Table 3.C.3.a. EB Base Rate', self.buildEBBaseRate, False, True, None, None),
            ('PDLD', 'W Table 3.C.4.b. Property Damage Liability Deductible Factor', self.buildPDDeductibleAmount, False, True, None, None),
            ('LL', 'W Table 3.C.4.d. Liability Limit Factor', self.buildLiabilityLimitFactor, False, True, None, None),
            ('LPGE', 'W Table 3.C.4.e. Liquefied Petroleum Gas (LPG) Exposures', self.buildLPGExposure, False, True, None, None),
            ('LS', 'W Table 3.C.4.f. Liability Size of Risk Modifier', self.buildLiabilitySizeRisk, False, True, 'LS_CURRENT', None),
            ('PLUS', 'W Table 4.A. Wholesale PLUS Endorsement', self.buildEndorsementCharge, False, True, None, None),
        ]

        total = len(sheetSpecs)
        for i, (tableCode, title, build, useIndex, useHeader, layoutKey, postFormat) in enumerate(sheetSpecs, start=1):
            if progress_callback:
                progress_callback(f"Building sheet {i}/{total}: {tableCode}...")
            print(f"  [{i}/{total}] Building sheet: {tableCode}")
            ws = Wholesale.generateWorksheet(tableCode, title, build(), useIndex, useHeader, layout_key=layoutKey)
            if postFormat:
                postFormat(ws)

        if progress_callback:
            progress_callback("Building Index sheet...")
        print(f"  [{total}/{total}] Building sheet: Index")
        Wholesale.createIndex()
        return Wholesale.getWB()
