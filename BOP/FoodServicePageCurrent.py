# This module builds and formats the Food Service State Page workbook (pre-2.0).
#
# Same idea as FoodServicePage.py (BP-2.0) — see that file's module docstring
# — but transcribed from the root-level FoodServicePageCurrent.py, which
# predates the Territory Definitions workbook (State Territory Multiplier
# tables are built straight from the ratebook's own
# BP7_Peril_TerritorialFactor table, same TRBG/TRPP/TRLB codes + PROGRAM_TR
# layout key already used by Hab/Auto Service/Retail/Service/Office/
# Wholesale's pre-2.0 pages), does not have the cat4/fire4 exclusions BP-2.0
# added to buildBaseRates, and already calls buildBaseRates(company)
# per-company correctly (unlike the 2.0 root, which hardcoded 'NGIC' for
# every company's tab — see [[bop_wholesale_port]]).

import numpy as np
import pandas as pd

from . import ExcelSettingsBOP


class Food:
    # Tables buildBaseRates() needs present for a given company before it can
    # be built without a KeyError.
    _BASE_RATE_TABLES = (
        'BP7_Peril_Building_Base_Rates',
        'BP7_Peril_BPP_Base_Rates',
        'BP7_Peril_Business_Income_Base_Rate',
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

        self.foodProgramCode = 40000

    # Builds a dataframe for the given table code
    # The hierarchy matches Business Auto: lower-level company (NACO/NAFF/NICOF)
    # first, then NGIC (state-level default), then CW as the country-wide
    # fallback. See [[bop-nesting-order]] — the root-level
    # FoodServicePageCurrent.py checked NGIC first, which is backwards.
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
        businessIncomeBaseRates = pd.DataFrame(data=self.rateTables[company]['BP7_Peril_Business_Income_Base_Rate'][1:], index=None, columns=self.rateTables[company]['BP7_Peril_Business_Income_Base_Rate'][0])
        filteredBuilingBaseRates = buildingBaseRates.query(f'Class_Code_Min == {self.foodProgramCode} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'BuildingBaseRate'])
        filteredBPPBaseRates = bppBaseRates.query(f'Class_Code_Min == {self.foodProgramCode} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'BPPBaseRate'])
        filteredLiabilityBaseRates = liabilityBaseRates.query(f'ClassCode_Min == {self.foodProgramCode} & `Peril TypeCode` in {self.perils} & OccupanyType != "tenant"'). \
                pivot(index='Peril TypeCode', columns='OccupanyType', values='LiabilityFactor').reset_index().rename_axis(None, axis=1)
        filteredBusinessIncomeBaseRates = businessIncomeBaseRates.query(f'`Peril TypeCode` in {self.perils}  & OccupanyType != "tenant"'). \
                    pivot(index='Peril TypeCode', columns='OccupanyType', values='BusinessIncomeBaseRate').reset_index().rename_axis(None, axis=1). \
                    rename(columns={"buildingOwnerLessorsrisk": "Bus Inc Lessor's Risk", "buildingOwnerOccupant": "Bus Inc Occupant"})
        baseRates = pd.merge(filteredBuilingBaseRates, filteredBPPBaseRates, how='inner', on='Peril TypeCode')
        mergedBaseRates = pd.merge(baseRates, filteredBusinessIncomeBaseRates, how='inner', on='Peril TypeCode')
        finalBaseRates = pd.merge(mergedBaseRates, filteredLiabilityBaseRates, how='outer', on='Peril TypeCode')
        return finalBaseRates.replace({'Peril TypeCode': self.perilsConversions}).rename(columns={"Peril TypeCode": "Peril", "BuildingBaseRate": "Building",
                "BPPBaseRate": "BPP", "buildingOwnerLessorsrisk": "Liability Lessor's Risk", "buildingOwnerOccupant": "Liability Occupant"}).sort_values(by='Peril')

    # Builds the territory multiplier table for the given coverage (either building, bpp, liability or business income)
    # Returns a dataframe
    def buildTerritoryMultiplier(self, coverage):
        territorialFactor = self.buildDataFrame("BP7_Peril_TerritorialFactor")
        filteredTerritorialFactor = territorialFactor.query(f'Class_Code_Min == {self.foodProgramCode} & `Peril TypeCode` in {self.perils}').replace({'Peril TypeCode': self.perilsConversions}).rename(columns={'TerritoryCode': 'Territory'})
        if coverage.casefold() == 'building':
            return filteredTerritorialFactor.pivot(index='Territory', columns='Peril TypeCode', values='BldgTerritoryFactor').reset_index('Territory'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})
        elif coverage.casefold() == 'bpp':
            return filteredTerritorialFactor.pivot(index='Territory', columns='Peril TypeCode', values='BPPTerritoryFactor').reset_index('Territory'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})
        elif coverage.casefold() == 'business income':
            return filteredTerritorialFactor.pivot(index='Territory', columns='Peril TypeCode', values='BITerritoryFactor').reset_index('Territory'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})
        elif coverage.casefold() == 'liability':
            return filteredTerritorialFactor.pivot(index='Territory', columns='Peril TypeCode', values='LiabilityTerritoryFactor').reset_index('Territory'). \
                    drop(columns=['L-Products', 'L-Violence', 'L-OtherMed', 'L-OtherPrem', 'WF', 'NC-BINC'], errors='ignore').rename(columns={'L-SlipFall': 'LIAB-Other'})

    # Builds the construction type table for the given coverage (either building or bpp)
    # Returns a dataframe
    def buildConstructionType(self, coverage):
        constructionType = self.buildDataFrame("BP7 Peril Construction_Type")
        filteredConstructionType = constructionType.query(f'Class_Code_Min == {self.foodProgramCode} & `Peril TypeCode` in {self.perils}').replace({'Peril TypeCode': self.perilsConversions}). \
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
        filteredTheftOptions = theftOptions.query(f'Class_Code_Min == {self.foodProgramCode} & `Peril TypeCode` in {self.perils} & `Theft Option` != "Full Theft"'). \
                replace({'Peril TypeCode': self.perilsConversions})
        filteredTheftOptions = filteredTheftOptions.drop(filteredTheftOptions[filteredTheftOptions['Peril TypeCode'].isin(['L-Products', 'NC-BINC', 'WF'])].index)
        return filteredTheftOptions.pivot(index='Peril TypeCode', columns='Theft Option', values='BPP Theft Options Factor').reset_index('Peril TypeCode'). \
                rename(columns={'Peril TypeCode': 'Peril', 'Excluded Theft': 'Excluded', 'Limited Theft': 'Limited'})

    # Builds the year built modifier table for the given coverage (either building, bpp or business income)
    # Returns a dataframe
    def buildYearBuiltModifier(self, coverage):
        yearBuiltModifier = pd.DataFrame()
        if coverage.casefold() == 'building':
            yearBuiltModifier = self.buildDataFrame("BP7 Peril_Building_Year_Built_Modifier")
        elif coverage.casefold() == 'bpp':
            yearBuiltModifier = self.buildDataFrame("BP7 Peril_BPP_Year_Built_Modifier")
        elif coverage.casefold() == 'business income':
            yearBuiltModifier = self.buildDataFrame("BP7 Peril_BI_Year_Built_Modifier")
        filteredYearBuiltModifier = yearBuiltModifier.query(f'Class_Code_Min == {self.foodProgramCode} & `Peril TypeCode` in {self.perils}').replace({'Peril TypeCode': self.perilsConversions}). \
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
        elif coverage.casefold() == 'business income':
            return filteredYearBuiltModifier.pivot(index='Year Built Range', columns='Peril TypeCode', values='BI_Year_Built_Factor').reset_index('Year Built Range'). \
                    drop(columns=['L-Products', 'WF', 'NC-BINC'], errors='ignore')

    # Builds the equipment breakdown base rate table
    # Returns a dataframe
    def buildEBBaseRate(self):
        ebBaseRate = self.buildDataFrame("BP7_EBBaseRate")
        return ebBaseRate.query(f'Class_Code_Min == {self.foodProgramCode}').rename(columns={'BaseRate': 'Rate'}).filter(items=['Rate'])

    # Builds the property damage liability deductible factor table
    # Returns a dataframe
    def buildPDDeductibleAmount(self):
        pdDeductibleAmount = self.buildDataFrame("BP7_Peril_Property_Damage_Liability_Factor")
        return pdDeductibleAmount.query(f'ClassCode_Min == {self.foodProgramCode} & `Peril TypeCode` == "liability1"').rename(columns={'PDDeductibleAmount': 'P.D. Deductible Amount', 'PDDeductibleFactor': 'Factor'}). \
                replace({'P.D. Deductible Amount': {'NoDeductible': '0'}}).astype({'P.D. Deductible Amount': 'int64'}).sort_values(by=['P.D. Deductible Amount']). \
                replace({'P.D. Deductible Amount': {0: 'No Deductible'}}).filter(items=['P.D. Deductible Amount', 'Factor'])

    # Builds the liability size of risk modifier table
    # Returns a dataframe
    def buildLiabilitySizeRisk(self):
        liabilitySizeRisk = self.buildDataFrame("BP7_Peril_Liability_Factor_Receipts_Limit")
        filteredLiabilitySizeRisk = liabilitySizeRisk.query(f'ClassCode_Min == {self.foodProgramCode} & `Peril TypeCode` == "liability1" & FoodServType != "N/A" & FoodServType != "EXQ"'). \
                rename(columns={'ReceiptMin': 'Min', 'ReceiptMax': 'Max'}).replace({'FoodServType': {'CAS': 'Casual', 'FAM': 'Family', 'FAS': 'Fast Food', 'FIN': 'Fine Dining', 'LTD': 'Limited Cooking'}})
        return filteredLiabilitySizeRisk.pivot(index=['Min', 'Max'], columns='FoodServType', values='LiabilityFactor').reset_index(['Min', 'Max']).fillna({'Max': 'and over'})

    # Builds the liability limit factor table
    # Returns a dataframe
    def buildLiabilityLimitFactor(self):
        liabilityLimitFactor = self.buildDataFrame("BP7_Peril_ILF_Factor")
        return liabilityLimitFactor.query(f'ClassCode_Min == {self.foodProgramCode} & `Peril TypeCode` == "liability1"').filter(items=['LiabilityLimit', 'LiabilityFactor']). \
                rename(columns={'LiabilityLimit': 'Liability Limit of Insurance', 'LiabilityFactor': 'Factor'}).astype({'Liability Limit of Insurance': 'int32'})

    # Builds the endorsement charge table
    # Returns a dataframe
    def buildEndorsementCharge(self):
        endorsementCharge = self.buildDataFrame("BP7_PlusEndorsementCharge")
        return endorsementCharge.query(f'ClassCodeMIn == {self.foodProgramCode}').filter(items=['PlusEndorsementCharge']).rename(columns={'PlusEndorsementCharge': 'Base premium for each Food Service premises'})

    # Builds the optional increased limits – spoilage from power outage table
    # Returns a dataframe
    def buildSpoilagePowerOutage(self):
        optionalCoverageRates = self.buildDataFrame("BP7_Optional_Coverage_Base_Rates")
        spoilagePowerOutage = optionalCoverageRates.query(f'CoverageName == "SpoilagePowerOutageIncreasedLimits"').rename(columns={'CoverageName': 'Coverage', 'BaseRate': 'Optional Increased Limit'})
        return spoilagePowerOutage.replace({'Coverage': {'SpoilagePowerOutageIncreasedLimits': 'Spoilage Power Outage'}})

    # Builds the optional increased limits - food contamination - loss of income table
    # Returns a dataframe
    def buildFoodContamination(self):
        miscBaseRates = self.buildDataFrame("BP7_Miscellaneous_Base_Rates")
        miscFactors = self.buildDataFrame("BP7_Miscellaneous_Factors_Table")
        foodContamEstablishment = self.buildDataFrame("BP7_Food_Contamination_Establishment")
        foodContamFranchise = self.buildDataFrame("BP7_Food_Contamination_Franchise")
        filteredMiscBaseRates = miscBaseRates.query(f'BaseRateName == "FoodContaminationLossOfIncome"')
        filteredMiscFactors = miscFactors.query(f'FactorName == "Receipts"')
        filteredFoodContamEst = foodContamEstablishment.query(f'FoodServiceType == "FIN"')
        filteredFoodContamFranch = foodContamFranchise.query(f'FranchiseInd == "Yes"')
        foodContamination = pd.concat([filteredMiscBaseRates, filteredMiscFactors, filteredFoodContamEst, filteredFoodContamFranch], ignore_index=True)
        foodContamination['Service'] = np.where(foodContamination['BaseRateName'].notnull(),
                                                'Base',
                                                np.where(foodContamination['FoodServiceType'].notnull(),
                                                         'Fine Dining',
                                                         np.where(foodContamination['FranchiseInd'].notnull(),
                                                                  'Franchise',
                                                                  'Receipts')))
        foodContamination['Rate or Factor'] = np.where(foodContamination['BaseRateName'].notnull(),
                                                       foodContamination['BaseRate'],
                                                       np.where(foodContamination['FoodServiceType'].notnull(),
                                                                foodContamination['FoodContaminationEstablishmentFactor'],
                                                                np.where(foodContamination['FranchiseInd'].notnull(),
                                                                         foodContamination['FoodContainmentFranchiseFactor'],
                                                                         foodContamination['Factor'])))
        return foodContamination.filter(items=['Service', 'Rate or Factor'])

    # Builds the off premises valet parking table
    # Returns a dataframe
    def buildValetParking(self):
        optionalCoverageRates = self.buildDataFrame("BP7_Optional_Coverage_Base_Rates")
        offPremValetPark = optionalCoverageRates.query(f'CoverageName == "OffPremisesValetParking"').rename(columns={'CoverageName': 'Coverage', 'BaseRate': 'Rate'})
        return offPremValetPark.replace({'Coverage': {'OffPremisesValetParking': 'Off Premises Valet Parking'}})

    # Builds the franchise upgrade endorsement table for the given program
    # Returns a dataframe
    def buildFranchiseUpgradeEndorsement(self):
        franchiseUpgradeBase = self.buildDataFrame("BP7_Franchise_Upgrade_Base")
        miscMinMaxPrem = self.buildDataFrame("BP7_Miscellaneous_Minimum/Maximum_Premium")
        filteredFranchiseUpgrade = franchiseUpgradeBase.query(f'MinClassCode == {self.foodProgramCode}')
        filteredMiscMinMaxPrem = miscMinMaxPrem.query(f'CoverageType == "BP7Pol_FranchiseUpgradeEndorsementCov_Ext"')
        franchiseUpgradeEndorsement = pd.concat([filteredFranchiseUpgrade, filteredMiscMinMaxPrem], ignore_index=True)
        franchiseUpgradeEndorsement['Rate or Premium'] = np.where(franchiseUpgradeEndorsement['RateType'].isnull(),
                                                                  'Minimum Premium',
                                                                  'Base Rate')
        franchiseUpgradeEndorsement['Per Building'] = np.where(franchiseUpgradeEndorsement['FranchiseUpgradeBase'].isnull(),
                                                               franchiseUpgradeEndorsement['Premium'],
                                                               franchiseUpgradeEndorsement['FranchiseUpgradeBase'])
        return franchiseUpgradeEndorsement.filter(items=['Rate or Premium', 'Per Building'])

    # Sets up the Food Service Excel file and creates a separate worksheet for
    # each of the given dataframes. progress_callback (optional) is called
    # with a short message before each sheet is built.
    # Returns the Excel workbook
    def buildFoodPage(self, progress_callback=None):
        companies = [c for c in self.rateTables.keys() if c != 'CW']

        FoodService = ExcelSettingsBOP.Excel(state=self.state, programName='Food Service', nEffective=self.nEffective, rEffective=self.rEffective, companyList=companies)

        sheetSpecs = []
        # A company can be present in rateTables (its ratebook was uploaded)
        # without having filed its own base-rate tables — a deviation
        # ratebook may only override a handful of tables. Check for the
        # specific tables buildBaseRates() needs, not just company
        # membership, or it KeyErrors on that company's missing table.
        for company, tab, label in (('NACO', 'BRNACO', 'NW Assurance'), ('NAFF', 'BRNAFF', 'NW Affinity'),
                                     ('NGIC', 'BRNGIC', 'NW General Insurance Company'), ('NICOF', 'BRNICOF', 'NICOF')):
            if company in self.rateTables and all(t in self.rateTables[company] for t in self._BASE_RATE_TABLES):
                sheetSpecs.append((tab, f'FS Table 3.B.1. {label} State Base Rates', lambda c=company: self.buildBaseRates(c), False, True, 'AS_BR', None))

        sheetSpecs += [
            ('TRBG', 'FS Table 3.C.1.a. State Territory Multiplier - Building', lambda: self.buildTerritoryMultiplier('Building'), False, True, 'PROGRAM_TR', None),
            ('TRPP', 'FS Table 3.C.1.a. State Territory Multiplier - BPP', lambda: self.buildTerritoryMultiplier('BPP'), False, True, 'PROGRAM_TR', None),
            ('TRLB', 'FS Table 3.C.1.a. State Territory Multiplier - Liability', lambda: self.buildTerritoryMultiplier('Liability'), False, True, 'PROGRAM_TR', None),
            ('CBG', 'FS Table 3.C.2.c. Construction Factor - Building', lambda: self.buildConstructionType('Building'), False, True, None, None),
            ('CPP', 'FS Table 3.C.2.c. Construction Factor - BPP and Bus Inc', lambda: self.buildConstructionType('BPP'), False, True, None, None),
            ('ET', 'FS Table 3.C.2.m. Exclude Theft Factor', self.buildTheftOptions, False, True, None, None),
            ('YBBG', 'FS Table 3.C.2.p. Year Built Modifier - Building', lambda: self.buildYearBuiltModifier('Building'), False, True, None, None),
            ('YBPP', 'FS Table 3.C.2.p. Year Built Modifier - BPP', lambda: self.buildYearBuiltModifier('BPP'), False, True, None, None),
            ('YBBI', 'FS Table 3.C.2.p. Year Built Modifier - Bus Inc', lambda: self.buildYearBuiltModifier('Business Income'), False, True, None, None),
            ('EBB', 'FS Table 3.C.3.a. EB Base Rate', self.buildEBBaseRate, False, True, None, None),
            ('PDLD', 'FS Table 3.C.4.b. Property Damage Liability Deductible Factor', self.buildPDDeductibleAmount, False, True, None, None),
            ('LS', 'FS Table 3.C.4.d. Liability Size of Risk Modifier', self.buildLiabilitySizeRisk, False, True, 'LS_CURRENT', None),
            ('LL', 'FS Table 3.C.4.e. Liability Limit Factor', self.buildLiabilityLimitFactor, False, True, None, None),
            ('PLUS', 'FS Table 4.A.1. Food Service PLUS Endorsement', self.buildEndorsementCharge, False, True, None, None),
            ('SPO', 'FS Table 4.A.2. Optional Increased Limits – Spoilage From Power Outage', self.buildSpoilagePowerOutage, False, True, None, None),
            ('C', 'FS Table 4.A.3. Optional Increased Limits – Food Contamination – Loss of Income', self.buildFoodContamination, False, True, None, None),
            ('VAL', 'FS Table 4.B. Off Premises Valet Parking', self.buildValetParking, False, True, None, None),
            ('FR', 'FS Table 4.C. Franchise Upgrade Endorsement', self.buildFranchiseUpgradeEndorsement, False, True, None, None),
        ]

        total = len(sheetSpecs)
        for i, (tableCode, title, build, useIndex, useHeader, layoutKey, postFormat) in enumerate(sheetSpecs, start=1):
            if progress_callback:
                progress_callback(f"Building sheet {i}/{total}: {tableCode}...")
            print(f"  [{i}/{total}] Building sheet: {tableCode}")
            ws = FoodService.generateWorksheet(tableCode, title, build(), useIndex, useHeader, layout_key=layoutKey)
            if postFormat:
                postFormat(ws)

        if progress_callback:
            progress_callback("Building Index sheet...")
        print(f"  [{total}/{total}] Building sheet: Index")
        FoodService.createIndex()
        return FoodService.getWB()
