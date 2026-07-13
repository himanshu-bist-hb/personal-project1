# This module builds and formats the All Peril State Page workbook (BP-2.0).
#
# The build*() methods below are unchanged business logic, transcribed from
# the root-level AllPerilPage.py: each pulls a table out of the ratebook data
# (via buildDataFrame's nesting waterfall) and shapes it into the DataFrame
# the rate page needs. Unlike All Programs, every table is filtered to the
# single "allperil" peril and broken out by PROGRAM (Class_Code_Min).
#
# All Excel formatting (fonts, column widths, sub-header labels, page setup)
# lives in ExcelSettingsBOP.py, driven by "BOP/BOP Input File.xlsx" — the
# All Peril-specific column profiles use the *_AP layout keys there.

import math

import pandas as pd

from . import ExcelSettingsBOP


class AllPeril:
    def __init__(self, state, rateTables, classCodes, protectionClassConversions, buildingCodes, nEffective, rEffective) -> None:
        self.state = state
        self.rateTables = rateTables
        self.classCodes = classCodes
        self.protectionClassConversions = protectionClassConversions
        self.buildingCodes = buildingCodes
        self.nEffective = nEffective # New business effective date
        self.rEffective = rEffective # Renewal business effective date

    # Builds a dataframe for the given table code
    # The hierarchy matches Business Auto: lower-level company (NACO/NAFF/NICOF)
    # first, since a company-specific filing should override the default;
    # then NGIC (the state-level default company); then CW as the final
    # country-wide fallback.
    # Returns the dataframe that was built
    def buildDataFrame(self, tableCode):
        if 'NACO' in self.rateTables.keys(): # Checking if NACO file was given
            if tableCode in self.rateTables['NACO'].keys(): # Checking if table exists in NACO
                return pd.DataFrame(data=self.rateTables['NACO'][tableCode][1:], index=None, columns=self.rateTables['NACO'][tableCode][0])
        if 'NAFF' in self.rateTables.keys(): # Checking if NAFF file was given
            if tableCode in self.rateTables['NAFF'].keys(): # Checking if table exists in NAFF
                return pd.DataFrame(data=self.rateTables['NAFF'][tableCode][1:], index=None, columns=self.rateTables['NAFF'][tableCode][0])
        if 'NICOF' in self.rateTables.keys(): # Checking if NICOF file was given
            if tableCode in self.rateTables['NICOF'].keys(): # Checking if table exists in NICOF
                return pd.DataFrame(data=self.rateTables['NICOF'][tableCode][1:], index=None, columns=self.rateTables['NICOF'][tableCode][0])
        if tableCode in self.rateTables['NGIC'].keys(): # Checking if table exists in the NGIC ratebook
            return pd.DataFrame(data=self.rateTables['NGIC'][tableCode][1:], index=None, columns=self.rateTables['NGIC'][tableCode][0])
        return pd.DataFrame(data=self.rateTables['CW'][tableCode][1:], index=None, columns=self.rateTables['CW'][tableCode][0]) # Returning the country-wide table if it wasn't found in any other company

    # Builds the Territory Multiplier factor table
    # Returns a dataframe
    def buildTerritoryMultiplier(self):
        TerritoryMultiplier = self.buildDataFrame("BP7_Peril_TerritorialFactor")
        filteredTerritoryMultiplier = TerritoryMultiplier.query(f'`Peril TypeCode` == "allperil"')
        return filteredTerritoryMultiplier.filter(items=['TerritoryCode', 'BldgTerritoryFactor', 'BPPTerritoryFactor', 'LiabilityTerritoryFactor', 'BITerritoryFactor']). \
            rename(columns={'TerritoryCode': 'Territory', 'BldgTerritoryFactor': 'Building', 'BPPTerritoryFactor': 'BPP', 'LiabilityTerritoryFactor': 'Liability', 'BITerritoryFactor': 'BI'})

    # Builds the sprinkler factor table
    # Returns a dataframe
    def buildSprinklerFactor(self):
        sprinklerPeril = self.buildDataFrame("BP7_Peril_Sprinkler_Discount")
        filteredSprinklerPeril = sprinklerPeril.query(f'`Peril TypeCode` == "allperil"').filter(items=['Class_Code_Min', 'Bldg Sprinkler Factor', 'BPP Sprinkler Factor'])
        filteredSprinklerPeril = filteredSprinklerPeril.replace({'Class_Code_Min': self.classCodes}).rename(columns={'Class_Code_Min': 'Program', 'Bldg Sprinkler Factor': 'Building', 'BPP Sprinkler Factor': 'BPP'}).sort_values(by=['Program'])
        return filteredSprinklerPeril

    # Builds the protection class table for the coverage (either Building or BPP)
    # Returns a dataframe
    def buildProtectionClass(self, coverage):
        protectionPeril = self.buildDataFrame("BP7_Peril_Protection_Class")
        filteredProtectionPeril = protectionPeril.query(f'`Peril TypeCode` == "allperil"').replace({'Class_Code_Min': self.classCodes, 'ProtectionClass': self.protectionClassConversions}).rename(columns={'ProtectionClass': 'Protection Class'})
        if coverage.casefold() == 'building': # Case-insensitive comparison
            pivotedProtectionPeril = filteredProtectionPeril.pivot(index='Protection Class', columns='Class_Code_Min', values='BldgProtectionClassFactor').reset_index('Protection Class')
            pivotedProtectionPeril['Protection Class Number'] = pivotedProtectionPeril['Protection Class'].apply(self.getProtectionClassValue)
            return pivotedProtectionPeril.sort_values(by=['Protection Class Number']).loc[:, pivotedProtectionPeril.columns != 'Protection Class Number']
        elif coverage.casefold() == 'bpp': # Case-insensitive comparison
            pivotedProtectionPeril = filteredProtectionPeril.pivot(index='Protection Class', columns='Class_Code_Min', values='BPPProtectionClassFactor').reset_index('Protection Class')
            pivotedProtectionPeril['Protection Class Number'] = pivotedProtectionPeril['Protection Class'].apply(self.getProtectionClassValue)
            return pivotedProtectionPeril.sort_values(by=['Protection Class Number']).loc[:, pivotedProtectionPeril.columns != 'Protection Class Number']

    # Builds the masonry veneer factor table for the given coverage (either Building or BPP)
    # Returns a dataframe
    def buildMasonryVeneer(self, coverage):
        masonryVeneerPeril = self.buildDataFrame("BP7_Peril_Masonry_Veneer")
        updatedMasonryVeneerPeril = masonryVeneerPeril.astype({'Masonry_Veneer_Min_Percent': 'int64', 'Masonry_Veneer_Max_Percent': 'int64'}). \
                astype({'Masonry_Veneer_Min_Percent': 'string', 'Masonry_Veneer_Max_Percent': 'string'}) # Converting to int first to get rid of decimal places
        updatedMasonryVeneerPeril["Masonry Veneer Percentage"] = updatedMasonryVeneerPeril["Masonry_Veneer_Min_Percent"] + ' - ' + updatedMasonryVeneerPeril["Masonry_Veneer_Max_Percent"] + '%' # Creating a single column for the percentage
        filteredMasonryVeneer = updatedMasonryVeneerPeril.query(f'`Peril TypeCode` == "allperil"')
        if coverage.casefold() == 'building': # Case-insensitive comparison
            return filteredMasonryVeneer.pivot(index='Masonry Veneer Percentage', columns='Peril TypeCode', values='Bldg_Masonry_Veneer_Factor').rename(columns={'allperil': 'All Programs'}).reset_index('Masonry Veneer Percentage')
        elif coverage.casefold() == 'bpp': # Case-insensitive comparison
            return filteredMasonryVeneer.pivot(index='Masonry Veneer Percentage', columns='Peril TypeCode', values='BPP_Masonry_Veneer_Factor').rename(columns={'allperil': 'All Programs'}).reset_index('Masonry Veneer Percentage')

    # Builds the property deductible table
    # Returns a dataframe
    def buildPropertyDeductible(self):
        propertyDedPeril = self.buildDataFrame("BP7_Peril_PropertyDeductible")
        filteredPropertyDeductible = propertyDedPeril.query(f'`Peril TypeCode` == "allperil"').replace({'Class_Code_Min': self.classCodes}). \
                rename(columns={'DeductibleAmount': 'Property Deductible', 'BPPTIB_AmtofInsurance_Min': 'BPP Min', 'BPPTIB_AmtOfInsurance_Max': 'BPP Max', 'BLDG_AmtOfInsurance': 'Building'})
        pivotedPropertyDeductible = filteredPropertyDeductible.filter(items=['Class_Code_Min', 'Property Deductible', 'BPP Min', 'BPP Max', 'Building', 'PropertyDeductibleFactor']). \
                pivot(index=['Property Deductible', 'BPP Min', 'BPP Max', 'Building'], columns='Class_Code_Min', values='PropertyDeductibleFactor').reset_index(['Property Deductible', 'BPP Min', 'BPP Max', 'Building'])
        sortedPropertyDeductible = pivotedPropertyDeductible.astype({'Property Deductible': 'int32'}).fillna({'BPP Max': '+'}).sort_values(by=['Property Deductible', 'BPP Min', 'BPP Max', 'Building'])
        propertyDeductibleDimensions = sortedPropertyDeductible.shape # A list where the first element represents the number of rows in the df and the second element represents the number of columns
        for row in range(propertyDeductibleDimensions[0]):
            for column in range(4, propertyDeductibleDimensions[1]): # Starting with column 4 because that is the first column with a peril
                if math.isnan(sortedPropertyDeductible.iloc[row, column]):
                    lowFactor = sortedPropertyDeductible.iloc[row - 1, column]
                    highFactor = sortedPropertyDeductible.iloc[row + 1, column]
                    lowBuildingAmount = sortedPropertyDeductible.iloc[row - 1, 3]
                    currentBuildingAmount = sortedPropertyDeductible.iloc[row, 3]
                    highBuildingAmount = sortedPropertyDeductible.iloc[row + 1, 3]
                    if math.isnan(lowFactor): # In case the low factor is missing too
                        lowFactor = sortedPropertyDeductible.iloc[row - 2, column]
                        lowBuildingAmount = sortedPropertyDeductible.iloc[row - 2, 3]
                    if math.isnan(highFactor): # In case the high factor is missing too
                        highFactor = sortedPropertyDeductible.iloc[row + 2, column]
                        highBuildingAmount = sortedPropertyDeductible.iloc[row + 2, 3]
                    sortedPropertyDeductible.iloc[row, column] = lowFactor + (highFactor - lowFactor) * ((currentBuildingAmount - lowBuildingAmount) / (highBuildingAmount - lowBuildingAmount))
        return sortedPropertyDeductible

    # Builds the wind/hail deductible factor table for the given coverage (either Building or BPP)
    # Returns a dataframe
    def buildWHDeductibleFactor(self, coverage):
        whDedPeril = self.buildDataFrame("BP7_Peril_WH_Deductible_Factor")
        if coverage.casefold() == 'building': # Case-insensitive comparison
            filteredWHDeductibleFactor = whDedPeril.query(f'`Peril TypeCode` == "allperil" & Coverage == "Building"').replace({'Class_Code_Min': self.classCodes}). \
                    rename(columns={'Class_Code_Min': 'Program', 'BPPTIB_AmtofInsurance_Min': 'BPP Min', 'BPPTIB_AmtofInsurance_Max': 'BPP Max', 'BLDG_AmtofInsurance': 'Building'})
        elif coverage.casefold() == 'bpp': # Case-insensitive comparison
            filteredWHDeductibleFactor = whDedPeril.query(f'`Peril TypeCode` == "allperil" & Coverage == "BPP"').replace({'Class_Code_Min': self.classCodes}). \
                    rename(columns={'Class_Code_Min': 'Program', 'BPPTIB_AmtofInsurance_Min': 'BPP Min', 'BPPTIB_AmtofInsurance_Max': 'BPP Max', 'BLDG_AmtofInsurance': 'Building'})
        updatedWHDeductibleFactor = filteredWHDeductibleFactor[filteredWHDeductibleFactor['WH_PercentOrAmount'].notnull()] # Filtering out the missing values from the WH_PercentOrAmount column
        sortedWHDeductibleFactor = updatedWHDeductibleFactor.fillna({'BPP Max': '+'}).sort_values(by=['Program', 'BPP Min', 'BPP Max', 'Building']) # Filling in missing values and sorting
        return sortedWHDeductibleFactor.astype({'WH_PercentOrAmount': 'int32'}).pivot(index=['Program', 'BPP Min', 'BPP Max', 'Building'], columns='WH_PercentOrAmount', values='WH Factor').reset_index(['Program', 'BPP Min', 'BPP Max', 'Building'])

    # Builds the wind/hail deductible per building factor table for the given coverage (either Building or BPP)
    # Returns a dataframe
    def buildWHDeductiblePerBuilding(self, coverage):
        whDedBldgPeril = self.buildDataFrame("BP7 Peril_WH_Deductible_Per_Building")
        if coverage.casefold() == 'building': # Case-insensitive comparison
            filteredWHDedBldg = whDedBldgPeril.query(f'{"`Peril TypeCode`".strip()} == "allperil" & Coverage == "Building" & WHDeductibleAmt != "1" & WHDeductibleAmt != "2" & WHDeductibleAmt != "3" & WHDeductibleAmt != "4" & WHDeductibleAmt != "5"'). \
                    replace({'Class_Code_Min': self.classCodes}).rename(columns={'WHDeductibleAmt': 'Wind / Hail Deductible Amount', 'AmtOfInsurance': 'Amount of Insurance'})
        elif coverage.casefold() == 'bpp': # Case-insensitive comparison
            filteredWHDedBldg = whDedBldgPeril.query(f'{"`Peril TypeCode`".strip()} == "allperil" & Coverage == "BPP" & WHDeductibleAmt != "1" & WHDeductibleAmt != "2" & WHDeductibleAmt != "3" & WHDeductibleAmt != "4" & WHDeductibleAmt != "5"'). \
                    replace({'Class_Code_Min': self.classCodes}).rename(columns={'WHDeductibleAmt': 'Wind / Hail Deductible Amount', 'AmtOfInsurance': 'Amount of Insurance'})
        return filteredWHDedBldg.pivot(index=['Wind / Hail Deductible Amount', 'Amount of Insurance'], columns='Class_Code_Min', values='WH_Deductible_Building_Factor').reset_index(['Wind / Hail Deductible Amount', 'Amount of Insurance']). \
                astype({'Wind / Hail Deductible Amount': 'int32'}).sort_values(by=['Wind / Hail Deductible Amount', 'Amount of Insurance'])

    # Builds the wind/hail percentage deductible factor table for the given coverage (either Building or BPP)
    # Returns a dataframe
    def buildWHDeductiblePercentage(self, coverage):
        whDedPeril = self.buildDataFrame("BP7 Peril_WH_Deductible_Per_Building")
        whDedPeril['Wind / Hail Deductible'] = whDedPeril['WHDeductibleAmt'] + '%'
        if coverage.casefold() == 'building': # Case-insensitive comparison
            filteredWHDeductiblePercentage = whDedPeril.query(f'`Peril TypeCode` == "allperil" & Coverage == "Building" & (WHDeductibleAmt == "1" | WHDeductibleAmt == "2" | WHDeductibleAmt == "3" | WHDeductibleAmt == "4" | WHDeductibleAmt == "5")'). \
                    rename(columns={'AmtOfInsurance': 'Amount of Insurance'})
        elif coverage.casefold() == 'bpp': # Case-insensitive comparison
            filteredWHDeductiblePercentage = whDedPeril.query(f'`Peril TypeCode` == "allperil" & Coverage == "BPP" & (WHDeductibleAmt == "1" | WHDeductibleAmt == "2" | WHDeductibleAmt == "3" | WHDeductibleAmt == "4" | WHDeductibleAmt == "5")'). \
                    rename(columns={'AmtOfInsurance': 'Amount of Insurance'})
        return filteredWHDeductiblePercentage.replace({'Class_Code_Min': self.classCodes}).pivot(index=['Wind / Hail Deductible', 'Amount of Insurance'], columns='Class_Code_Min', values='WH_Deductible_Building_Factor').reset_index(['Wind / Hail Deductible', 'Amount of Insurance'])

    # Builds the burglar alarm factor table
    # Returns a dataframe
    def buildBurglarAlarmFactor(self):
        burglarAlarmPeril = self.buildDataFrame("BP7_Peril_Burglar_Alarm_Factor")
        filteredBurglarAlarmPeril = burglarAlarmPeril.query(f'`Peril TypeCode` == "allperil" & (`Burglar Alarm Type` == "local" | `Burglar Alarm Type` == "central")').replace({'Class_Code_Min': self.classCodes}). \
            rename(columns={'Class_Code_Min': 'Program'})
        return filteredBurglarAlarmPeril.pivot(index='Program', columns='Burglar Alarm Type', values='Burglar Alarm Factor').rename(columns={'central': 'Central Station Alarm', 'local': 'Local Alarm'}).reset_index('Program')

    # Builds the fire alarm factor table
    # Returns a dataframe
    def buildFireAlarmFactor(self):
        fireAlarmPeril = self.buildDataFrame("BP7_Peril_Fire_Alarm_Factor")
        fireAlarmPeril = fireAlarmPeril.query(f'`Peril TypeCode` == "allperil" & FireAlarmType == "Central Station Fire Alarm"').replace({'Peril TypeCode': {'allperil': 'All Programs'}}).rename(columns={'Peril TypeCode': 'Program'})
        fireAlarmPeril = fireAlarmPeril.pivot(index='Program', columns='Coverage', values='FireAlarmFactor').rename(columns={'Business Income': 'Bus Inc'})
        return fireAlarmPeril

    # Builds the building age modifier table for the given coverage (either Building, BPP, or Business Income)
    # Returns a dataframe
    def buildBuildingAgeModifier(self, coverage):
        bldgAgePeril = self.buildDataFrame("BP7 Peril Building_Age_Modifier")
        if coverage.casefold() == 'building': # Case-insensitive comparison
            filteredBldgAgePeril = bldgAgePeril.query(f'`Peril TypeCode` == "allperil" & Coverage == "Building"')
        elif coverage.casefold() == 'bpp': # Case-insensitive comparison
            filteredBldgAgePeril = bldgAgePeril.query(f'`Peril TypeCode` == "allperil" & Coverage == "BPP"')
        elif coverage.casefold() == 'business income': # Case-insensitive comparison
            filteredBldgAgePeril = bldgAgePeril.query(f'`Peril TypeCode` == "allperil" & Coverage == "Business Income"')
        return filteredBldgAgePeril.replace({'Class_Code_Min': self.classCodes, 'Building_Age_Min': {101: '101-1000'}}).rename(columns={'Building_Age_Min': 'Building Age Years'}). \
            pivot(index='Building Age Years', columns='Class_Code_Min', values='BuildingAge_Factor').reset_index('Building Age Years')

    # Builds the building AOI table
    # Returns a dataframe
    def buildBuildingAOI(self):
        aoiBldgPeril = self.buildDataFrame("BP7_Peril_Building_Amt_Insurance")
        filteredAOIBldgPeril = aoiBldgPeril.query(f'`Peril TypeCode` == "allperil" & Building_Limit < 10000000').replace({'Class_Code_Min': self.classCodes}).rename(columns={'Building_Limit': 'Lower Bound'}). \
            pivot(index='Lower Bound', columns='Class_Code_Min', values='AmountOfInsuranceFactor').reset_index('Lower Bound')
        filteredAOIBldgPeril['Upper Bound'] = filteredAOIBldgPeril['Lower Bound'] - 1 # Creating the upper bound column
        filteredAOIBldgPeril['Upper Bound'] = filteredAOIBldgPeril['Upper Bound'].shift(periods=-1, fill_value='and over') # Shifting the upper bound column to align it with lower bound
        cols = filteredAOIBldgPeril.columns.tolist()
        cols = cols[0:1] + cols[-1:] + cols[1:-1] # Rearranging the order of columns
        return filteredAOIBldgPeril[cols]

    # Builds the BPP AOI table
    # Returns a dataframe
    def buildBPPAOI(self):
        aoiBPPPeril = self.buildDataFrame("BP7_Peril_BPP_Amt_Insurance")
        filteredAOIBPPPeril = aoiBPPPeril.query(f'`Peril TypeCode` == "allperil" & BPP_Limit < 10000000').replace({'Class_Code_Min': self.classCodes}).rename(columns={'BPP_Limit': 'Lower Bound'}). \
            pivot(index='Lower Bound', columns='Class_Code_Min', values='AmountOfInsuranceFactor').reset_index('Lower Bound')
        filteredAOIBPPPeril['Upper Bound'] = filteredAOIBPPPeril['Lower Bound'] - 1 # Creating the upper bound column
        filteredAOIBPPPeril['Upper Bound'] = filteredAOIBPPPeril['Upper Bound'].shift(periods=-1, fill_value='and over') # Shifting the upper bound column to align it with lower bound
        cols = filteredAOIBPPPeril.columns.tolist()
        cols = cols[0:1] + cols[-1:] + cols[1:-1] # Rearranging the order of columns
        return filteredAOIBPPPeril[cols]

    # Builds the building code effectiveness grade (BCEG) table
    # Returns a dataframe
    def buildBCEG(self):
        bceg = self.buildDataFrame("BP7_Peril_BCEG_Factor")
        filteredBCEG = bceg.query(f'`Peril TypeCode` == "allperil"').copy() # Creating a new copy of the data here
        filteredBCEG['Grade'] = filteredBCEG['Tier Grade'] # Creating a new column for the grade and using the old one for sorting
        updatedBCEG = filteredBCEG.astype({'Tier Grade': 'int32'}).replace({'Grade': {'98': 'Non-Particip', '99': 'Ungraded'}})
        if self.state in self.buildingCodes: # Multiple building codes, so different table format
            updatedBCEG['Territory Group'] = updatedBCEG['Territory Code'].apply(self.getBuildingCode)
            finalBCEG = updatedBCEG.pivot(index=['Territory Group', 'Territory Code', 'Grade', 'Tier Grade'], columns='Peril TypeCode', values='BCEGFactor'). \
                    reset_index(['Territory Group', 'Territory Code', 'Grade', 'Tier Grade']).groupby(by=['Territory Group', 'Grade', 'Tier Grade']).mean(numeric_only=True). \
                    reset_index(['Territory Group', 'Grade', 'Tier Grade']).sort_values(by=['Territory Group', 'Tier Grade']).rename(columns={'allperil': 'All Programs'})
            del finalBCEG['Tier Grade'] # Deleting the extra column used for sorting
            finalBCEG['ISO Territory'] = finalBCEG['Territory Group'].apply(self.getTerritoryCodes) # Creating a column for the territory
            cols = finalBCEG.columns.tolist()
            cols = cols[0:1] + cols[-1:] + cols[1:-1] # Rearranging the order of columns
            return finalBCEG[cols]
        finalBCEG = updatedBCEG.pivot(index=['Grade', 'Tier Grade'], columns='Peril TypeCode', values='BCEGFactor').reset_index(['Grade', 'Tier Grade']).sort_values(by=['Tier Grade']).rename(columns={'allperil': 'All Programs'})
        del finalBCEG['Tier Grade'] # Deleting the extra column used for sorting
        return finalBCEG

    # Builds the tenants improvements and betterments factor table
    # Returns a dataframe
    def buildTenantsImprovements(self):
        tenantsImprovementsPeril = self.buildDataFrame("BP7_Peril_Tenants_Improvements_and_Betterments_Factor")
        tenantsImprovementsPeril = tenantsImprovementsPeril.query(f'`Peril TypeCode` == "allperil"').replace({'Peril TypeCode': {'allperil': 'All Programs'}}).rename(columns={'Peril TypeCode': 'Program', 'TIBFactor': 'Factor'}). \
                filter(items=['Program', 'Factor'])
        return tenantsImprovementsPeril

    # Calculates the "value" of each protection class by multiplying the class number by 200 and adds the Unicode code of the letter if there is one
    # Used for sorting the protection class table
    # Returns the value associated with the given protection class
    def getProtectionClassValue(self, protectionClass):
        if len(protectionClass) == 3:
            return int(protectionClass[:-1]) * 200 + ord(protectionClass[-1])
        elif len(protectionClass) == 2:
            if protectionClass == '10':
                return int(protectionClass) * 200
            return int(protectionClass[:-1]) * 200 + ord(protectionClass[-1])
        return int(protectionClass) * 200

    # Returns the building code (A, B, C or D) from the given territory code
    # Varies by state
    def getBuildingCode(self, territoryCode):
        for buildingCode in self.buildingCodes[self.state]:
            if territoryCode in self.buildingCodes[self.state][buildingCode]:
                return buildingCode

    # Returns the list of territories associated with the given building code separated by a comma
    def getTerritoryCodes(self, buildingCode):
        return ', '.join(self.buildingCodes[self.state][buildingCode])

    # Sets up the All Peril Excel file and creates a separate worksheet for each of the given dataframes
    # progress_callback (optional) is called with a short message before each sheet is built
    # Returns the Excel workbook
    def buildAllPerilPage(self, progress_callback=None):
        companies = [c for c in self.rateTables.keys() if c != 'CW'] # country-wide is not a company, so ignoring it

        AllPeril = ExcelSettingsBOP.Excel(state=self.state, programName='All Peril', nEffective=self.nEffective, rEffective=self.rEffective, companyList=companies)

        # BCEG's column layout depends on whether this state has multiple
        # building-code groups. The old root-level AllPerilPage.formatBCEG had
        # a TODO for multi-code states; here that case reuses the BCEG_MULTI
        # profile All Programs already defines (the table shape is the same).
        bcegLayoutKey = 'BCEG_MULTI' if self.state in self.buildingCodes else 'BCEG_AP_SINGLE'

        # (tab name, page title, builder callable, useIndex, useHeader, layout_key)
        # layout_key=None reuses the All Programs profile of the same code
        # (transcribed widths matched exactly); *_AP keys are the All
        # Peril-specific profiles in BOP Input File.xlsx.
        sheetSpecs = [
            ('SPR', 'AS, FS, H, O, R, S, W Table 3.C.2.a. Sprinkler Factor', self.buildSprinklerFactor, False, True, None),
            ('PCBG', 'AS, FS, H, O, R, S, W Table 3.C.2.b. Protection Class Factor - Building', lambda: self.buildProtectionClass('Building'), False, True, 'PCBG_AP'),
            ('PCPP', 'AS, FS, H, O, R, S, W Table 3.C.2.b. Protection Class Factor - BPP', lambda: self.buildProtectionClass('BPP'), False, True, 'PCPP_AP'),
            ('MVBG', 'AS, FS, H, O, R, S, W Table 3.C.2.c.1. Masonry Veneer Factor - Building', lambda: self.buildMasonryVeneer('Building'), False, True, 'MVBG_AP'),
            ('MVPP', 'AS, FS, H, O, R, S, W Table 3.C.2.c.1. Masonry Veneer Factor - BPP', lambda: self.buildMasonryVeneer('BPP'), False, True, 'MVPP_AP'),
            ('PD', 'AS, FS, H, O, R, S, W Table 3.C.2.f. Property Deductible Factor', self.buildPropertyDeductible, False, True, 'PD_AP'),
            ('WHOBG', 'AS, FS, H, O, R, S, W Table 3.C.2.g.(1). Windstorm or Hail Deductible Factor - Per Occurrence Fixed Deductible Amount - Building', lambda: self.buildWHDeductibleFactor('Building'), False, True, None),
            ('WHOPP', 'AS, FS, H, O, R, S, W Table 3.C.2.g.(1). Windstorm or Hail Deductible Factor - Per Occurrence Fixed Deductible Amount - BPP', lambda: self.buildWHDeductibleFactor('BPP'), False, True, None),
            ('WHBBG', 'AS, FS, H, O, R, S, W Table 3.C.2.g.(2). Windstorm or Hail Deductible Factor - Per Building Fixed Deductible Amount - Building', lambda: self.buildWHDeductiblePerBuilding('Building'), False, True, 'WHBBG_AP'),
            ('WHBPP', 'AS, FS, H, O, R, S, W Table 3.C.2.g.(2). Windstorm or Hail Deductible Factor - Per Building Fixed Deductible Amount - BPP', lambda: self.buildWHDeductiblePerBuilding('BPP'), False, True, 'WHBPP_AP'),
            ('WHPBG', 'AS, FS, H, O, R, S, W Table 3.C.2.g.(3). Windstorm or Hail Deductible Factor - Percentage Deductible - Building', lambda: self.buildWHDeductiblePercentage('Building'), False, True, 'WHPBG_AP'),
            ('WHPPP', 'AS, FS, H, O, R, S, W Table 3.C.2.g.(3). Windstorm or Hail Deductible Factor - Percentage Deductible - BPP', lambda: self.buildWHDeductiblePercentage('BPP'), False, True, 'WHPPP_AP'),
            ('BA', 'AS, FS, H, O, R, S, W Table 3.C.2.h. Burglar Alarm Modifier', self.buildBurglarAlarmFactor, False, True, None),
            ('CSFA', 'AS, FS, H, O, R, S, W Table 3.C.2.i. Central Station Fire Alarm Modifier', self.buildFireAlarmFactor, True, True, 'CSFA_AP'),
            ('BABG', 'AS, FS, H, O, R, S, W Table 3.C.2.j. Building Age Modifier - Building', lambda: self.buildBuildingAgeModifier('Building'), False, True, 'BABG_AP'),
            ('BAPP', 'AS, FS, H, O, R, S, W Table 3.C.2.j. Building Age Modifier - BPP', lambda: self.buildBuildingAgeModifier('BPP'), False, True, 'BAPP_AP'),
            ('BABI', 'FS Table 3.C.2.j. Building Age Modifier - Bus Inc', lambda: self.buildBuildingAgeModifier('Business Income'), False, True, 'BABI_AP'), # Business Income is only applicable in the food program
            ('AIBG', 'AS, FS, H, O, R, S, W Table 3.C.2.l.(1). AOI (Amount of Insurance) Relativity Factor - Building', self.buildBuildingAOI, False, True, 'AIBG_AP'),
            ('AIPP', 'AS, FS, H, O, R, S, W Table 3.C.2.l.(1). AOI (Amount of Insurance) Relativity Factor - BPP', self.buildBPPAOI, False, True, 'AIPP_AP'),
            ('BCEG', 'H Table 3.C.2.n., AS, FS, O, R, S, W Table 3.C.2.o. Building Code Effectiveness Grading', self.buildBCEG, False, True, bcegLayoutKey),
            ('TIB', 'H Table 3.C.2.p., AS, FS, O, R, S, W Table 3.C.2.q. Tenants Improvements and Betterments Factor', self.buildTenantsImprovements, False, True, 'TIB_AP'),
            ('TR', 'AS, FS, H, O, R, S, W Table 3.C.1.a. State Territory Multiplier - All Perils', self.buildTerritoryMultiplier, False, True, 'TR_AP'),
        ]

        total = len(sheetSpecs)
        for i, (tableCode, title, build, useIndex, useHeader, layoutKey) in enumerate(sheetSpecs, start=1):
            if progress_callback:
                progress_callback(f"Building sheet {i}/{total}: {tableCode}...")
            print(f"  [{i}/{total}] Building sheet: {tableCode}")
            AllPeril.generateWorksheet(tableCode, title, build(), useIndex, useHeader, layout_key=layoutKey)

        if progress_callback:
            progress_callback("Building Index sheet...")
        print(f"  [{total}/{total}] Building sheet: Index")
        AllPeril.createIndex()
        return AllPeril.getWB()
