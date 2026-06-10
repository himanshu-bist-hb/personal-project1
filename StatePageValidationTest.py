import pandas as pd

class ValidationTest:
    def __init__(self, rateTables, perils, programCodes) -> None:
        self.rateTables = rateTables
        self.perils = perils
        self.programCodes = programCodes

    # Builds a dataframe for the given table code
    # The hierarchy is as follows: NGIC > Migration > CW
    # Returns the dataframe that was built
    def buildDataFrame(self, tableCode):
        if tableCode in self.rateTables['NGIC'].keys(): # Checking if table exists in the NGIC ratebook
            return pd.DataFrame(data=self.rateTables['NGIC'][tableCode][1:], index=None, columns=self.rateTables['NGIC'][tableCode][0])
        if 'NACO' in self.rateTables.keys(): # Checking if NACO file was given
            if tableCode in self.rateTables['NACO'].keys(): # Checking if table exists in NACO
                return pd.DataFrame(data=self.rateTables['NACO'][tableCode][1:], index=None, columns=self.rateTables['NACO'][tableCode][0])
        if 'NAFF' in self.rateTables.keys(): # Checking if NAFF file was given
            if tableCode in self.rateTables['NAFF'].keys(): # Checking if table exists in NAFF
                return pd.DataFrame(data=self.rateTables['NAFF'][tableCode][1:], index=None, columns=self.rateTables['NAFF'][tableCode][0])
        if 'NICOF' in self.rateTables.keys(): # Checking if NICOF file was given
            if tableCode in self.rateTables['NICOF'].keys(): # Checking if tabl exists in NICOF
                return pd.DataFrame(data=self.rateTables['NICOF'][tableCode][1:], index=None, columns=self.rateTables['NICOF'][tableCode][0])
        return pd.DataFrame(data=self.rateTables['CW'][tableCode][1:], index=None, columns=self.rateTables['CW'][tableCode][0]) # Returning the country-wide table if it wasn't found in any other company

    # Sorts the given dataframe by the given parameters and resets the index so the dataframe can be compared to others
    # Returns a new dataframe is returned
    def cleanDataFrame(self, df, sortingArray):
        return df.sort_values(by=sortingArray).reset_index(drop=True) # Dropping the old index since it's meaningless

    # Builds the sprinkler factor table for the given program
    # Returns a dataframe
    def buildSprinklerFactor(self, program):
        sprinklerPeril = self.buildDataFrame("BP7_Peril_Sprinkler_Discount")
        return sprinklerPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'Bldg Sprinkler Factor', 'BPP Sprinkler Factor'])

    # Builds the protection class table for the given program
    # Returns a dataframe
    def buildProtectionClass(self, program):
        protectionPeril = self.buildDataFrame("BP7_Peril_Protection_Class")
        return protectionPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'ProtectionClass', 'BldgProtectionClassFactor', 'BPPProtectionClassFactor'])

    # Builds the property deductible table for the given program
    # Returns a dataframe
    def buildPropertyDeductible(self, program):
        propertyDedPeril = self.buildDataFrame("BP7_Peril_PropertyDeductible")
        return propertyDedPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'DeductibleAmount', 'BPPTIB_AmtofInsurance_Min', 'BPPTIB_AmtOfInsurance_Max', 'BLDG_AmtOfInsurance', 'PropertyDeductibleFactor'])

    # Builds the wind/hail deductible factor table for the given program
    # Returns a dataframe
    def buildWHDeductibleFactor(self, program):
        whDedPeril = self.buildDataFrame("BP7_Peril_WH_Deductible_Factor")
        return whDedPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'Coverage', 'BPPTIB_AmtofInsurance_Min', 'BPPTIB_AmtofInsurance_Max', 'BLDG_AmtofInsurance', 'WH_PercentOrAmount', 'WH Factor'])

    # Builds the wind/hail deductible per building factor table for the given program
    # Returns a dataframe
    def buildWHDeductiblePerBuilding(self, program):
        whDedBldgPeril = self.buildDataFrame("BP7 Peril_WH_Deductible_Per_Building")
        return whDedBldgPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'WHDeductibleAmt', 'Coverage', 'AmtOfInsurance', 'WH_Deductible_Building_Factor'])

    # Builds the burglar alarm factor table for the given program
    # Returns a dataframe
    def buildBurglarAlarmFactor(self, program):
        burglarAlarmPeril = self.buildDataFrame("BP7_Peril_Burglar_Alarm_Factor")
        return burglarAlarmPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'Burglar Alarm Type', 'Burglar Alarm Factor'])

    # Builds the building age modifier table for the given program
    # Returns a dataframe
    def buildBuildingAgeModifier(self, program):
        bldgAgePeril = self.buildDataFrame("BP7 Peril Building_Age_Modifier")
        return bldgAgePeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils} & Coverage != "Business Income"').filter(items=['Peril TypeCode', 'Coverage', 'Building_Age_Min', 'Building_Age_Max', 'BuildingAge_Factor'])

    # Builds the building AOI table for the given program
    # Returns a dataframe
    def buildBuildingAOI(self, program):
        aoiBldgPeril = self.buildDataFrame("BP7_Peril_Building_Amt_Insurance")
        return aoiBldgPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'Building_Limit', 'AmountOfInsuranceFactor'])

    # Builds the BPP AOI table for the given program
    # Returns a dataframe
    def buildBPPAOI(self, program):
        aoiBPPPeril = self.buildDataFrame("BP7_Peril_BPP_Amt_Insurance")
        return aoiBPPPeril.query(f'Class_Code_Min == {self.programCodes[program]} & `Peril TypeCode` in {self.perils}').filter(items=['Peril TypeCode', 'BPP_Limit', 'AmountOfInsuranceFactor'])

    # Builds the equipment breakdown limit relativity table for the given program
    # Returns a dataframe
    def buildEBLimitRelativity(self, program):
        ebLimitRelativity = self.buildDataFrame("BP7_EBLimitsRelativityModifier")
        return ebLimitRelativity.query(f'Class_Code_Min == {self.programCodes[program]}').filter(items=['TotalPropertyLimitMin', 'TotalPropertyLimitMax', 'LimitRelativityModifier'])

    # Builds the equipment breakdown deductible factor table for the given program
    # Returns a dataframe
    def buildEBDeductibleFactor(self, program):
        ebDeductibleFactor = self.buildDataFrame("BP7_EBDeductibleFactor")
        return ebDeductibleFactor.query(f'Class_Code_Min == {self.programCodes[program]}').filter(items=['DeductibleAmt', 'Factor'])

    # Compares the given 7 tables (from each of the 7 programs), in a dataframe format, against each other for equality
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareTables(self, autoTable, foodTable, habTable, officeTable,  serviceTable, retailTable, wholesaleTable):
        return autoTable.equals(foodTable) & autoTable.equals(habTable) & autoTable.equals(officeTable) & autoTable.equals(serviceTable) & autoTable.equals(retailTable) & autoTable.equals(wholesaleTable) \
                & foodTable.equals(habTable) & foodTable.equals(officeTable) & foodTable.equals(serviceTable) & foodTable.equals(retailTable) & foodTable.equals(wholesaleTable) \
                & habTable.equals(officeTable) & habTable.equals(serviceTable) & habTable.equals(retailTable) & habTable.equals(wholesaleTable) \
                & officeTable.equals(serviceTable) & officeTable.equals(retailTable) & officeTable.equals(wholesaleTable) \
                & serviceTable.equals(retailTable) & serviceTable.equals(wholesaleTable) \
                & retailTable.equals(wholesaleTable)

    # Compares all 7 sprinkler factor tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareSprinklerFactor(self):
        sortingArray = ['Peril TypeCode']
        return self.compareTables(self.cleanDataFrame(self.buildSprinklerFactor('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildSprinklerFactor('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildSprinklerFactor('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildSprinklerFactor('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildSprinklerFactor('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildSprinklerFactor('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildSprinklerFactor('Wholesale'), sortingArray))

    # Compares all 7 protection class tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareProtectionClass(self):
        sortingArray = ['Peril TypeCode', 'ProtectionClass']
        return self.compareTables(self.cleanDataFrame(self.buildProtectionClass('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildProtectionClass('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildProtectionClass('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildProtectionClass('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildProtectionClass('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildProtectionClass('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildProtectionClass('Wholesale'), sortingArray))

    # Compares all 7 property deductible tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def comparePropertyDeductible(self):
        sortingArray = ['Peril TypeCode', 'DeductibleAmount', 'BPPTIB_AmtofInsurance_Min', 'BPPTIB_AmtOfInsurance_Max', 'BLDG_AmtOfInsurance']
        return self.compareTables(self.cleanDataFrame(self.buildPropertyDeductible('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildPropertyDeductible('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildPropertyDeductible('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildPropertyDeductible('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildPropertyDeductible('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildPropertyDeductible('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildPropertyDeductible('Wholesale'), sortingArray))

    # Compares all 7 wind/hail deductible factor tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareWHDeductibleFactor(self):
        sortingArray = ['Peril TypeCode', 'Coverage', 'BPPTIB_AmtofInsurance_Min', 'BPPTIB_AmtofInsurance_Max', 'BLDG_AmtofInsurance', 'WH_PercentOrAmount']
        return self.compareTables(self.cleanDataFrame(self.buildWHDeductibleFactor('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductibleFactor('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductibleFactor('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductibleFactor('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductibleFactor('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductibleFactor('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductibleFactor('Wholesale'), sortingArray))

    # Compares all 7 wind/hail deductible per building tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareWHDeductiblePerBuilding(self):
        sortingArray = ['Peril TypeCode', 'WHDeductibleAmt', 'Coverage', 'AmtOfInsurance']
        return self.compareTables(self.cleanDataFrame(self.buildWHDeductiblePerBuilding('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductiblePerBuilding('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductiblePerBuilding('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductiblePerBuilding('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductiblePerBuilding('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductiblePerBuilding('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildWHDeductiblePerBuilding('Wholesale'), sortingArray))

    # Compares all 7 burglar alarm factor tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareBurglarAlarmFactor(self):
        sortingArray = ['Peril TypeCode', 'Burglar Alarm Type']
        return self.compareTables(self.cleanDataFrame(self.buildBurglarAlarmFactor('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildBurglarAlarmFactor('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildBurglarAlarmFactor('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildBurglarAlarmFactor('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildBurglarAlarmFactor('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildBurglarAlarmFactor('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildBurglarAlarmFactor('Wholesale'), sortingArray))

    # Compares all 7 building age modifier tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareBuildingAgeModifier(self):
        sortingArray = ['Peril TypeCode', 'Coverage', 'Building_Age_Min', 'Building_Age_Max']
        return self.compareTables(self.cleanDataFrame(self.buildBuildingAgeModifier('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAgeModifier('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAgeModifier('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAgeModifier('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAgeModifier('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAgeModifier('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAgeModifier('Wholesale'), sortingArray))

    # Compares all 7 Building AOI (amount of insurance) tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareBuildingAOI(self):
        sortingArray = ['Peril TypeCode', 'Building_Limit']
        return self.compareTables(self.cleanDataFrame(self.buildBuildingAOI('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAOI('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAOI('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAOI('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAOI('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAOI('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildBuildingAOI('Wholesale'), sortingArray))

    # Compares all 7 BPP AOI (amount of insurance) tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareBPPAOI(self):
        sortingArray = ['Peril TypeCode', 'BPP_Limit']
        return self.compareTables(self.cleanDataFrame(self.buildBPPAOI('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildBPPAOI('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildBPPAOI('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildBPPAOI('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildBPPAOI('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildBPPAOI('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildBPPAOI('Wholesale'), sortingArray))

    # Compares all 7 equipment breakdown limit relativity tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareEBLimitRelativity(self):
        sortingArray = ['TotalPropertyLimitMin', 'TotalPropertyLimitMax']
        return self.compareTables(self.cleanDataFrame(self.buildEBLimitRelativity('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildEBLimitRelativity('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildEBLimitRelativity('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildEBLimitRelativity('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildEBLimitRelativity('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildEBLimitRelativity('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildEBLimitRelativity('Wholesale'), sortingArray))

    # Compares all 7 equipment breakdown deductible factor tables
    # Returns True IF AND ONLY IF all tables are found equal to each other
    # Otherwise, returns False
    def compareEBDeductibleFactor(self):
        sortingArray = ['DeductibleAmt']
        return self.compareTables(self.cleanDataFrame(self.buildEBDeductibleFactor('Auto'), sortingArray),
                                  self.cleanDataFrame(self.buildEBDeductibleFactor('Food'), sortingArray),
                                  self.cleanDataFrame(self.buildEBDeductibleFactor('Hab'), sortingArray),
                                  self.cleanDataFrame(self.buildEBDeductibleFactor('Office'), sortingArray),
                                  self.cleanDataFrame(self.buildEBDeductibleFactor('Retail'), sortingArray),
                                  self.cleanDataFrame(self.buildEBDeductibleFactor('Service'), sortingArray),
                                  self.cleanDataFrame(self.buildEBDeductibleFactor('Wholesale'), sortingArray))

    def compareCompanies(self):
        allTablesEqual = True
        tablesExcluded = ['BP7EarthquakeDivisionFiveLossCostMultiplier', 'BP7LiabilityChargesForRelatedAddnlExposures', 'BP7PerilBldgHabitationalSwimmingPoolsPropertyBaseRate', 'BP7_Peril_BPP_Base_Rates', 'BP7_Peril_Building_Base_Rates',
                          'BP7_Peril_Business_Income_Base_Rate', 'BP7_Peril_Liability_Base_Rates', 'BP7_CompanyCode'] # Excluding certain tables from comparison since they are expected to be different
        tablesTested = set() # Using a set to avoid duplicate entries to keep track of which tables have been tested
        # Comparing NGIC tables to migration company tables
        for tableCode in self.rateTables['NGIC'].keys():
            if tableCode in tablesExcluded: # Skipping excluded tables
                continue
            ngicTable = pd.DataFrame(data=self.rateTables['NGIC'][tableCode][1:], index=None, columns=self.rateTables['NGIC'][tableCode][0])
            ngicTable = ngicTable.sort_values(by=ngicTable.columns.values.tolist())
            if 'NACO' in self.rateTables.keys(): # Checking if NACO file was given
                if tableCode in self.rateTables['NACO'].keys(): # Checking if table exists in NACO
                    nacoTable = pd.DataFrame(data=self.rateTables['NACO'][tableCode][1:], index=None, columns=self.rateTables['NACO'][tableCode][0])
                    nacoTable = nacoTable.sort_values(by=nacoTable.columns.values.tolist())
                    tablesTested.add(tableCode)
                    if not ngicTable.equals(nacoTable):
                        print(f'{tableCode} is not equal between NGIC and NACO')
                        print()
                        allTablesEqual = False
            if 'NAFF' in self.rateTables.keys(): # Checking if NAFF file was given
                if tableCode in self.rateTables['NAFF'].keys(): # Checking if table exists in NAFF
                    naffTable = pd.DataFrame(data=self.rateTables['NAFF'][tableCode][1:], index=None, columns=self.rateTables['NAFF'][tableCode][0])
                    naffTable = naffTable.sort_values(by=naffTable.columns.values.tolist())
                    tablesTested.add(tableCode)
                    if not ngicTable.equals(naffTable):
                        print(f'{tableCode} is not equal between NGIC and NAFF')
                        print()
                        allTablesEqual = False
            if 'NICOF' in self.rateTables.keys(): # Checking if NICOF file was given
                if tableCode in self.rateTables['NICOF'].keys(): # Checking if table exists in NICOF
                    nicofTable = pd.DataFrame(data=self.rateTables['NICOF'][tableCode][1:], index=None, columns=self.rateTables['NICOF'][tableCode][0])
                    nicofTable = nicofTable.sort_values(by=nicofTable.columns.values.tolist()) 
                    tablesTested.add(tableCode)
                    if not ngicTable.equals(nicofTable):
                        print(f'{tableCode} is not equal between NGIC and NICOF')
                        print()
                        allTablesEqual = False

        # Comparing NACO tables to NAFF and NICOF tables
        if 'NACO' in self.rateTables.keys():
            for tableCode in self.rateTables['NACO'].keys():
                if tableCode in tablesTested or tableCode in tablesExcluded: # Skipping any tables that have been tested already or excluded
                    continue
                nacoTable = pd.DataFrame(data=self.rateTables['NACO'][tableCode][1:], index=None, columns=self.rateTables['NACO'][tableCode][0])
                nacoTable = nacoTable.sort_values(by=nacoTable.columns.values.tolist())
                if 'NAFF' in self.rateTables.keys(): # Checking if NAFF file was given
                    if tableCode in self.rateTables['NAFF'].keys(): # Checking if table exists in NAFF
                        naffTable = pd.DataFrame(data=self.rateTables['NAFF'][tableCode][1:], index=None, columns=self.rateTables['NAFF'][tableCode][0])
                        naffTable = naffTable.sort_values(by=naffTable.columns.values.tolist())
                        if not nacoTable.equals(naffTable):
                            print(f'{tableCode} is not equal between NACO and NAFF')
                            print()
                            allTablesEqual = False
                if 'NICOF' in self.rateTables.keys(): # Checking if NICOF file was given
                    if tableCode in self.rateTables['NICOF'].keys(): # Checking if table exists in NICOF
                        nicofTable = pd.DataFrame(data=self.rateTables['NICOF'][tableCode][1:], index=None, columns=self.rateTables['NICOF'][tableCode][0])
                        nicofTable = nicofTable.sort_values(by=nicofTable.columns.values.tolist()) 
                        if not nacoTable.equals(nicofTable):
                            print(f'{tableCode} is not equal between NACO and NICOF')
                            print()
                            allTablesEqual = False

        # Comparing NAFF tables to NICOF tables
        if 'NAFF' in self.rateTables.keys() and 'NICOF' in self.rateTables.keys():
            for tableCode in self.rateTables['NAFF'].keys():
                if tableCode in tablesTested or tableCode in tablesExcluded: # Skipping any tables that have been tested already or excluded
                    continue
                naffTable = pd.DataFrame(data=self.rateTables['NAFF'][tableCode][1:], index=None, columns=self.rateTables['NAFF'][tableCode][0])
                naffTable = naffTable.sort_values(by=naffTable.columns.values.tolist())
                if tableCode in self.rateTables['NICOF'].keys():
                    nicofTable = pd.DataFrame(data=self.rateTables['NICOF'][tableCode][1:], index=None, columns=self.rateTables['NICOF'][tableCode][0])
                    nicofTable = nicofTable.sort_values(by=nicofTable.columns.values.tolist())
                    if not naffTable.equals(nicofTable):
                        print(f'{tableCode} is not equal between NAFF and NICOF')
                        print()
                        allTablesEqual = False
        
        if allTablesEqual:
            print('All tables are equal across all companies')
            print()

    def runValidationTest(self):
        print('**************************************************************')
        print('START OF CROSS PROGRAM VALIDATION')
        print('**************************************************************')
        print()

        print('Testing Sprinkler Factor tables...')
        if self.compareSprinklerFactor():
            print('Sprinkler Factor tables are all equal.')
        else:
            print('There is a difference in Sprinkler Factor tables.')
        print()

        print('Testing Protection Class tables...')
        if self.compareProtectionClass():
            print('Protection Class tables are all equal.')
        else:
            print('There is a difference in Protection Class tables.')
        print()

        print('Testing Property Deductible tables...')
        if self.comparePropertyDeductible():
            print('Property Deductible tables are all equal.')
        else:
            print('There is a difference in Property Deductible tables.')
        print()

        print('Testing Wind/Hail Deductible Factor tables...')
        if self.compareWHDeductibleFactor():
            print('Wind/Hail Deductible Factor tables are all equal.')
        else:
            print('There is a difference in Wind/Hail Deductible Factor tables.')
        print()

        print('Testing Wind/Hail Deductible Per Building Factor tables...')
        if self.compareWHDeductiblePerBuilding():
            print('Wind/Hail Deductible Per Building Factor tables are all equal.')
        else:
            print('There is a difference in Wind/Hail Deductible Per Building Factor tables.')
        print()

        print('Testing Burglar Alarm Factor tables...')
        if self.compareBurglarAlarmFactor():
            print('Burglar Alarm Factor tables are all equal.')
        else:
            print('There is a difference in Burglar Alarm Factor tables.')
        print()

        print('Testing Building Age Modifier tables...')
        if self.compareBuildingAgeModifier():
            print('Building Age Modifier tables are all equal.')
        else:
            print('There is a difference in Building Age Modifier tables.')
        print()

        print('Testing Building AOI tables...')
        if self.compareBuildingAOI():
            print('Building AOI tables are all equal.')
        else:
            print('There is a difference in Building AOI tables.')
        print()

        print('Testing BPP AOI tables...')
        if self.compareBPPAOI():
            print('BPP AOI tables are all equal.')
        else:
            print('There is a difference in BPP AOI tables.')
        print()

        print('Testing Equipment Breakdown Limit Relativity Modifier tables...')
        if self.compareEBLimitRelativity():
            print('Equipment Breakdown Limit Relativity Modifier tables are all equal.')
        else:
            print('There is a difference in Equipment Breakdown Limit Relativity Modifier tables.')
        print()

        print('Testing Equipment Breakdown Deductible Factor tables...')
        if self.compareEBDeductibleFactor():
            print('Equipment Breakdown Deductible Factor tables are all equal.')
        else:
            print('There is a difference in Equipment Breakdown Deductible Factor tables.')
        print()

        print('**************************************************************')
        print('END OF CROSS PROGRAM VALIDATION')
        print('**************************************************************')
        print()

        print('**************************************************************')
        print('START OF CROSS COMPANY VALIDATION')
        print('**************************************************************')
        print()

        self.compareCompanies()

        print('**************************************************************')
        print('END OF CROSS COMPANY VALIDATION')
        print('**************************************************************')