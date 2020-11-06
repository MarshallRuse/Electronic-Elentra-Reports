#!/usr/bin/env python
# coding: utf-8

import numpy as np
import pandas as pd
import os
import json
import itertools
# import warnings
import CommUtils as comms

# warnings.filterwarnings("default")
# # Format Elentra Extract 
# ### Load the extract into memory
# 
# Extract the date range and the program from the first two rows of the csv.

class ElentraReport:

    def __init__(self, data=None, LookupWorkbook=None, options=None, reportType=None):
        comms.sendMessage("ElentraReport instantiated")
        # Extract Data
        self.data = data

        # Lookup Tables
        self.LookupWorkbook = LookupWorkbook
        self.EPALookup = None
        self.IMTraineeLookup = None
        #self.SubspecTraineeLookup = None
        self.SiteLookup = None
        self.BlockLookup = None

        if options is not None:
            self.options = options
        else:
            pass
            # TODO: throw an error, options will never be None 
            #self.options = cs.loadConfigSettings()

        # Values derived from the first few (immediately deleted) lines of the extract,
        # used to name the resulting report
        self.dateRange = ""
        self.program = ""

        self.FormattedExtract = None
        self.ResidentAnalysis = None
        self.PTResidentAnalysis = None
        self.PTBlock = None
        self.PTSite = None

        self.saveLocation = options["saveReportFolderPath"] + options["pathSeparator"]

        self.reportType = reportType

        self.commsCount = 0

    def setExtractData(self, path):
        # if self.data is None:
        #     comms.sendMessage("Self.data is NONE", "progMsg")
        # else:
        #     comms.sendMessage("Self.data IS NOT NONE", "progMsg")
        comms.sendMessage("Grabbing your extract data...", "progMsg")
        self.commsCount += 1
        comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
    
        with open(path, encoding='utf-8') as f:
            for x in range(3):
                if x == 1:
                    line = f.readline()
                    self.dateRange = line[line.find(",")+1:].strip()[1:-1]
                elif x == 2:
                    line = f.readline()
                    self.program = line[line.find(",")+1:].strip()[1:-1]
                else:
                    f.readline()
        
        self.data = pd.read_csv(path, 
            error_bad_lines=False,
            skiprows=3)

        comms.sendMessage("Extract data grabbed...", "progMsg")
        self.commsCount += 1
        comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
        # comms.sendMessage("Warning Message: " + Warning)
        return

    def setLookupWorkbook(self, path=None, LookupWorkbook=None):
        comms.sendMessage("Opening up the Lookup Table workbook...", "progMsg")
        self.commsCount += 1
        comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
        if LookupWorkbook is not None:
            self.LookupWorkbook = LookupWorkbook
            return
        if path is not None and LookupWorkbook is None:
            self.LookupWorkbook = pd.read_excel(path, sheet_name=["VLOOKUP MASTER", "IM Trainees", "Site", "BLOCK"])
            return
        
    def setLookupTables(self, LookupWorkbook=None):
        comms.sendMessage("Gleaning the Lookup Table details...", "progMsg")
        self.commsCount += 1
        comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
        if LookupWorkbook is not None:
            self.EPALookup = self.LookupWorkbook["VLOOKUP MASTER"]
            self.IMTraineeLookup = self.LookupWorkbook["IM Trainees"]
            #self.SubspecTraineeLookup = self.LookupWorkbook["Trainees 2020-2021 (POWER)"]
            self.SiteLookup = self.LookupWorkbook["Site"]
            self.BlockLookup = self.LookupWorkbook["BLOCK"]
            return
        elif LookupWorkbook is None and self.LookupWorkbook is not None:
            self.EPALookup = self.LookupWorkbook["VLOOKUP MASTER"]
            self.IMTraineeLookup = self.LookupWorkbook["IM Trainees"]
            #self.SubspecTraineeLookup = self.LookupWorkbook["Trainees 2020-2021 (POWER)"]
            self.SiteLookup = self.LookupWorkbook["Site"]
            self.BlockLookup = self.LookupWorkbook["BLOCK"]
            return
        else:
            return "No Lookup Workbook has been provided, and none is set on the ElentraReport class."

    def calibrateTrainingLevel(self, tl, calibration):
        level = int(tl[3]) # 4th character of PGY#
        suffix = tl[4:] if len(tl) > 4 else "" # Account for the C suffix that sometimes accompanies
        newLevel = level + calibration
        newLevel = newLevel if newLevel > 0 else 1
        newTl = "PGY" + str(newLevel) + suffix
        return newTl
    
    def createFormattedExtract(self):
        comms.sendMessage("Formatting the extract...", "progMsg")
        self.commsCount += 1
        comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
        if self.data is not None:
            # ### Replace the default "Entrustment / Overall Category" values with their formatted ones.
            # The only values in the "Entrustment / Overall Category" column should be:
            #     1. Intervention
            #     2. Direction
            #     3. Support
            #     4. Competent
            #     5. Proficient
            self.data.replace({"Entrustment / Overall Category": {"Intervention": "1. Intervention", 
                                               "Direction": "2. Direction",
                                               "Support": "3. Support",
                                               "Autonomy": "4. Competent",
                                               "Competent": "4. Competent",
                                               "Excellence": "5. Proficient",
                                               "Proficient": "5. Proficient"}}, inplace=True)

            supportedVals = ['1. Intervention', '2. Direction', '3. Support', '4. Competent', '5. Proficient', np.nan]
            unsupportedVals = [val for val in list(self.data['Entrustment / Overall Category'].unique()) if val not in supportedVals]
            self.fixUnsupportedVals(unsupportedVals)

            # ### Remove the forms with no Date of Submission
            if self.options['Options']["removeUnsubmitted"]:
                self.data.dropna(subset=["Date of Assessment Form Submission"], inplace=True)


            # ### Create a column called "EPA Code and Name" and use the Assessment Form ID / form_id to populate.

            self.data.insert(1, "EPA Code and Name", self.data["Assessment Form Code"].map(self.EPALookup.set_index("form_id")["Stage and Name"]))


            # ### Create a column called "Site" based on the Contextual Variable Site
            self.data.insert(2, "Site", self.data["CV ID 9533 : Site"].map(self.SiteLookup.set_index("Lookup Site")["Formatted Site"]))


            # ### Create a column called "Block" based on the Date of encounter
            self.data.sort_values(by="Date of encounter", inplace=True)
            self.data["Date of encounter"] = pd.to_datetime(self.data["Date of encounter"])
            self.BlockLookup.sort_values(by="Start Date", inplace=True)
            self.data = pd.merge_asof(self.data, self.BlockLookup[["Start Date", "Year and Block"]], left_on="Date of encounter", right_on="Start Date")
            # Drop the unnecessary Start Date column, rename Year and Block
            self.data.drop("Start Date", axis=1, inplace=True)
            self.data.rename(columns={"Year and Block": "Block"}, inplace=True)


            # ### Format the Assessee and Assessors names
            self.data["Resident"] = self.data["Assessee Lastname"].apply(lambda x: x.upper()) + ", " + self.data["Assessee Firstname"]
            self.data["Assessor Fullname"] = self.data["Assessor Lastname"].apply(lambda x: x.upper()) + ", " + self.data["Assessor Firstname"]

            # Remove the forms with self-assessments or Procedure Log POSTMD as the assessor
            if self.options['Options']["removeProcedures"]:
                #self.data.dropna(subset=["Entrustment / Overall Category"], inplace=True)
                self.data.drop(self.data[(self.data['Resident'] == self.data['Assessor Fullname']) | self.data['Assessor Fullname'].isin(['POSTMD, Procedure Log'])].index, inplace=True)

            # ### Redo the Assessee Training Level column to have accurate levels
            if (self.program == "Internal Medicine"):
                self.IMTraineeLookup.dropna(axis=0, subset=["Elentra ID"], inplace=True)
                self.IMTraineeLookup.reset_index(drop=True, inplace=True)
                self.IMTraineeLookup["Elentra ID"] = self.IMTraineeLookup["Elentra ID"].astype(int).astype(str)

                self.data["Assessee User ID"] = self.data["Assessee User ID"].astype(int).astype(str)

                del self.data["Assessee Training Level"]

                self.data.insert(2, "Assessee Training Level", self.data["Assessee User ID"].map(self.IMTraineeLookup.set_index("Elentra ID")["Training Level"]))

                mask2017_18 = (self.data['Date of encounter'] >= '2017-7-1') & (self.data['Date of encounter'] < '2018-7-1')
                mask2018_19 = (self.data['Date of encounter'] >= '2018-7-1') & (self.data['Date of encounter'] < '2019-7-1')
                mask2019_20 = (self.data['Date of encounter'] >= '2019-7-1') & (self.data['Date of encounter'] < '2020-7-1')
                mask2020_21 = (self.data['Date of encounter'] >= '2020-7-1') & (self.data['Date of encounter'] < '2021-7-1')

                self.data["Academic Year"] = ""
                self.data.loc[mask2017_18, "Academic Year"] = "2017-18"
                self.data.loc[mask2018_19, "Academic Year"] = "2018-19"
                self.data.loc[mask2019_20, "Academic Year"] = "2019-20"
                self.data.loc[mask2020_21, "Academic Year"] = "2020-21"

                self.data.loc[mask2017_18, "Assessee Training Level"] = self.data["Assessee Training Level"].apply(self.calibrateTrainingLevel, args=(-3,))
                self.data.loc[mask2018_19, "Assessee Training Level"] = self.data["Assessee Training Level"].apply(self.calibrateTrainingLevel, args=(-2,))
                self.data.loc[mask2019_20, "Assessee Training Level"] = self.data["Assessee Training Level"].apply(self.calibrateTrainingLevel, args=(-1,))

                self.data["Week of Year"] = 0
                yearStart2017 = pd.Timestamp('2017-7-1')
                self.data.loc[mask2017_18, "Week of Year"] = (self.data["Date of encounter"] - yearStart2017).dt.days // 7
                yearStart2018 = pd.Timestamp('2018-7-1')
                self.data.loc[mask2018_19, "Week of Year"] = (self.data["Date of encounter"] - yearStart2018).dt.days // 7
                yearStart2019 = pd.Timestamp('2019-7-1')
                self.data.loc[mask2019_20, "Week of Year"] = (self.data["Date of encounter"] - yearStart2019).dt.days // 7
                yearStart2020 = pd.Timestamp('2020-7-1')
                self.data.loc[mask2020_21, "Week of Year"] = (self.data["Date of encounter"] - yearStart2020).dt.days // 7


            # ### Set Assessment ID as the index for the table
            #self.data.set_index("Assessment ID", inplace=True)


            # ### Drop columns with no values to reduce size of final table
            if self.options['Options']["removeEmptyColumns"]:
                comms.sendMessage("Removing the columns full of nothing...", "progMsg")
                self.commsCount += 1
                comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
                self.data.dropna(axis=1, how='all', inplace=True)


            if self.reportType == "Formatted Extract" or self.options['Options']['createSpinoffExtract'] or self.options['Options']["includeExtractDataInReport"]:
                self.FormattedExtract = self.data.copy()

            # ### Save a copy of the Formatted Extract for seperate export
            if self.options['Options']["createSpinoffExtract"]:
                comms.sendMessage("Spinning off a formatted extract...", "progMsg")
                self.commsCount += 1
                comms.sendMessage(self.progressUpdate(self.commsCount), "prog")

            # Use the xlsxwriter engine to format the xlsx output.  Format the sheet as a table
            if self.options['Options']["createSpinoffExtract"] and self.options['Options']["spinoffExtractAsTable"]:
                comms.sendMessage("Making a pretty table...", "progMsg")
                self.commsCount += 1
                comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
                
                finalExtractFileName = self.fileExists(self.saveLocation + self.program + " " + self.dateRange + " - FormattedExtract", "xlsx")
                FEWriter = pd.ExcelWriter(finalExtractFileName, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
                self.FormattedExtract.to_excel(FEWriter, sheet_name="DataExtract", index=False, startrow=1, header=False)
                FEWorkbook = FEWriter.book
                FEWorksheet = FEWriter.sheets['DataExtract']
                (maxRow, maxCol) = self.FormattedExtract.shape
                columnSettings = [{'header': column} for column in self.FormattedExtract.columns]
                FEWorksheet.add_table(0, 0, maxRow, maxCol - 1, {
                    'columns': columnSettings,
                    'name': 'ExtractData'
                })
                FEWorksheet.set_column(0, maxCol - 1, 30) # set a more comfortable column width
                FEWriter.save()
            elif self.options['Options']["createSpinoffExtract"] and not self.options['Options']["spinoffExtractAsTable"]:
                comms.sendMessage("Giving you a plain ole boring formatted extract...", "progMsg")
                finalExtractFileName = self.fileExists(self.saveLocation + self.program + " " + self.dateRange + " - FormattedExtract", "xlsx")
                comms.sendMessage(finalExtractFileName)
                #FEWriter = pd.ExcelWriter(finalExtractFileName) # pylint: disable=abstract-class-instantiated
                self.FormattedExtract.to_excel(finalExtractFileName, sheet_name="DataExtract", index=False)

# # Create a Resident Analysis Table
    def createResidentAnalysis(self):
        comms.sendMessage("Generating a report...", "progMsg")
        self.commsCount += 1
        comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
        if self.data is not None:

            # This resource is very helpful for slicing multi-index dataframes: https://stackoverflow.com/questions/53927460/select-rows-in-pandas-multiindex-dataframe
            self.ResidentAnalysis = self.data[['Resident', 'EPA Code and Name', 'Assessment Form Code']].drop_duplicates(subset=['Resident', 'EPA Code and Name'])

            self.ResidentAnalysis.sort_values(by=['Resident', 'EPA Code and Name'], inplace=True)
            self.ResidentAnalysis.reset_index(drop=True, inplace=True)


            ResEntrustmentLevels = self.data.groupby(['Resident', 'EPA Code and Name', 'Entrustment / Overall Category']).size().unstack(fill_value=0)

            ResEntrustmentLevels.reset_index(level=['Resident', 'EPA Code and Name'], inplace=True)

            self.ResidentAnalysis = pd.merge(self.ResidentAnalysis, ResEntrustmentLevels, how="left", on=['Resident', 'EPA Code and Name'])

            # Make sure that all of the entrustment levels are there
            for ind, el in enumerate(['1. Intervention', '2. Direction', '3. Support', '4. Competent', '5. Proficient'], 2):
                if not el in self.ResidentAnalysis:
                    self.ResidentAnalysis.insert(ind, el, 0)
            #Make sure the columns are in the correct order
            self.ResidentAnalysis = self.ResidentAnalysis[['Resident', 'EPA Code and Name', 'Assessment Form Code', '1. Intervention', '2. Direction', '3. Support', '4. Competent', '5. Proficient']]

            self.ResidentAnalysis["Number of Entrustments"] = self.ResidentAnalysis["4. Competent"] + self.ResidentAnalysis["5. Proficient"]


            # ### Add the "Target Entrustments" Column
            comms.sendMessage("Setting the target entrustments...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            self.ResidentAnalysis.insert(len(self.ResidentAnalysis.columns), "Target Entrustments", self.ResidentAnalysis["Assessment Form Code"].map(self.EPALookup.set_index("form_id")["Target EPA"]))


            # ### Convert String Number Columns to Numeric Columns
            self.ResidentAnalysis['Target Entrustments'].replace('2 (in each procedure)', 2, inplace=True)

            self.ResidentAnalysis[['1. Intervention', '2. Direction', '3. Support', '4. Competent', '5. Proficient', 'Number of Entrustments','Target Entrustments']] = self.ResidentAnalysis[['1. Intervention', '2. Direction', '3. Support', '4. Competent', '5. Proficient', 'Number of Entrustments','Target Entrustments']].apply(pd.to_numeric)


            # ### Add the "Entrustments Acquired" Column
            comms.sendMessage("Checking whose acquired the entrustments...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            self.ResidentAnalysis['Entrustments Acquired'] = np.where(pd.to_numeric(self.ResidentAnalysis['Number of Entrustments']) >= pd.to_numeric(self.ResidentAnalysis['Target Entrustments']), 'Yes', 'No')

            # ### Create a Unique Assessors column
            comms.sendMessage("How many assessors are there?...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            NumUniqueAssessors = self.data[['Resident', 'EPA Code and Name', 'Assessor Fullname']].groupby(['Resident', 'EPA Code and Name']).agg({'Assessor Fullname': 'nunique'})
            self.ResidentAnalysis = pd.merge(self.ResidentAnalysis, NumUniqueAssessors, how="left", on=['Resident', 'EPA Code and Name'])
            self.ResidentAnalysis.rename(columns={'Assessor Fullname': 'Number of Assessors (Unique Count)'}, inplace=True)

            # ### Create a Direct Entrustments column
            comms.sendMessage("Counting Direct Entrustments...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            DirectEntrustmentsMask = self.data[(self.data['CV ID 9524 : Type of Assessment'] == 'Direct Observation') & ((self.data['Entrustment / Overall Category'] == '4. Competent') | (self.data['Entrustment / Overall Category'] == '5. Proficient'))]
            NumDirectEntrustments = DirectEntrustmentsMask[['Resident', 'EPA Code and Name', 'CV ID 9524 : Type of Assessment']].groupby(['Resident', 'EPA Code and Name']).agg({'CV ID 9524 : Type of Assessment': 'count'})
            self.ResidentAnalysis = pd.merge(self.ResidentAnalysis, NumDirectEntrustments, how="left", on=['Resident', 'EPA Code and Name'])
            self.ResidentAnalysis.rename(columns={'CV ID 9524 : Type of Assessment': 'Direct Entrustments'}, inplace=True)
            self.ResidentAnalysis['Direct Entrustments'] = self.ResidentAnalysis['Direct Entrustments'].fillna(0).astype('int64')


            # ### Create a Comments Table
            comms.sendMessage("Sprinkling in some comments...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            CommentsLookup = self.data[['Resident', 'EPA Code and Name', '2 - 3 Strengths', '2 - 3 Actions or areas for improvement']]

            CommentsLookup['2 - 3 Strengths'] = CommentsLookup['2 - 3 Strengths'].apply(lambda x: str(x))
            CommentsLookup['2 - 3 Actions or areas for improvement'] = CommentsLookup['2 - 3 Actions or areas for improvement'].apply(lambda x: str(x))

            StrengthsLookup = CommentsLookup.groupby(['Resident', 'EPA Code and Name'])['2 - 3 Strengths'].apply(','.join).reset_index()
            WeaknessesLookup = CommentsLookup.groupby(['Resident', 'EPA Code and Name'])['2 - 3 Actions or areas for improvement'].apply(','.join).reset_index()

            ResidentAnalysisComments = pd.merge(StrengthsLookup, WeaknessesLookup, how="outer", on=["Resident", "EPA Code and Name"]) 

            self.ResidentAnalysis = pd.merge(self.ResidentAnalysis, ResidentAnalysisComments, how="left", on=['Resident', 'EPA Code and Name'])

            self.ResidentAnalysis.rename(columns={'2 - 3 Strengths': 'Strengths', '2 - 3 Actions or areas for improvement': 'Weaknesses'}, inplace=True)

            PivotTable = pd.concat([self.ResidentAnalysis.assign(**{x: '' for x in ['Resident', 'EPA Code and Name'][i:]}).groupby(['Resident', 'EPA Code and Name']).sum() for i in range(1, 3)]).sort_index()

            self.PTResidentAnalysis = pd.merge(PivotTable, ResidentAnalysisComments, how="left", on=['Resident', 'EPA Code and Name'])

            self.PTResidentAnalysis.rename(columns={'2 - 3 Strengths': 'Strengths', '2 - 3 Actions or areas for improvement': 'Weaknesses'}, inplace=True)


            # Now that the Pivot Tables been merged, its multi-level index has collapsed.  This will be corrected for the final Excel workbook appearance artificially with merging of Excel cells, but in the meantime it allows us to conveniently add the "Entrustments Acquired" column to the PivotTable, which got dropped by the aggregating sum function during PivotTable's assignment.
            EAColumn = np.where(self.PTResidentAnalysis['EPA Code and Name'] != '', np.where(pd.to_numeric(self.PTResidentAnalysis['Number of Entrustments']) >= pd.to_numeric(self.PTResidentAnalysis['Target Entrustments']), 'Yes', 'No'), '')
            self.PTResidentAnalysis.insert(10, "Entrustments Acquired", EAColumn)


            # Drop the Assessment Form Code column from the ResidentAnalysis sheets
            self.ResidentAnalysis.drop('Assessment Form Code', axis=1, inplace=True)
            self.PTResidentAnalysis.drop('Assessment Form Code', axis=1, inplace=True)


            # # Create Block Pivot Table
            comms.sendMessage("Pivoting on some blocks...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            self.PTBlock = pd.pivot_table(self.data, index=['Resident', 'Block'], columns='Entrustment / Overall Category', values='Assessment ID', aggfunc=lambda x: len(x.unique())).fillna(0)


            # Add the subtotal rows
            self.PTBlock = pd.concat([self.PTBlock.reset_index(level=['Resident', 'Block']).assign(**{x: '' for x in ['Resident', 'Block'][i:]}).groupby(['Resident', 'Block']).sum() for i in range(1, 3)]).sort_index()

            # Reset the self.PTBlock indices to collapese the multi-index.  Re-merge cells below.
            self.PTBlock.reset_index(level=['Resident', 'Block'], inplace=True)
            #self.PTBlock = self.PTBlock[self.PTBlock['Resident'] != 'All']


            # # Create Site Pivot Table
            comms.sendMessage("Pivoting on some sites...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            self.PTSite = pd.pivot_table(self.data, index=['Resident', 'Site'], columns='Entrustment / Overall Category', values='Assessment ID', aggfunc=lambda x: len(x.unique())).fillna(0)


            # Add the subtotal rows
            self.PTSite = pd.concat([self.PTSite.reset_index(level=['Resident', 'Site']).assign(**{x: '' for x in ['Resident', 'Site'][i:]}).groupby(['Resident', 'Site']).sum() for i in range(1, 3)]).sort_index()


            # Reset the self.PTSite indices to collapese the multi-index.  Re-merge cells below.
            self.PTSite.reset_index(level=['Resident', 'Site'], inplace=True)
            #self.PTSite = self.PTSite[self.PTSite['Resident'] != 'All']

            # # Write the tables to an Excel Workbook
            comms.sendMessage("Writing everything down...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            finalReportFileName = self.fileExists(self.saveLocation + self.program + " " + self.dateRange, "xlsx")
            writer = pd.ExcelWriter(finalReportFileName) # pylint: disable=abstract-class-instantiated
                
            self.ResidentAnalysis.to_excel(writer, sheet_name="ResidentAnalysis", index=False, startrow=1, header=False)
            self.PTResidentAnalysis.to_excel(writer, sheet_name="ResidentAnalysisPT", index=False, startrow=1, header=False)
            self.PTBlock.to_excel(writer, sheet_name="BlockAnalysis", index=False, startrow=1, header=False)
            self.PTSite.to_excel(writer, sheet_name='SiteAnalysis', index=False, startrow=1, header=False)
            if self.options['Options']["includeExtractDataInReport"]:
                self.FormattedExtract.to_excel(writer, sheet_name="DataExtract", index=False)

            workbook = writer.book
            comms.sendMessage("Making it look pretty...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            # Header Row Format
            FormatHeader = workbook.add_format({
                'font_color': '#FFFFFF',
                'bold': True,
                'bg_color': '#2F75B5',
                'align': 'center'
            })

            # Alignment Formats
            FormatCenteredValues = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter'
            })
            FormatLeftHCenteredV = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter',
                'text_wrap': True
            })

            # Comments Formats
            FormatComments = workbook.add_format({
                'align': 'left',
                'valign': 'top',
                'text_wrap': True,
                'font_size': 10
            })

            # EPAs by CBD stage Formats
            FormatEPATTD = workbook.add_format({
                'bg_color': '#00B0F0'
            })
            FormatEPAFOD = workbook.add_format({
                'bg_color': '#92D050'
            })
            FormatEPACOD = workbook.add_format({
                'bg_color': '#FFC000'
            })

            # Bad Cells Formats
            FormatExcelBad = workbook.add_format({
                'bg_color':   '#FFC7CE',
                'font_color': '#9C0006'
            })

            FormatMergedResidents = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#9BC2E6'})


            #### Format the ResidentAnalysis sheet ####
            ResAnSheet = writer.sheets['ResidentAnalysis']
            (maxRow, maxCol) = self.ResidentAnalysis.shape

            # Format the header row
            # Write the column headers with the defined format.
            for col_num, value in enumerate(self.ResidentAnalysis.columns.values):
                ResAnSheet.write(0, col_num, value, FormatHeader)

            # Resize the columns and apply formats
            ResAnSheet.set_column(0,0,30, FormatLeftHCenteredV) # Set the Resident column to a width of 30
            ResAnSheet.set_column(1,1,60, FormatLeftHCenteredV) # Set the EPA Code and Name column to a width of 60
            ResAnSheet.set_column(2, 6, 13, FormatCenteredValues) # Set the Entrustment Level columns to a width of 13
            ResAnSheet.set_column(7, 11, 21, FormatCenteredValues) # Set the Entrustment assessment columns to a width of 21
            ResAnSheet.set_column(12, 13, 80, FormatComments) # Set the comments columns to a width of 80


            ResAnSheet.conditional_format(1, 1, maxRow, 1, {
                'type': 'text',
                'criteria': 'containing',
                'value': '1. TTD',
                'format': FormatEPATTD
            })
            ResAnSheet.conditional_format(1, 1, maxRow, 1, {
                'type': 'text',
                'criteria': 'containing',
                'value': '2. FOD',
                'format': FormatEPAFOD
            })
            ResAnSheet.conditional_format(1, 1, maxRow, 1, {
                'type': 'text',
                'criteria': 'containing',
                'value': '3. COD',
                'format': FormatEPACOD
            })

            ResAnSheet.conditional_format(1, 9, maxRow, 9, {
                'type': 'text',
                'criteria': 'containing',
                'value': 'No',
                'format': FormatExcelBad
            })

            #### Format the ResidentAnalysisPT sheet ####
            ResAnPTSheet = writer.sheets['ResidentAnalysisPT']
            (maxRow, maxCol) = self.PTResidentAnalysis.shape

            # Format the header row
            # Write the column headers with the defined format.
            for col_num, value in enumerate(self.PTResidentAnalysis.columns.values):
                ResAnPTSheet.write(0, col_num, value, FormatHeader)
                

    
            # Merge the Resident Column cells based on name to mimic Pivot Table appearance
            for resident in self.PTResidentAnalysis['Resident'].unique():
                firstIndex = self.PTResidentAnalysis[self.PTResidentAnalysis['Resident'] == resident].index[0]
                lastIndex = self.PTResidentAnalysis[self.PTResidentAnalysis['Resident'] == resident].index[-1]
                ResAnPTSheet.merge_range(firstIndex + 1, 0, lastIndex + 1, 0, resident, FormatMergedResidents)
                for col in range(1, maxCol):
                    cellVal = self.PTResidentAnalysis[self.PTResidentAnalysis.columns[col]].iloc[firstIndex]
                    if pd.isnull(cellVal):
                        cellVal = ""
                    ResAnPTSheet.write(firstIndex + 1, col, cellVal, FormatMergedResidents)

            # Resize the columns and apply formats
            ResAnPTSheet.set_column(0,0,30, FormatLeftHCenteredV) # Set the Resident column to a width of 30
            ResAnPTSheet.set_column(1,1,60, FormatLeftHCenteredV) # Set the EPA Code and Name column to a width of 60
            ResAnPTSheet.set_column(2, 6, 13, FormatCenteredValues) # Set the Entrustment Level columns to a width of 13
            ResAnPTSheet.set_column(7, 11, 21, FormatCenteredValues) # Set the Entrustment assessment columns to a width of 21
            ResAnPTSheet.set_column(12, 13, 80, FormatComments) # Set the comments columns to a width of 80


            ResAnPTSheet.conditional_format(1, 1, maxRow, 1, {
                'type': 'text',
                'criteria': 'containing',
                'value': '1. TTD',
                'format': FormatEPATTD
            })
            ResAnPTSheet.conditional_format(1, 1, maxRow, 1, {
                'type': 'text',
                'criteria': 'containing',
                'value': '2. FOD',
                'format': FormatEPAFOD
            })
            ResAnPTSheet.conditional_format(1, 1, maxRow, 1, {
                'type': 'text',
                'criteria': 'containing',
                'value': '3. COD',
                'format': FormatEPACOD
            })

            ResAnPTSheet.conditional_format(1, 9, maxRow, 9, {
                'type': 'text',
                'criteria': 'containing',
                'value': 'No',
                'format': FormatExcelBad
            })


            #### Format the Block sheet ####

            BlockSheet = writer.sheets['BlockAnalysis']
            (maxRow, maxCol) = self.PTBlock.shape

            # Format the header row
            # Write the column headers with the defined format.
            for col_num, value in enumerate(self.PTBlock.columns.values):
                BlockSheet.write(0, col_num, value, FormatHeader)
                
            # Merge the Resident Column cells based on name to mimic Pivot Table appearance
            for resident in self.PTBlock['Resident'].unique():
                firstIndex = self.PTBlock[self.PTBlock['Resident'] == resident].index[0]
                lastIndex = self.PTBlock[self.PTBlock['Resident'] == resident].index[-1]
                BlockSheet.merge_range(firstIndex + 1, 0, lastIndex + 1, 0, resident, FormatMergedResidents)
                for col in range(1, maxCol):
                    cellVal = self.PTBlock[self.PTBlock.columns[col]].iloc[firstIndex]
                    if pd.isnull(cellVal):
                        cellVal = ""
                    BlockSheet.write(firstIndex + 1, col, cellVal, FormatMergedResidents)

            # Resize the columns and apply formats
            BlockSheet.set_column(0,0,30, FormatLeftHCenteredV)
            BlockSheet.set_column(1,1,60, FormatLeftHCenteredV) # Set the EPA Code and Name column to a width of 60
            BlockSheet.set_column(2, 7, 13, FormatCenteredValues) # Set the Entrustment Level columns to a width of 13


            #### Format the Site sheet ####

            SiteSheet = writer.sheets['SiteAnalysis']
            (maxRow, maxCol) = self.PTSite.shape

            # Format the header row
            # Write the column headers with the defined format.
            for col_num, value in enumerate(self.PTSite.columns.values):
                SiteSheet.write(0, col_num, value, FormatHeader)
    
            # Merge the Resident Column cells based on name to mimic Pivot Table appearance
            for resident in self.PTSite['Resident'].unique():
                firstIndex = self.PTSite[self.PTSite['Resident'] == resident].index[0]
                lastIndex = self.PTSite[self.PTSite['Resident'] == resident].index[-1]
                SiteSheet.merge_range(firstIndex + 1, 0, lastIndex + 1, 0, resident, FormatMergedResidents)
                for col in range(1, maxCol):
                    cellVal = self.PTSite[self.PTSite.columns[col]].iloc[firstIndex]
                    if pd.isnull(cellVal):
                        cellVal = ""
                    SiteSheet.write(firstIndex + 1, col, cellVal, FormatMergedResidents)

            # Resize the columns and apply formats
            SiteSheet.set_column(0,0,30, FormatLeftHCenteredV)
            SiteSheet.set_column(1,1,60, FormatLeftHCenteredV) # Set the EPA Code and Name column to a width of 60
            SiteSheet.set_column(2, 7, 13, FormatCenteredValues) # Set the Entrustment Level columns to a width of 13

            comms.sendMessage("Saving our hard work...", "progMsg")
            self.commsCount += 1
            comms.sendMessage(self.progressUpdate(self.commsCount), "prog")
            writer.save()
    
    def progressUpdate(self, fraction):
        if self.reportType == "Formatted Extract":
            return str(round((fraction / 11) * 100))
        elif self.reportType == "Full Report":
            return str(round((fraction / 20) * 100))

    def fileExists(self, originalFile, ext):
        actualname = "%s.%s" % (originalFile, ext)
        c = itertools.count(start=1)
        while os.path.exists(actualname):
            actualname = "%s (%d).%s" % (originalFile, next(c), ext)
        return actualname

    def fixUnsupportedVals(self, unsupportedVals):
        if len(unsupportedVals) > 0:
            comms.sendMessage({"unsupportedVals": unsupportedVals}, "requireInput")
            responseDict = json.loads(input())
            for key, val in responseDict["selectedChoices"].items():
                self.data.replace({"Entrustment / Overall Category": {key: val}}, inplace=True)
