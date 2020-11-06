import sys, json
import ElentraReport as ER
import CommUtils as comms
from datetime import datetime

def createFormattedExtract(commands):
    comms.sendMessage(commands)
    Report = ER.ElentraReport(options=commands, reportType="Formatted Extract")
    Report.setExtractData(path=commands["extractDataFilePath"])
    Report.setLookupWorkbook(path=commands["lookupTableFilePath"])
    Report.setLookupTables()
    Report.createFormattedExtract()
    comms.sendMessage("Your extract has been formatted!", "progMsg")
    comms.sendMessage("100", "prog")
    del Report

def createGeneratedReport(commands):
    comms.sendMessage(commands)
    Report = ER.ElentraReport(options=commands, reportType="Full Report")
    Report.setExtractData(path=commands["extractDataFilePath"])
    Report.setLookupWorkbook(path=commands["lookupTableFilePath"])
    Report.setLookupTables()
    Report.createFormattedExtract()
    Report.createResidentAnalysis()
    comms.sendMessage("Your report has been generated!", "progMsg")
    comms.sendMessage("100", "prog")
    del Report

for line in sys.stdin:
    commands = json.loads(line)
    if commands["func"] == "createFormattedExtract": 
        createFormattedExtract(commands)
    elif commands["func"] == "createGeneratedReport":
        createGeneratedReport(commands)
