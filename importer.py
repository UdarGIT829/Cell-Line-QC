
import xlrd
from xlsxwriter.workbook import Workbook
import csv
import os
import math
import re
import sys

from pathlib import Path

def importCellCountXLSX():
    directory = os.path.dirname(__file__)

    numberOfMasterFiles,masterFiles = 0, list() #Should only be 1!!!
    masterFile = str()
    for path in Path(directory).rglob('*.xlsx'):
        numberOfMasterFiles += 1
        masterFiles.append(path)
        if numberOfMasterFiles > 1:
            print("Too many master files, there can only be 1 master file")
            print("Master files seen: ", masterFiles)
            sys.exit()
        tempNameList = list()
        tempIdentifier = os.path.basename(path).replace(" ","")
        with xlrd.open_workbook(path) as workbook:
            nameOverride = ""
            names = workbook.sheet_names()
            sheetAmount = len(names)
            for index in range(sheetAmount):
                print("Reading sheet #: ", index+1)
                sheet = workbook.sheet_by_index(index)
                tempSheetName = tempIdentifier + "sheet" + str(index+1) + ".csv"
                tempNameList.append(tempSheetName)
                with open(tempSheetName, 'a+', newline="") as file:
                    col = csv.writer(file)
                    for row in range(sheet.nrows):
                        col.writerow(sheet.row_values(row))

    ###Begin importing from Cell count csv's
    cellCountFiles = {}
    pattern_NewTemplate = re.compile("\s-\s([a-zA-Z0-9]*)_")
    pattern_other = re.compile("\d\D([A-Z0-9]*)_")

    for path in Path(directory).rglob('*.csv'):
        if 'Master' in str(path):
            pass
        else:
            finding1 = pattern_NewTemplate.search(str(path))
            finding2 = pattern_other.search(str(path))
            if finding1:
                dayBC = finding1.group(0)[3:-2]
            elif finding2:
                dayBC = finding2.group(0)[2:-2]
            else:
                print(path)
            cellCountFiles[dayBC] = path

    cellCountData = {}
    for key in cellCountFiles:
        with open(cellCountFiles[key], 'r') as file:
            rows = csv.reader(file)
            cellCountData[key] = rows
    ###End importing from Cell Count csv's
    ###Format of cellCountData is {DayBarcode : iterable of all rows}    

    #Import from master sheet to list called masterSheetRows
    masterSheetRows = list()
    with open(tempSheetName,'r') as file:
        rows = csv.reader(file)
        for row in rows:
            masterSheetRows.append(row)
    
    #Clean all temporary files
    for toRemove in tempNameList:
        if os.path.exists(toRemove):
            os.remove(toRemove)
            #pass

    return masterSheetRows, cellCountData

def parseMasterData(sheetRows):
    masterData = {}
    rowNum= 0
    CL_IDIndex = 3
    Day1BCIndex = 8
    ColumnNumberIndex = 9
    Day7BCIndex = 10
    for row in sheetRows:
        if rowNum > 0:
            CL_ID = row[CL_IDIndex]
            Day1BC = row[Day1BCIndex]
            ColumnNumber = int(float(row[ColumnNumberIndex]))
            Day7BC = row[Day7BCIndex]
            dayBCs = [Day1BC, Day7BC]
            for dayBC in dayBCs:
                if dayBC not in masterData:
                    masterData[ dayBC ] = { ColumnNumber : CL_ID }
                else: 
                    masterData[ dayBC ][ ColumnNumber ] = CL_ID
        rowNum += 1
    return masterData

masterSheetRows, cellCountData = importCellCountXLSX()
masterData = parseMasterData(masterSheetRows)
print(masterData)
