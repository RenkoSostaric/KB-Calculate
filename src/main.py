from bs4 import BeautifulSoup
import re
import pandas as pd;
import openpyxl as pyxl;
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path

df = pd.DataFrame(columns=["MACHINE","NC PROGRAM PATH", "MATERIAL", "BLANK", "WEIGHT", "MACHINING TIME", "TOTAL CUTTING LENGTH", "SCRAP"])
setupPlanList = ["1560-P00-594_SAP 320417_XK.HTML"]
setupPlanDirectory = Path("testing/program")
excelFileDirectory = Path("testing/xlsx")
excelSourceName = "main.xlsx"
excelOutputName = "main_output.xlsx"
excelWorkbookName = "podatki"


def readFiles(setupPlanDirectory, setupPlanList):
    for file in setupPlanList:
        setupPlanPath = Path(setupPlanDirectory / file)
        if not setupPlanPath.is_file():
            print("Napaka datoteka", file, "ne obstaja")
            return
        with open(setupPlanPath) as fp:
            setupPlan = BeautifulSoup(fp, "html.parser")
        getSinglePartTable(setupPlan)
        # dataframeAppendFile(setupPlan)

def getSinglePartTable(setupPlan):
    singlePartTable = setupPlan.find(string=re.compile("INFORMATION ON SINGLE PART")).find_parent("table")
    startingImage = singlePartTable.find_next("img").find_parent("tr")
    endingImage = startingImage.find_next_sibling("tr").find_next("img").find_parent("tr")
    currentTableRow = startingImage
    singlePartData = BeautifulSoup("<table></table>", "html.parser")
    while currentTableRow != endingImage:
        if currentTableRow is not None:
            singlePartData.table.append(currentTableRow.copy())
        currentTableRow = currentTableRow.find_next_sibling("tr")
    print(singlePartData)

def dataframeAppendFile(setupPlan):
    global df
    tempRow=[] 
    for column in df:
        searchElement = setupPlan.find("td", string=re.compile(column))
        foundElement = searchElement.find_next_sibling("td")
        innerText = foundElement.text
        formattedText = " ".join(re.findall(r"\S+(?<!\s)", innerText))
        tempRow.append(formattedText)
    dfTemp = pd.DataFrame([tempRow])
    dfTemp.columns = df.columns
    df = pd.concat([df, dfTemp], ignore_index=True)

def writeToXlsx(excelFileDirectory):
    workbook = pyxl.load_workbook(excelFileDirectory / excelSourceName)
    worksheet = workbook[excelWorkbookName];
    for row in dataframe_to_rows(df, header = True, index = False):
        worksheet.append(row)
    workbook.save(excelFileDirectory / excelOutputName)
    workbook.close()

readFiles(setupPlanDirectory, setupPlanList)
# writeToXlsx(excelFileDirectory)
# print(df)