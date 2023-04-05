from bs4 import BeautifulSoup
import re
import pandas as pd;
import openpyxl as pyxl;
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path

df = pd.DataFrame(columns=["MACHINE","NC PROGRAM PATH", "MATERIAL ID", "BLANK", "WEIGHT", "MACHINING TIME", "LASER TOTAL CUTTING LENGTH", "SCRAP"])
setupPlanList = ["9500.html", "81011.html", "319930.html", "500585.html", "3712951.html", "test.html"]
setupPlanDirectory = Path("testing/html")
excelFileDirectory = Path("testing/xlsx")


def readFiles(setupPlanDirectory, setupPlanList):
    for file in setupPlanList:
        setupPlanPath = Path(setupPlanDirectory / file)
        if not setupPlanPath.is_file():
            print("Napaka datoteka", file, "ne obstaja")
            return
        with open(setupPlanPath) as fp:
            setupPlan = BeautifulSoup(fp, "html.parser")
        dataframeAppendFile(setupPlan)

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
    workbook = pyxl.load_workbook(excelFileDirectory / "main.xlsx")
    worksheet = workbook["podatki"];
    for row in dataframe_to_rows(df, header = True, index = False):
        worksheet.append(row)
    workbook.save(excelFileDirectory / "main_output.xlsx")
    workbook.close()

readFiles(setupPlanDirectory, setupPlanList)
writeToXlsx(excelFileDirectory)
print(df)