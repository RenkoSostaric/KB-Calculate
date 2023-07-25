from bs4 import BeautifulSoup
from bs4 import Comment
from lxml import etree
import re
import pandas as pd;
import openpyxl as pyxl;
import sqlite3;
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

setupPlanCollumns = ["Count", "Processed", "ProgramNumber", "ProgramRuns", "PartNumber", "DrawingNumber", "DrawingName", "Customer", "NumberOfParts", "DimensionX", "DimensionY", "Area", "Weight", "GeoFilename", "BmpFilename", "JobName", "MaterialId", "SheetSizeZ"]
setupPlanList = ["1560-P00-594_SAP 320417_XK.HTML"]
setupPlanDirectory = Path("testing/program")
excelFileDirectory = Path("testing/xlsx")
excelSourceName = "main.xlsx"
excelOutputName = "main_output.xlsx"
excelWorkbookName = "podatki"


def readFiles(setupPlanDirectory, setupPlanList):
    for file in setupPlanList:
        try:
            setupPlanPath = Path(setupPlanDirectory / file)
            if not setupPlanPath.is_file():
                print("Error: File", file, "does not exist")
                return
            with open(setupPlanPath) as fp:
                setupPlan = BeautifulSoup(fp, "html.parser")
            singlePartData = getSinglePartTable(setupPlan)
            singlePartData = singlePartData.replace("CREATE TABLE LabelPartData;\nALTER TABLE LabelPartData ADD COLUMN Count COUNTER PRIMARY KEY;", "CREATE TABLE LabelPartData (Count INTEGER PRIMARY KEY AUTOINCREMENT);")
            if(singlePartData == ""):
                print("Error: File", file, "is not valid")
                return
            dataframeAppendFile(singlePartData)
        except Exception as e:
            print("Error:", e)
            return


def getSinglePartTable(singlePartData):
    try:
        comments = singlePartData.findAll(string=lambda string:isinstance(string, Comment))
        for comment in comments:        
            commentSoup = BeautifulSoup(comment, "lxml")    
            sqlRow = commentSoup.find("sql", string="CREATE TABLE LabelPartData")
            if sqlRow is not None:
                singlePartData = sqlRow.text + ";\n"
                while sqlRow.next_sibling is not None:
                    sqlRow = sqlRow.next_sibling
                    if hasattr(sqlRow, 'name') and sqlRow.name == "sql": # type: ignore
                        singlePartData += sqlRow.text + ";\n"
                return singlePartData
        return ""
    except Exception as e:
        print("Error:", e)
        exit(1)

def dataframeAppendFile(setupPlan):
    global df
    try:
        database = sqlite3.connect('src/cache.db')
        cursor = database.cursor()
        cursor.execute("DROP TABLE IF EXISTS LabelPartData")    
        cursor.executescript(setupPlan)
        database.commit()
        sql_query = pd.read_sql_query ('''
            SELECT
            *
            FROM LabelPartData
        ''', database)
        df = pd.DataFrame(sql_query)
    except Exception as e:
        print("Error:", e)
        exit(1)

def writeToXlsx():
    try:
        workbook = pyxl.load_workbook(excelFileDirectory / excelSourceName)
        worksheet = workbook[excelWorkbookName];
        for row in dataframe_to_rows(df, header = True, index = False):
            worksheet.append(row)
        workbook.save(excelFileDirectory / excelOutputName)
        workbook.close()
    except Exception as e:
        print("Error:", e)
        exit(1)

readFiles(setupPlanDirectory, setupPlanList)
writeToXlsx()
print(df)