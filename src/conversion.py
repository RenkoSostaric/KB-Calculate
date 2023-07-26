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

class exception(Exception):
    pass

setupPlanCollumns = ["GeoFilename", "NumberOfParts", "DimensionX", "DimensionY", "Area", "Weight", "MachiningTime"]
setupPlanList = ["C:/Users/Uporabnik/Documents/KB Calculate/testing/program/1560-P00-594_SAP 320417_XK.HTML", "C:/Users/Uporabnik/Documents/KB Calculate/testing/program/1560-P00-594_SAP 320417_XX.HTML"]
excelFileDirectory = Path("testing/xlsx")
excelSourceName = "main.xlsx"
excelOutputName = "main_output.xlsx"
excelWorkbookName = "podatki"

def filesToDataframe(setupPlanList, df):
    database = sqlite3.connect('src/cache.db')
    cursor = database.cursor()
    cursor.execute("DROP TABLE IF EXISTS LabelPartData")    
    try:
        for file in setupPlanList:
            file = Path(file)
            if not file.is_file():
                raise exception("Error: File", file, "does not exist")
            if not (file.suffix == ".html" or file.suffix == ".HTML"):
                raise exception("Error: File", file, "is not a html file")
            with open(file) as fp:
                setupPlan = BeautifulSoup(fp, "html.parser")
            singlePartData = getSinglePartSQL(setupPlan)
            singlePartData = singlePartData.replace("CREATE TABLE LabelPartData;\nALTER TABLE LabelPartData ADD COLUMN Count COUNTER PRIMARY KEY;", "CREATE TABLE LabelPartData (Count INTEGER PRIMARY KEY AUTOINCREMENT);")
            singlePartData += getSinglePartMachiningTime(setupPlan)
            if(singlePartData == ""):
                raise exception("Error: File", file, "is not valid")
            df = dataframeAppendFile(singlePartData, df)
        return df
    except Exception as e:
        print("Error:", e)
        return

def getSinglePartMachiningTime(setupPlan):
    singlePartSQL = "ALTER TABLE LabelPartData ADD COLUMN MachiningTime VARCHAR(255);\n"
    singlePartTable = setupPlan.find(string=re.compile("INFORMATION ON SINGLE PART")).find_parent("table")
    machiningTimesRows = singlePartTable.find_all("font", string=re.compile("MACHINING TIME:"))
    for i in range(len(machiningTimesRows)):
        machiningTime = machiningTimesRows[i].find_parent("tr").find_all("td")[1].text
        if(machiningTime != ""):
            singlePartSQL += "UPDATE LabelPartData SET MachiningTime='" + re.findall(r"\d+\.\d+ min", machiningTime)[0] + "' WHERE COUNT = " + str(i+1) + ";"
    return singlePartSQL
    

def getSinglePartSQL(setupPlan):
    try:
        comments = setupPlan.findAll(string=lambda string:isinstance(string, Comment))
        for comment in comments:        
            commentSoup = BeautifulSoup(comment, "lxml")    
            sqlRow = commentSoup.find("sql", string="CREATE TABLE LabelPartData")
            if sqlRow is not None:
                setupPlan = sqlRow.text + ";\n"
                while sqlRow.next_sibling is not None:
                    sqlRow = sqlRow.next_sibling
                    if hasattr(sqlRow, 'name') and sqlRow.name == "sql": # type: ignore
                        setupPlan += sqlRow.text + ";\n"
                return setupPlan
        return ""
    except Exception as e:
        print("Error:", e)
        exit(1)

def dataframeAppendFile(setupPlan, df):
    try:
        database = sqlite3.connect('src/cache.db')
        cursor = database.cursor()
        cursor.execute("DROP TABLE IF EXISTS LabelPartData")    
        cursor.executescript(setupPlan)
        database.commit()
        sql_query = pd.read_sql_query ('''
            SELECT
            '''+ ", ".join(setupPlanCollumns) + '''
            FROM LabelPartData
        ''', database)
        if df.empty:
            df = pd.DataFrame(sql_query)
        else:
            df = pd.concat([df, pd.DataFrame(sql_query)], ignore_index=True)
        df["GeoFilename"] = df["GeoFilename"].replace(to_replace=r".*/(.+\.GEO)", value=r"\g<1>", regex=True)
        return df

    except Exception as e:
        print("Error:", e)
        exit(1)

def writeToXlsx(df):
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

def mainConversion(setupPlanList):
    df = pd.DataFrame(columns=setupPlanCollumns)
    df = filesToDataframe(setupPlanList, df)
    writeToXlsx(df)
    return df