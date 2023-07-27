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
import configparser
import os
warnings.filterwarnings("ignore", category=UserWarning)

class exception(Exception):
    pass

config = configparser.ConfigParser()
config.read('config\config.ini')
setupPlanCollumns = config['DEFAULT'].get('setupPlanCollumns', 'GeoFilename,NumberOfParts,DimensionX,DimensionY,Area,Weight,MachiningTime').split(',')
excelFileDirectory = Path(config['DEFAULT'].get('excelFileDirectory', 'testing/xlsx'))
excelSourceName = config['DEFAULT'].get('excelSourceName', 'main.xlsx')
excelWorkbookName = config['DEFAULT'].get('excelWorkbookName', 'podatki')
excelLastFile = config['DEFAULT'].get('excelLastFile', 'main_output.xlsx')
excelOutputDir = Path(config['DEFAULT'].get('excelOutputDir', 'testing/xlsx'))

# Define the "Settings" menu
def saveSettingsToFile():
    # Save the current values of the variables to the configuration file
    config['DEFAULT']['setupPlanCollumns'] = ','.join(setupPlanCollumns)
    config['DEFAULT']['excelFileDirectory'] = str(excelFileDirectory)
    config['DEFAULT']['excelSourceName'] = excelSourceName
    config['DEFAULT']['excelWorkbookName'] = excelWorkbookName
    config['DEFAULT']['excelLastFile'] = excelLastFile
    config['DEFAULT']['excelOutputDir'] = str(excelOutputDir)
    with open('config\config.ini', 'w') as configfile:
        config.write(configfile)

def loadSettingsFromFile():
    config = configparser.ConfigParser()
    config.read('config\config.ini')
    global setupPlanCollumns, excelFileDirectory, excelSourceName, excelWorkbookName, excelLastFile, excelOutputDir
    setupPlanCollumns = config['DEFAULT'].get('setupPlanCollumns', 'GeoFilename,NumberOfParts,DimensionX,DimensionY,Area,Weight,MachiningTime').split(',')
    excelFileDirectory = Path(config['DEFAULT'].get('excelFileDirectory', 'testing/xlsx'))
    excelSourceName = config['DEFAULT'].get('excelSourceName', 'main.xlsx')
    excelWorkbookName = config['DEFAULT'].get('excelWorkbookName', 'podatki')
    excelLastFile = config['DEFAULT'].get('excelLastFile', 'main_output.xlsx')
    excelOutputDir = Path(config['DEFAULT'].get('excelOutputDir', 'testing/xlsx'))

def filesToDataframe(setupPlanList, df):
    database = sqlite3.connect('src/cache.db')
    cursor = database.cursor()
    cursor.execute("DROP TABLE IF EXISTS LabelPartData")    
    for file in setupPlanList:
        file = Path(file)
        if not file.is_file():
            raise exception("Datoteka ne obstja")
        if not (file.suffix == ".html" or file.suffix == ".HTML"):
            raise exception("Datoteka ni HTML")
        with open(file) as fp:
            setupPlan = BeautifulSoup(fp, "html.parser")
        singlePartData = getSinglePartSQL(setupPlan)
        singlePartData = singlePartData.replace("CREATE TABLE LabelPartData;\nALTER TABLE LabelPartData ADD COLUMN Count COUNTER PRIMARY KEY;", "CREATE TABLE LabelPartData (Count INTEGER PRIMARY KEY AUTOINCREMENT);")
        singlePartData += getSinglePartMachiningTime(setupPlan)
        if(singlePartData == ""):
            raise exception("Datoteka ne vsebuje podatkov o posameznih kosih")
        df = dataframeAppendFile(singlePartData, df)
    return df

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
    
def dataframeAppendFile(setupPlan, df):
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

def writeToXlsx(df, excelOutputName):
    try:
        workbook = pyxl.load_workbook(excelFileDirectory / excelSourceName)
        worksheet = workbook[excelWorkbookName]
        for row in dataframe_to_rows(df, header=True, index=False):
            worksheet.append(row)
        workbook.save(excelOutputDir / excelOutputName)
        workbook.close()
    except Exception as e:
        raise exception("Ni mogoče zapisati v excel datoteko")

def mainConversion(setupPlanList):
    loadSettingsFromFile()
    df = pd.DataFrame(columns=setupPlanCollumns)
    df = filesToDataframe(setupPlanList, df)
    global excelLastFile
    excelLastFile = os.path.basename(setupPlanList[0].replace(".HTML", ".xlsx"))
    writeToXlsx(df, excelLastFile)
    saveSettingsToFile()
    return df