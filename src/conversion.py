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

# Define global constants
OUTPUT_PATH = Path(__file__).parent
CACHE_PATH = OUTPUT_PATH / Path("cache")
CONFIG_PATH = OUTPUT_PATH / Path("config")

# Define the custom exception
class exception(Exception):
    pass

# Get configuration from the config file
config = configparser.ConfigParser()
config.read(CONFIG_PATH / Path("config.ini"))
setupPlanCollumns = config['DEFAULT'].get('setupPlanCollumns', 'GeoFilename,NumberOfParts,DimensionX,DimensionY,Area,Weight,MachiningTime').split(',')
excelFileDirectory = Path(config['DEFAULT'].get('excelFileDirectory', 'testing/xlsx'))
excelSourceName = config['DEFAULT'].get('excelSourceName', 'main.xlsx')
excelWorkbookName = config['DEFAULT'].get('excelWorkbookName', 'podatki')
excelLastFile = config['DEFAULT'].get('excelLastFile', 'main_output.xlsx')
excelOutputDir = Path(config['DEFAULT'].get('excelOutputDir', 'testing/xlsx'))

# Function to save the current settings to the config file
def saveSettingsToFile():
    config['DEFAULT']['setupPlanCollumns'] = ','.join(setupPlanCollumns)
    config['DEFAULT']['excelFileDirectory'] = str(excelFileDirectory)
    config['DEFAULT']['excelSourceName'] = excelSourceName
    config['DEFAULT']['excelWorkbookName'] = excelWorkbookName
    config['DEFAULT']['excelLastFile'] = excelLastFile
    config['DEFAULT']['excelOutputDir'] = str(excelOutputDir)
    with open(CONFIG_PATH / Path("config.ini"), 'w') as configfile:
        config.write(configfile)

# Function to load the settings from the config file
def loadSettingsFromFile():
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH / Path("config.ini"))
    global setupPlanCollumns, excelFileDirectory, excelSourceName, excelWorkbookName, excelLastFile, excelOutputDir
    setupPlanCollumns = config['DEFAULT'].get('setupPlanCollumns', 'GeoFilename,NumberOfParts,DimensionX,DimensionY,Area,Weight,MachiningTime').split(',')
    excelFileDirectory = Path(config['DEFAULT'].get('excelFileDirectory', 'testing/xlsx'))
    excelSourceName = config['DEFAULT'].get('excelSourceName', 'main.xlsx')
    excelWorkbookName = config['DEFAULT'].get('excelWorkbookName', 'podatki')
    excelLastFile = config['DEFAULT'].get('excelLastFile', 'main_output.xlsx')
    excelOutputDir = Path(config['DEFAULT'].get('excelOutputDir', 'testing/xlsx'))
    
# Function that goes through all the files and converts them to a dataframe
def filesToDataframe(setupPlanList, df):
    # Connect to the temp database
    database = sqlite3.connect(CACHE_PATH / Path("cache.db"))
    cursor = database.cursor()
    cursor.execute("DROP TABLE IF EXISTS LabelPartData")
    # Iterate through all the files 
    for file in setupPlanList:
        file = Path(file)
        # Check if the file is an HTML file
        # Check if the file exists
        if not file.is_file():
            raise exception("Datoteka ne obstaja")
        if not (file.suffix == ".html" or file.suffix == ".HTML"):
            raise exception("Datoteka ni HTML")
        # Open the file and parse it to a BeautifulSoup object
        with open(file) as fp:
            setupPlan = BeautifulSoup(fp, "html.parser")
        # Get the SQL part of the file
        singlePartData = getSinglePartSQL(setupPlan)
        # Format the SQL so it works with sqlite3
        singlePartData = singlePartData.replace("CREATE TABLE LabelPartData;\nALTER TABLE LabelPartData ADD COLUMN Count COUNTER PRIMARY KEY;", "CREATE TABLE LabelPartData (Count INTEGER PRIMARY KEY AUTOINCREMENT);")
        # Get the machining time for each part becouse it is not in the SQL
        singlePartData += getSinglePartMachiningTime(setupPlan)
        # Check if the file contains data about the parts
        if(singlePartData == ""):
            raise exception("Datoteka ne vsebuje podatkov o posameznih kosih")
        # Execute the SQL and update the dataframe
        df = dataframeAppendFile(singlePartData, df)
    return df

# Function that goes through the html file, gets the machining time for each part and adds it to the SQL script
def getSinglePartMachiningTime(setupPlan):
    # Create the new column in SQL
    singlePartSQL = "ALTER TABLE LabelPartData ADD COLUMN MachiningTime VARCHAR(255);\n"
    # Find the table with the single part data
    singlePartTable = setupPlan.find_all(string=re.compile(r"(INFORMATION ON SINGLE PART)|(EINZELTEILINFORMATION)"))[-1].find_parent("table")
    # Find all the rows that contain the machining times
    machiningTimesRows = singlePartTable.find_all(string=re.compile(r"(MACHINING TIME)|(BEARBEITUNGSZEIT)"))
    # Iterate through all the rows and add the machining time to the SQL script
    for i in range(len(machiningTimesRows)):
        machiningTime = machiningTimesRows[i].find_parent("tr").find_all("td")[1].text
        # Check if the machining time is not empty and add the data to the SQL script
        if(machiningTime != ""):
            singlePartSQL += "UPDATE LabelPartData SET MachiningTime='" + re.findall(r"\d+\.\d+ min", machiningTime)[0] + "' WHERE COUNT = " + str(i+1) + ";"
    return singlePartSQL
    
# Function that goes through the html file and gets the SQL scipt for the single part data
def getSinglePartSQL(setupPlan):
    # Find all the comments in the file
    comments = setupPlan.findAll(string=lambda string:isinstance(string, Comment))
    # Iterate through all the comments and find the one that contains the SQL script
    for comment in comments:        
        commentSoup = BeautifulSoup(comment, "lxml")    
        sqlRow = commentSoup.find("sql", string="CREATE TABLE LabelPartData")
        # Check if the comment contains the SQL script
        if sqlRow is not None:
            # Get the SQL script
            setupPlan = sqlRow.text + ";\n"
            while sqlRow.next_sibling is not None:
                sqlRow = sqlRow.next_sibling
                if hasattr(sqlRow, 'name') and sqlRow.name == "sql": # type: ignore
                    setupPlan += sqlRow.text + ";\n"
            return setupPlan
    return ""

# Function that executes the SQL script and updates the dataframe
def dataframeAppendFile(setupPlan, df):
    # Connect to the temp database
    database = sqlite3.connect(CACHE_PATH / Path("cache.db"))
    cursor = database.cursor()
    cursor.execute("DROP TABLE IF EXISTS LabelPartData")    
    cursor.executescript(setupPlan)
    database.commit()
    # Get the data from the database
    sql_query = pd.read_sql_query ('''
        SELECT
        '''+ ", ".join(setupPlanCollumns) + '''
        FROM LabelPartData
    ''', database)
    # Add the data to the dataframe if it's empty create a new one otherwise append it
    if df.empty:
        df = pd.DataFrame(sql_query)
    else:
        df = pd.concat([df, pd.DataFrame(sql_query)], ignore_index=True)
    # Format the GeoFilename so it only contains the filename not the whole path
    df["GeoFilename"] = df["GeoFilename"].replace(to_replace=r".*/(.+\.GEO)", value=r"\g<1>", regex=True)
    return df

# Function that saves the data from the dataframe to an excel file
def writeToXlsx(df, excelOutputName):
    try:
        # Open the main excel file
        workbook = pyxl.load_workbook(excelFileDirectory / excelSourceName)
        worksheet = workbook[excelWorkbookName]
        # Iterate through all the rows in the dataframe and add them to the excel file
        for row in dataframe_to_rows(df, header=True, index=False):
            worksheet.append(row)
        # Save the excel file as a new file
        workbook.save(excelOutputDir / excelOutputName)
        workbook.close()
    except Exception as e:
        raise exception("Ni mogoče zapisati v excel datoteko")

# Function that controls the conversion process and returns the dataframe for the GUI preview
def mainConversion(setupPlanList):
    # Load confinguration from the config file
    loadSettingsFromFile()
    # Create and fill the dataframe
    df = pd.DataFrame(columns=setupPlanCollumns)
    df = filesToDataframe(setupPlanList, df)
    # Create the excel file
    global excelLastFile
    excelLastFile = os.path.basename(setupPlanList[0].replace(".HTML", ".xlsx"))
    writeToXlsx(df, excelLastFile)
    # Save the settings with created excel file name to the config file
    saveSettingsToFile()
    return df