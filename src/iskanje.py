from html.parser import HTMLParser
import pandas as pd;
import openpyxl;
import os
from openpyxl.utils.dataframe import dataframe_to_rows

htmlList=["1a.html", "77582.html", "77581.html"]

iskaniPodatki=["MACHINE:", "BLANK:", "WEIGHT:", "MACHINING TIME:", "SCRAP:", "LASER TOTAL CUTTING LENGTH:"]
najdeniPodatki=[[0, [int, int]], [0, [int, int]], [0, [int, int]], [0, [int, int]], [0, [int, int]], [0, [int, int]]]

df = pd.DataFrame()

def ustvariSeznamLokacij(htmlParse):
    global podatkiHTML
    podatkiHTML = pd.read_html(htmlParse)
    for i in range(len(podatkiHTML)):
        tabelaHTML=pd.DataFrame(podatkiHTML[i])
        for j in range(len(iskaniPodatki)):
            listPozicij=[]
            result = tabelaHTML.isin([iskaniPodatki[j]])
            seriesObj = result.any()
            columnNames = list(seriesObj[seriesObj == True].index)
            for col in columnNames:
                rows = list(result[col][result[col] == True].index)
                for row in rows:
                    listPozicij=[row, col]
            if listPozicij and tabelaHTML.loc[row, col]==iskaniPodatki[j]:
                najdeniPodatki[j][0]=i
                najdeniPodatki[j][1][0]=row
                najdeniPodatki[j][1][1]=col

def ustvariGlavoTabele():
    for glava in reversed(iskaniPodatki):
        df.insert(len(df), glava, [], True) 

def vstaviVrstico():
    global df
    vrstica=[]
    for i in range(len(najdeniPodatki)):
        x=podatkiHTML[najdeniPodatki[i][0]].iloc[najdeniPodatki[i][1][0], najdeniPodatki[i][1][1]+1]
        vrstica.append(x)
    dfTemp = pd.DataFrame([vrstica])
    dfTemp.columns = df.columns
    df = pd.concat([df, dfTemp], axis=0, ignore_index=True)

def izvoziExcel():
    file = r"test.xlsx"
    if os.path.isfile(file):
        workbook = openpyxl.load_workbook(file)
        sheet = workbook['podatki']
        for row in dataframe_to_rows(df, header = False, index = False):
            sheet.append(row)
        workbook.save("test_after.xlsx")
        workbook.close()


ustvariGlavoTabele()
for html in htmlList:
    ustvariSeznamLokacij(html)
    vstaviVrstico()
izvoziExcel()
print("-" * 200)
print(df)
print("-" * 200)