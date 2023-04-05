from bs4 import BeautifulSoup
import re
import pandas as pd;

searchList=["MACHINE","NC PROGRAM PATH", "MATERIAL ID", "BLANK", "WEIGHT", "MACHINING TIME", "LASER TOTAL CUTTING LENGTH", "NUMBER OF PROGRAMME RUNS", "SCRAP"]
dataList=[""] * 9

""" df = pd.DataFrame({
    "MACHINE" : [],
    "NC PROGRAM PATH" : [],
    "MATERIAL ID" : [],
    "BLANK" : [],
    "WEIGHT" : [],
    "MACHINING TIME" : [],
    "LASER TOTAL CUTTING LENGTH" : [],
    "NUMBER OF PROGRAMME RUNS" : [],
    "SCRAP" : []}); """
    
df = pd.DataFrame(columns=["MACHINE","NC PROGRAM PATH", "MATERIAL ID", "BLANK", "WEIGHT", "MACHINING TIME", "LASER TOTAL CUTTING LENGTH", "NUMBER OF PROGRAMME RUNS", "SCRAP"])

with open("77581.html") as fp:
    setupDocument = BeautifulSoup(fp, "html.parser")
    
row=[] 
for column in df:
    searchElement = setupDocument.find("td", string=re.compile(column))
    foundElement = searchElement.find_next_sibling("td")
    clean = foundElement.text
    cell = " ".join(re.findall(r"\S+(?<!\s)", clean))
    row.append(cell)
dfTemp = pd.DataFrame([row])
dfTemp.columns = df.columns
df = pd.concat([df, dfTemp], axis=0, ignore_index=True) 
print(df)