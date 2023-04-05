import pandas as pd;

htmlji=["1a.html", "1aleks.html"]
for html in htmlji:
    table = pd.read_html(html)
    #print(table)
    df={
        "Machine": [table[1].iloc[0,1].split(" ")[0]],
        "Blank": [table[1].iloc[9,1]],
        "Weight": [table[1].iloc[10,1]],
        "Machining time": [table[4].iloc[9,2]], 
        "Scrap": [table[1].iloc[15,1]],
        "Surface": [table[4].iloc[7,2]],
        "Št. lukenj": [(table[3].iloc[2,13])],
        "Laser": [table[4].iloc[10,2]]
        }
    data = pd.DataFrame(df)
    print(data)