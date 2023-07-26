from pathlib import Path
import pandas as pd
import tkinter as tk
from  tkinter import ttk
import ctypes
import tkinter.font as tkFont

ctypes.windll.shcore.SetProcessDpiAwareness(1)


OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


window = tk.Tk()

window.geometry("800x500")
window.configure(bg = "#FFFFFF")


canvas = tk.Canvas(
    window,
    bg = "#FFFFFF",
    height = 500,
    width = 800,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
canvas.create_rectangle(
    0.0,
    0.0,
    260.0,
    500.0,
    fill="#FFFFFF",
    outline="")

canvas.create_rectangle(
    260.0,
    0.0,
    262.0,
    500.0,
    fill="#D9D9D9",
    outline="")

image_image_1 = tk.PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    129.0,
    249.99996948242188,
    image=image_image_1
)

image_image_2 = tk.PhotoImage(
    file=relative_to_assets("image_2.png"))
image_2 = canvas.create_image(
    129.0,
    249.0,
    image=image_image_2
)

button_image_1 = tk.PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = tk.Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: print("button_1 clicked"),
    relief="flat"
)
button_1.place(
    x=670.0,
    y=450.0,
    width=120.0,
    height=40.0
)

button_image_2 = tk.PhotoImage(
    file=relative_to_assets("button_2.png"))
button_2 = tk.Button(
    image=button_image_2,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: print("button_2 clicked"),
    relief="flat"
)
button_2.place(
    x=540.0,
    y=450.0,
    width=120.0,
    height=40.0
)
window.resizable(False, False)

df = pd.read_excel('C:\\Users\\Uporabnik\\Documents\\KB Calculate\\testing\\xlsx\\main_output.xlsx')
#if any of the values in the first collumn of df more then 26 charracters then show only first 26 and add ... at the end
df[df.columns[0]] = df[df.columns[0]].apply(lambda x: x[:26] + '...' if len(x) > 26 else x)


print(df)

# create a treeview with data from dataframe

style = ttk.Style()
style.configure("style.Treeview", borderwidth=0, highlightthickness=0, font=('Montserrat', 10))
style.configure("style.Treeview.Heading", font=('Montserrat Semibold', 8))
style.layout("style.Treeview", [('style.Treeview.treearea', {'sticky': 'nswe'})])



tree = ttk.Treeview(
    window, 
    columns=(df.columns[0], df.columns[2], df.columns[3],df.columns[-1]), #type: ignore
    show="headings",
    style="style.Treeview",
    height=18
)



tree.place(x=270, y=30)
    

tree.column(df.columns[0], width=220) #type: ignore
tree.column(df.columns[2], width=100) #type: ignore
tree.column(df.columns[3], width=100) #type: ignore
tree.column(df.columns[-1], width=100) #type: ignore

# Add the rows to the Treeview
for col in tree["columns"]:
    tree.heading(col, text=col)
for index, row in df.iterrows():
    tree.insert("", "end", values=(row[df.columns[0]], row[df.columns[2]], row[df.columns[3]], row[df.columns[-1]])) #type: ignore

window.mainloop()