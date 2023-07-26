import tkinter as tk
import tkinter.font as font
from tkinterdnd2 import DND_FILES, TkinterDnD
from  tkinter import ttk
from pathlib import Path
from PIL import Image, ImageTk
from PIL import Image, ImageTk, ImageFilter
import ctypes
import math
import conversion

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("../assets")
fileDataframe = None
ctypes.windll.shcore.SetProcessDpiAwareness(1)

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

def set_appwindow(window):
    hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
    stylew = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    stylew = stylew & ~WS_EX_TOOLWINDOW
    stylew = stylew | WS_EX_APPWINDOW
    res = ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, stylew)
    window.wm_withdraw()
    window.after(10, lambda: window.wm_deiconify())

def mainWindow():
    global window
    window = TkinterDnD.Tk()
    window_width = 800
    window_height = 530
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    window.title("KB Calculate")
    window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    window.overrideredirect(True)
    window.resizable(False, False)
    window.after(10, lambda: set_appwindow(window))
    window.iconbitmap("assets/icon.ico")
    
    navBar = tk.Frame(
        window,
        bg = "#EAEAEA",
        height = 30,
        width = 800,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
    )
    navBar.pack(side = "top", anchor="nw")
    
    def exitGUI():
        window.destroy()
    
    exitButton = tk.Button(
        navBar,
        text = "\uf00d",
        font = font.Font(family="FontAwesome", size=14),
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        command = exitGUI
    )
    exitButton.place(relx=0.99, rely=0.5, anchor="e")
    
    windowTitle = tk.Label(
        navBar,
        text = "KB Calculate",
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        font = font.Font(family="Montserrat Medium", size=12)
    )
    windowTitle.place(relx = 0.5, rely = 0.5, anchor = "center")
        
    windowIconImage = Image.open("assets/icon.png")
    windowIconImage = windowIconImage.resize((55, 35), resample=Image.BICUBIC)
    windowIconImage = ImageTk.PhotoImage(windowIconImage)
    windowIcon = tk.Label(
        navBar,
        image=windowIconImage,
        bg = "#EAEAEA",
        bd = 0,
        highlightthickness = 0,
        
    )
    windowIcon.place(relx=0.035, rely=0.5, anchor="center")
    
    global leftCanvas
    leftCanvas = tk.Canvas(
        window,
        bg = "#FFFFFF",
        height = 500,
        width = 260,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
    )
    leftCanvas.place(x = 0, y = 30)
    leftCanvas.create_line(259, 0, 259, 500, fill="#D9D9D9", width=2)
    
    global rightCanvas
    rightCanvas = tk.Canvas(
        window,
        bg = "#FFFFFF",
        height = 500,
        width = 540,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
    )
    rightCanvas.place(x = 260, y = 30)  
    
    backgroundLogo = tk.PhotoImage(
        file=relative_to_assets("background_logo.png")
    )
    rightCanvas.create_image(
        270, 
        250, 
        image=backgroundLogo, 
        anchor="center",
        tags="backgroundLogo"
    )
    fileWindow()
    window.mainloop()
    
    
def fileWindow():
    fileUploadImage = tk.PhotoImage(
        file=relative_to_assets("file_upload.png")
    )
    leftCanvas.fileUploadImage = fileUploadImage # type: ignore
    leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=fileUploadImage,
        tags="fileUploadImage"
    )
    
    leftCanvas.create_text(
        130.0,
        340.0,
        anchor="center",
        text="Povleci in spusti datoteke",
        fill="#333333",
        font=("Montserrat Medium", 16 * -1),
        tags="fileUploadText"
    )
    
    dropped_files = []
    def handle_drop(event):
        file_paths = event.data.replace('{', '').replace('}', '\n').split('\n')
        file_paths = [path.strip() for path in file_paths if path.strip()]
        for file_path in file_paths:
            dropped_files.append(file_path)
        if len(dropped_files) > 0:
            waitWindow()
            fileDataframe = conversion.mainConversion(dropped_files)
            window.after(1000, lambda: resultsWindow(fileDataframe))
    
    leftCanvas.drop_target_register(DND_FILES) # type: ignore
    leftCanvas.dnd_bind('<<Drop>>', handle_drop) # type: ignore
    window.update()
    
def waitWindow():
    leftCanvas.delete("fileUploadImage", "fileUploadText")
    throbberBackgroundImage = tk.PhotoImage(
        file=relative_to_assets("throbber_background.png")
    )
    leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=throbberBackgroundImage,
        tags="throbberBackgroundImage"
    )
    leftCanvas.throbberBackgroundImage = throbberBackgroundImage #type: ignore
    throbberImage = tk.PhotoImage(
        file=relative_to_assets("throbber.png")
    )
    throbberItem=leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=throbberImage,
        tags="throbberImage"
    )
    def rotate_image(angle=0):
        throbberImage = ImageTk.PhotoImage(
            Image.open(relative_to_assets("throbber.png")).rotate(angle, resample=Image.BICUBIC)
        )
        leftCanvas.itemconfig(throbberItem, image=throbberImage)
        leftCanvas.throbberImage = throbberImage  # type: ignore
        window.after(10, rotate_image, angle - 10)
    rotate_image()
    leftCanvas.create_text(
        130.0,
        340.0,
        anchor="center",
        text="Nalaganje...",
        fill="#333333",
        font=("Montserrat Medium", 16 * -1),
        tags="throbberText"
    )
    window.update()
    
def resultsWindow(fileDataframe):
    leftCanvas.delete("throbberImage", "throbberBackgroundImage", "throbberText")
    rightCanvas.delete("backgroundLogo")
    checkmarkImage = tk.PhotoImage(
        file=relative_to_assets("checkmark.png")
    )
    leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=checkmarkImage,
        tags="checkmarkImage"
    )
    leftCanvas.checkmarkImage = checkmarkImage # type: ignore
    leftCanvas.create_text(
        130.0,
        340.0,
        anchor="center",
        text="Nalaganje končano!",
        fill="#333333",
        font=("Montserrat Medium", 16 * -1)
    )
    buttonCancelImage = Image.open(relative_to_assets("button_cancel.png"))
    buttonCancelImage = buttonCancelImage.resize((93, 40), resample=Image.BICUBIC)
    buttonCancelImage = ImageTk.PhotoImage(buttonCancelImage)
    rightCanvas.button_cancel_image = buttonCancelImage # type: ignore
    button_cancel = tk.Button(
        image=buttonCancelImage,
        borderwidth=0,
        highlightthickness=0,
        background="#FFFFFF",
        command=lambda: print("button_cancel clicked"),
        relief="flat"
    )
    button_cancel.place(    
        x=570.0,
        y=480.0,
        width=93.0,
        height=40.0
    )
    
    
    
    buttonXlsxImage = Image.open(relative_to_assets("button_xlsx.png"))
    buttonXlsxImage = buttonXlsxImage.resize((120, 40), resample=Image.BICUBIC)
    buttonXlsxImage = ImageTk.PhotoImage(buttonXlsxImage)
    rightCanvas.button_xlsx_image = buttonXlsxImage # type: ignore
    button_xlsx = tk.Button(
        image=buttonXlsxImage,
        borderwidth=0,
        highlightthickness=0,
        background="#FFFFFF",
        command=lambda: print("button_xlsx clicked"),
        relief="flat"
    )
    button_xlsx.place(    
        x=670.0,
        y=480.0,
        width=120.0,
        height=40.0
    )
    
    fileDataframe[fileDataframe.columns[0]] = fileDataframe[fileDataframe.columns[0]].apply(lambda x: x[:26] + '...' if len(x) > 26 else x)
    style = ttk.Style()
    style.configure("style.Treeview", borderwidth=0, highlightthickness=0, font=('Montserrat', 10))
    style.configure("style.Treeview.Heading", font=('Montserrat Semibold', 9))
    style.layout("style.Treeview", [('style.Treeview.treearea', {'sticky': 'nswe'})])
    tree = ttk.Treeview(
        rightCanvas, 
        columns=(fileDataframe.columns[0], fileDataframe.columns[2], fileDataframe.columns[3],fileDataframe.columns[-1]), #type: ignore
        show="headings",
        style="style.Treeview",
        height=18
    )
    tree.place(x=0, y=0)
        
    tree.column(fileDataframe.columns[0], width=220, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[2], width=100, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[3], width=100, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[-1], width=120, stretch=False) #type: ignore

    for col in tree["columns"]:
        tree.heading(col, text=col)    
    for index, row in fileDataframe.iterrows():
        tree.insert("", "end", values=(row[fileDataframe.columns[0]], row[fileDataframe.columns[2]], row[fileDataframe.columns[3]], row[fileDataframe.columns[-1]])) #type: ignore

    window.update()

if __name__ == "__main__":
    mainWindow()