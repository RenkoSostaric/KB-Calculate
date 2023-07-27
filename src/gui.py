import tkinter as tk
import tkinter.font as font
from tkinterdnd2 import DND_FILES, TkinterDnD
from  tkinter import ttk
from pathlib import Path
from PIL import Image, ImageTk
from PIL import Image, ImageTk, ImageFilter
import ctypes
import conversion
import time
import threading
import configparser
import os

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("../assets")
fileDataframe = None
ctypes.windll.shcore.SetProcessDpiAwareness(1)

# Load the values of the variables from the configuration file
config = configparser.ConfigParser()
config.read('config\config.ini')
setupPlanCollumns = config['DEFAULT'].get('setupPlanCollumns', 'GeoFilename,NumberOfParts,DimensionX,DimensionY,Area,Weight,MachiningTime').split(',')
excelFileDirectory = Path(config['DEFAULT'].get('excelFileDirectory', 'testing/xlsx'))
excelSourceName = config['DEFAULT'].get('excelSourceName', 'main.xlsx')
excelWorkbookName = config['DEFAULT'].get('excelWorkbookName', 'podatki')
excelLastFile = config['DEFAULT'].get('excelLastFile', 'main_output.xlsx')
excelOutputDir = Path(config['DEFAULT'].get('excelOuputDir', 'testing/xlsx'))

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

def saveSettings(setupPlanCollumnsInput, excelFileDirectoryInput, excelSourceNameInput, excelWorkbookNameInput, excelOutputDirInput):
    global setupPlanCollumns, excelFileDirectory, excelSourceName, excelWorkbookName, excelOutputDir
    setupPlanCollumns = [col.strip() for col in setupPlanCollumnsInput.get().split(',')]
    excelFileDirectory = Path(excelFileDirectoryInput.get())
    excelSourceName = excelSourceNameInput.get()
    excelWorkbookName = excelWorkbookNameInput.get()
    excelOutputDir = Path(excelOutputDirInput.get())
    saveSettingsToFile()

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

def setAppWindow(window):
    hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
    stylew = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    stylew = stylew & ~WS_EX_TOOLWINDOW
    stylew = stylew | WS_EX_APPWINDOW
    res = ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, stylew)
    window.wm_withdraw()
    window.after(10, lambda: window.wm_deiconify())

def mainWindow():
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
    window.after(10, lambda: setAppWindow(window))
    window.iconbitmap("assets/icon.ico")
    window.lift()
    window.attributes('-topmost',True)
    window.after_idle(window.attributes,'-topmost',False)
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
    
    def menuPopup(window):
        menu = tk.Menu(navBar, tearoff=0)
        menu.add_command(label="Nastavitve", command=lambda: openSettingsWindow(window))
        menu.add_command(label="Izhod", command=exitGUI)
        menu.tk_popup(navBar.winfo_rootx(), navBar.winfo_rooty()+30)
        
    menuButton = tk.Button(
        navBar,
        text = "\uf0c9",
        font = font.Font(family="FontAwesome", size=12),
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        command = lambda: menuPopup(window)
    )
    menuButton.place(relx=0.01, rely=0.5, anchor="w")
    
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
    fileWindow(window)
    window.mainloop()
    
    
def fileWindow(window):
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
            waitWindow(window)
            conversion_thread = threading.Thread(target=mainConversionThread, args=(window, dropped_files))
            conversion_thread.start()
    
    leftCanvas.drop_target_register(DND_FILES) # type: ignore
    leftCanvas.dnd_bind('<<Drop>>', handle_drop) # type: ignore
    window.update()
    
def mainConversionThread(window, file_list):
    try:
        fileDataframe = conversion.mainConversion(file_list)
        window.after(0, resultsWindow, window, fileDataframe)    
    except conversion.exception as e:
        errorWindow(window, e)
    
def waitWindow(window):
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
    def rotate_image():
        for angle in range(0, 720, 10):
            throbberImage = ImageTk.PhotoImage(
                Image.open(relative_to_assets("throbber.png")).rotate(angle, resample=Image.BICUBIC)
            )
            leftCanvas.itemconfig(throbberItem, image=throbberImage)
            leftCanvas.throbberImage = throbberImage  # type: ignore
            window.update()
            time.sleep(0.01)
    leftCanvas.create_text(
        130.0,
        340.0,
        anchor="center",
        text="Nalaganje...",
        fill="#333333",
        font=("Montserrat Medium", 16 * -1),
        tags="throbberText"
    )
    rotate_image()
    leftCanvas.drop_target_unregister()
    window.update()
    
def resultsWindow(window, fileDataframe):
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
        font=("Montserrat Medium", 16 * -1),
        tags="checkmarkText"
    ) 
    buttonCancelImage = Image.open(relative_to_assets("button_cancel.png"))
    buttonCancelImage = buttonCancelImage.resize((92, 40), resample=Image.BICUBIC)
    buttonCancelImage = ImageTk.PhotoImage(buttonCancelImage) 
    rightCanvas.button_cancel_image = buttonCancelImage # type: ignore
    button_cancel = tk.Button(
        image=buttonCancelImage,
        borderwidth=0,
        highlightthickness=0,
        background="#FFFFFF",
        activebackground="#FFFFFF",
        command=lambda: resetWindow(window),
        relief="flat"
    )
    rightCanvas.create_window(350, 460, anchor="center", window=button_cancel, width=97, height=45)
        
    def openXlsxFile():
        loadSettingsFromFile()
        os.system(f'start excel "{excelOutputDir}/{excelLastFile}"')
    buttonXlsxImage = Image.open(relative_to_assets("button_xlsx.png"))
    buttonXlsxImage = buttonXlsxImage.resize((120, 40), resample=Image.BICUBIC)
    buttonXlsxImage = ImageTk.PhotoImage(buttonXlsxImage)
    rightCanvas.button_xlsx_image = buttonXlsxImage # type: ignore
    button_xlsx = tk.Button(
        image=buttonXlsxImage,
        borderwidth=0,
        highlightthickness=0,
        background="#FFFFFF",
        activebackground="#FFFFFF",
        command= openXlsxFile,
        relief="flat"
    )
    rightCanvas.create_window(465, 460, anchor="center", window=button_xlsx, width=125, height=45)
    
    
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
        height=18,
        selectmode="none",
    )
    rightCanvas.create_window(0, 0, anchor="nw", window=tree, width=540, height=450)
        
    tree.column(fileDataframe.columns[0], width=220, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[2], width=100, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[3], width=100, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[-1], width=120, stretch=False) #type: ignore

    for col in tree["columns"]:
        tree.heading(col, text=col)    
    for index, row in fileDataframe.iterrows():
        tree.insert("", "end", values=(row[fileDataframe.columns[0]], row[fileDataframe.columns[2]], row[fileDataframe.columns[3]], row[fileDataframe.columns[-1]])) #type: ignore
    
    window.update()
    
def resetWindow(window):
    leftCanvas.delete("checkmarkImage", "checkmarkText", "throbberText", "throbberImage", "throbberBackgroundImage")
    rightCanvas.delete("all")
    fileWindow(window)
    
def errorWindow(window, error):
    errorWindow = tk.Toplevel(window)
    windowWidth = 600
    windowHeight = 400
    screenWidth = errorWindow.winfo_screenwidth()
    screenHeight = errorWindow.winfo_screenheight()
    x = (screenWidth // 2) - (windowWidth // 2)
    y = (screenHeight // 2) - (windowHeight // 2)
    errorWindow.title("Napaka")
    errorWindow.geometry(f"{windowWidth}x{windowHeight}+{x}+{y}")
    errorWindow.overrideredirect(True)
    errorWindow.resizable(False, False)
    errorWindow.iconbitmap("assets/icon.ico")
    errorWindow.lift()
    errorWindow.attributes('-topmost',True)
    canvas = tk.Canvas(errorWindow, width=windowWidth, height=windowHeight, bg="#FFFFFF")
    canvas.pack()
    
    errorMessage = tk.Label(canvas, text=error, font=("Montserrat Medium", 16), bg="#FFFFFF", fg="#333333")
    errorMessage.place(relx=0.5, rely=0.7, anchor="center")
    
    errorImage = tk.PhotoImage(file=relative_to_assets("error.png"))
    canvas.create_image(windowWidth // 2, windowHeight // 2 - 50, image=errorImage)
    canvas.errorImage = errorImage # type: ignore

    buttonCancelImage = Image.open(relative_to_assets("button_again.png"))
    buttonCancelImage = buttonCancelImage.resize((160, 40), resample=Image.BICUBIC)
    buttonCancelImage = ImageTk.PhotoImage(buttonCancelImage) 
    canvas.buttonCancelImage = buttonCancelImage # type: ignore
    buttonCancel = tk.Button(
        canvas,
        image=buttonCancelImage,
        borderwidth=0,
        highlightthickness=0,
        background="#FFFFFF",
        activebackground="#FFFFFF",
        command=lambda: [resetWindow(window), errorWindow.destroy()],
        relief="flat"
    )
    buttonCancel.place(relx=0.5, rely=0.9, anchor="center")
    window.update()

def openSettingsWindow(window):
    settingsWindow = tk.Toplevel(window)
    windowWidth = 400
    windowHeight = 360
    screenWidth = settingsWindow.winfo_screenwidth()
    screenHeight = settingsWindow.winfo_screenheight()
    x = (screenWidth // 2) - (windowWidth // 2)
    y = (screenHeight // 2) - (windowHeight // 2)
    settingsWindow.title("Nastavitve")
    settingsWindow.geometry(f"{windowWidth}x{windowHeight}+{x}+{y}")
    settingsWindow.overrideredirect(True)
    settingsWindow.resizable(False, False)
    settingsWindow.iconbitmap("assets/icon.ico")
    settingsWindow.lift()
    settingsWindow.attributes('-topmost',True)

    setupPlanCollumnsLabel = tk.Label(settingsWindow, text="Setup Plan Columns:", font=font.Font(family="Montserrat Medium", size=10))
    setupPlanCollumnsLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    setupPlanCollumnsInput = tk.Entry(settingsWindow, width=50)
    setupPlanCollumnsInput.insert(0, ','.join(setupPlanCollumns))
    setupPlanCollumnsInput.pack(side="top", fill=tk.X, padx=20)

    excelFileDirectoryLabel = tk.Label(settingsWindow, text="Excel File Directory:", font=font.Font(family="Montserrat Medium", size=10))
    excelFileDirectoryLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelFileDirectoryInput = tk.Entry(settingsWindow, width=50)
    excelFileDirectoryInput.insert(0, str(excelFileDirectory))
    excelFileDirectoryInput.pack(side="top", fill=tk.X, padx=20)

    excelSourceNameLabel = tk.Label(settingsWindow, text="Excel Source Name:", font=font.Font(family="Montserrat Medium", size=10))
    excelSourceNameLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelSourceNameInput = tk.Entry(settingsWindow, width=50)
    excelSourceNameInput.insert(0, excelSourceName)
    excelSourceNameInput.pack(side="top", fill=tk.X, padx=20)

    excelWorkbookNameLabel = tk.Label(settingsWindow, text="Excel Workbook Name:", font=font.Font(family="Montserrat Medium", size=10))
    excelWorkbookNameLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelWorkbookNameInput = tk.Entry(settingsWindow, width=50)
    excelWorkbookNameInput.insert(0, excelWorkbookName)
    excelWorkbookNameInput.pack(side="top", fill=tk.X, padx=20)
    
    excelOuputDirLabel = tk.Label(settingsWindow, text="Excel Ouput Directory:", font=font.Font(family="Montserrat Medium", size=10))
    excelOuputDirLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelOuputDirInput = tk.Entry(settingsWindow, width=50)
    excelOuputDirInput.insert(0, excelOutputDir)
    excelOuputDirInput.pack(side="top", fill=tk.X, padx=20)

    close_button = tk.Button(settingsWindow, text="Zapri", command=lambda:settingsWindow.destroy())
    close_button.place(relx=0.78, rely=0.94, anchor="center")
    save_button = tk.Button(settingsWindow, text="Shrani", command=lambda:saveSettings(setupPlanCollumnsInput, excelFileDirectoryInput, excelSourceNameInput, excelWorkbookNameInput, excelOuputDirInput))
    save_button.place(relx=0.9, rely=0.94, anchor="center")

if __name__ == "__main__":
    mainWindow()