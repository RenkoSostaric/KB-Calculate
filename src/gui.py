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
import pyglet
import sys
# pyglet will use the win32 gdi font renderer instead of the default freetype renderer
pyglet.options["win32_gdi_font"] = True

# Set the DPI awareness of TKInter
ctypes.windll.shcore.SetProcessDpiAwareness(1)

if getattr(sys, 'frozen', False):
    import pyi_splash # type: ignore # Import the splash screen if the app is compiled


# Define some constants
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("assets")
RESOURCES_PATH = OUTPUT_PATH / Path("resources")
CONFIG_PATH = OUTPUT_PATH / Path("config")

# Define the global dataFrame variable
fileDataframe = None

# Import fonts from the resources folder
pyglet.font.add_file(os.path.normpath(RESOURCES_PATH / Path("FontAwesome-Solid.ttf")))
pyglet.font.add_file(os.path.normpath(RESOURCES_PATH / Path("Montserrat-Medium.ttf")))
pyglet.font.add_file(os.path.normpath(RESOURCES_PATH / Path("Montserrat-Semibold.ttf")))

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
        print("Saving settings")
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

# Function to save the settings from the GUI
def saveSettings(setupPlanCollumnsInput, excelFileDirectoryInput, excelSourceNameInput, excelWorkbookNameInput, excelOutputDirInput):
    global setupPlanCollumns, excelFileDirectory, excelSourceName, excelWorkbookName, excelOutputDir
    setupPlanCollumns = [col.strip() for col in setupPlanCollumnsInput.get().split(',')]
    excelFileDirectory = Path(excelFileDirectoryInput.get())
    excelSourceName = excelSourceNameInput.get()
    excelWorkbookName = excelWorkbookNameInput.get()
    excelOutputDir = Path(excelOutputDirInput.get())
    saveSettingsToFile()

# Function that returns the path to the assets folder
def relativeToAssets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

# Function to set the window in the taskbar
def setAppWindow(window):
    hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
    stylew = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    stylew = stylew & ~WS_EX_TOOLWINDOW
    stylew = stylew | WS_EX_APPWINDOW
    res = ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, stylew)
    window.wm_withdraw()
    window.after(10, lambda: window.wm_deiconify())

# Function to create the main window
def mainWindow():
    # Create the main window and configure it
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
    window.iconbitmap(relativeToAssets("ikona.ico"))
    window.lift()
    window.attributes('-topmost',True)
    # Add the navBar
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
    def moveApp(e):
        window.geometry(f'+{e.x_root}+{e.y_root}')
    navBar.bind('<B1-Motion>', moveApp)
    # Function for the exit button
    def exitGUI():
        window.destroy()
    # Add the exit button
    exitButton = tk.Button(
        navBar,
        text = "\uf00d",
        font = ("Font Awesome 6 Free Solid", 12),
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        command = exitGUI
    )
    exitButton.place(relx=0.99, rely=0.5, anchor="e")
    # Function to open the settings window
    def menuPopup(window):
        menu = tk.Menu(navBar, tearoff=0)
        menu.add_command(label="Nastavitve", command=lambda: openSettingsWindow(window))
        menu.add_command(label="Izhod", command=exitGUI)
        menu.tk_popup(navBar.winfo_rootx(), navBar.winfo_rooty()+30)
    # Add the menu button
    menuButton = tk.Button(
        navBar,
        text = "\uf0c9",
        font = ("Font Awesome 6 Free Solid", 12),
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        command = lambda: menuPopup(window)
    )
    menuButton.place(relx=0.01, rely=0.5, anchor="w")
    # Add the window title
    windowTitle = tk.Label(
        navBar,
        text = "KB Calculate",
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        font = ("Montserrat Medium", 12)
    )
    windowTitle.place(relx = 0.5, rely = 0.5, anchor = "center")
    # Add the left and right canvas
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
    # Add the background logo
    backgroundLogo = tk.PhotoImage(
        file=relativeToAssets("background_logo.png")
    )
    rightCanvas.create_image(
        270, 
        250, 
        image=backgroundLogo, 
        anchor="center",
        tags="backgroundLogo"
    )
    # Call the fileWindow function to draw the file window
    fileWindow(window)
    # Close the splash screen
    if getattr(sys, 'frozen', False):
        pyi_splash.close()
    window.mainloop()

# Function to draw the file window
def fileWindow(window):
    # Add the file window drag and drop image
    fileUploadImage = tk.PhotoImage(
        file=relativeToAssets("file_upload.png")
    )
    leftCanvas.fileUploadImage = fileUploadImage # type: ignore # Prevent garbage collection
    leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=fileUploadImage,
        tags="fileUploadImage"
    )
    # Add the file window drag and drop text
    leftCanvas.create_text(
        130.0,
        340.0,
        anchor="center",
        text="Povleci in spusti datoteke",
        fill="#333333",
        font=("Montserrat Medium", 12),
        tags="fileUploadText"
    )
    # Variable that saves the dropped files
    dropped_files = []
    # Function to handle the drop event
    def handle_drop(event):
        file_paths = event.data.replace('.HTML', '.HTML;').replace('{', '').replace('}', '').split(';')
        file_paths = [path.strip() for path in file_paths if path.strip()]
        for file_path in file_paths:
            if(file_path != ''):
                dropped_files.append(file_path)
        if len(dropped_files) > 0:
            waitWindow(window)
            conversionThread = threading.Thread(target=mainConversionThread, args=(window, dropped_files))
            conversionThread.start()
    # Bind the drop event to the left canvas
    leftCanvas.drop_target_register(DND_FILES) # type: ignore
    leftCanvas.dnd_bind('<<Drop>>', handle_drop) # type: ignore
    window.update()

# Function to run the waitWindow and calculation functions in a thread to prevent the GUI from freezing 
def mainConversionThread(window, file_list):
    try:
        fileDataframe = conversion.mainConversion(file_list)
        window.after(10, resultsWindow, window, fileDataframe)    
    except conversion.exception as e:
        errorWindow(window, e)

# Function to draw the results window
def waitWindow(window):
    # Reset the left canvas
    leftCanvas.delete("fileUploadImage", "fileUploadText")
    # Add the throbber background image
    throbberBackgroundImage = tk.PhotoImage(
        file=relativeToAssets("throbber_background.png")
    )
    leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=throbberBackgroundImage,
        tags="throbberBackgroundImage"
    )
    leftCanvas.throbberBackgroundImage = throbberBackgroundImage #type: ignore # Prevent garbage collection
    # Add the throbber image
    throbberImage = tk.PhotoImage(
        file=relativeToAssets("throbber.png")
    )
    throbberItem=leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=throbberImage,
        tags="throbberImage"
    )
    # Function to rotate the throbber image
    def rotate_image():
        for angle in range(2880, 0, -10):
            throbberImage = ImageTk.PhotoImage(
                Image.open(relativeToAssets("throbber.png")).rotate(angle, resample=Image.BICUBIC)
            )
            leftCanvas.itemconfig(throbberItem, image=throbberImage)
            leftCanvas.throbberImage = throbberImage  # type: ignore # Prevent garbage collection
            window.update()
            time.sleep(0.01)
    # Add the throbber text
    leftCanvas.create_text(
        130.0,
        340.0,
        anchor="center",
        text="Nalaganje...",
        fill="#333333",
        font=("Montserrat Medium", 12),
        tags="throbberText"
    )
    # Start the throbber rotation
    rotate_image()
    # Unregister the drop event
    leftCanvas.drop_target_unregister()
    window.update()

# Function to draw the error window
def resultsWindow(window, fileDataframe):
    # Reset the left and right canvas
    leftCanvas.delete("throbberImage", "throbberBackgroundImage", "throbberText")
    # Add the results window background image
    rightCanvas.delete("backgroundLogo")
    checkmarkImage = tk.PhotoImage(
        file=relativeToAssets("checkmark.png")
    )
    leftCanvas.create_image(
        130.0,
        240.0,
        anchor="center",
        image=checkmarkImage,
        tags="checkmarkImage"
    )
    leftCanvas.checkmarkImage = checkmarkImage # type: ignore # Prevent garbage collection
    # Add the results window text
    leftCanvas.create_text(
        130.0,
        340.0,
        anchor="center",
        text="Nalaganje končano!",
        fill="#333333",
        font=("Montserrat Medium", 12),
        tags="checkmarkText"
    )
    # Add the cancel button with image becauce TKinter doesn't support rounded buttons
    buttonCancelImage = Image.open(relativeToAssets("button_cancel.png"))
    buttonCancelImage = buttonCancelImage.resize((92, 40), resample=Image.BICUBIC)
    buttonCancelImage = ImageTk.PhotoImage(buttonCancelImage) 
    rightCanvas.button_cancel_image = buttonCancelImage # type: ignore # Prevent garbage collection
    buttonCancel = tk.Button(
        image=buttonCancelImage,
        borderwidth=0,
        highlightthickness=0,
        background="#FFFFFF",
        activebackground="#FFFFFF",
        command=lambda: resetWindow(window),
        relief="flat"
    )
    rightCanvas.create_window(350, 460, anchor="center", window=buttonCancel, width=97, height=45)
    # Function to open the excel file
    def openXlsxFile():
        loadSettingsFromFile()
        os.system(f'start excel "{excelOutputDir}/{excelLastFile}"')
    # Add the xlsx button with image becauce TKinter doesn't support rounded buttons
    buttonXlsxImage = Image.open(relativeToAssets("button_xlsx.png"))
    buttonXlsxImage = buttonXlsxImage.resize((120, 40), resample=Image.BICUBIC)
    buttonXlsxImage = ImageTk.PhotoImage(buttonXlsxImage)
    rightCanvas.button_xlsx_image = buttonXlsxImage # type: ignore # Prevent garbage collection
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
    # Add the treeview for results preview
    fileDataframe[fileDataframe.columns[0]] = fileDataframe[fileDataframe.columns[0]].apply(lambda x: x[:26] + '...' if len(x) > 26 else x)
    style = ttk.Style()
    style.configure("style.Treeview", borderwidth=0, highlightthickness=0, font=("Montserrat Medium", 10))
    style.configure("style.Treeview.Heading", font=("Montserrat SemiBold", 10))
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
    # Set the treeview columns
    tree.column(fileDataframe.columns[0], width=220, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[2], width=100, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[3], width=100, stretch=False) #type: ignore
    tree.column(fileDataframe.columns[-1], width=120, stretch=False) #type: ignore
    # Insert the treeview data
    for col in tree["columns"]:
        tree.heading(col, text=col)    
    for index, row in fileDataframe.iterrows():
        tree.insert("", "end", values=(row[fileDataframe.columns[0]], row[fileDataframe.columns[2]], row[fileDataframe.columns[3]], row[fileDataframe.columns[-1]])) #type: ignore
    window.update()

# Function to reset the window to the default state
def resetWindow(window):
    # Reset the left and right canvas
    leftCanvas.delete("checkmarkImage", "checkmarkText", "throbberText", "throbberImage", "throbberBackgroundImage")
    rightCanvas.delete("all")
    # Call the fileWindow function
    fileWindow(window)

# Function to draw the error window
def errorWindow(window, error):
    # Create and configure the error window
    errorWindow = tk.Toplevel(window)
    windowWidth = 500
    windowHeight = 350
    screenWidth = errorWindow.winfo_screenwidth()
    screenHeight = errorWindow.winfo_screenheight()
    x = (screenWidth // 2) - (windowWidth // 2)
    y = (screenHeight // 2) - (windowHeight // 2)
    errorWindow.title("Napaka")
    errorWindow.geometry(f"{windowWidth}x{windowHeight}+{x}+{y}")
    errorWindow.overrideredirect(True)
    errorWindow.resizable(False, False)
    errorWindow.iconbitmap(relativeToAssets("ikona.ico"))
    errorWindow.lift()
    errorWindow.attributes('-topmost',True)
    canvas = tk.Canvas(errorWindow, width=windowWidth, height=windowHeight, bg="#FFFFFF")
    canvas.pack()
    # Add the error window text and image
    errorMessage = tk.Label(canvas, text=error, font=("Montserrat Medium", 12), bg="#FFFFFF", fg="#333333")
    errorMessage.place(relx=0.5, rely=0.7, anchor="center")
    errorImage = tk.PhotoImage(file=relativeToAssets("error.png"))
    canvas.create_image(windowWidth // 2, windowHeight // 2 - 50, image=errorImage)
    canvas.errorImage = errorImage # type: ignore
    # Add the error window button with image becauce TKinter doesn't support rounded buttons
    buttonCancelImage = Image.open(relativeToAssets("button_again.png"))
    buttonCancelImage = buttonCancelImage.resize((160, 40), resample=Image.BICUBIC)
    buttonCancelImage = ImageTk.PhotoImage(buttonCancelImage)
    canvas.buttonCancelImage = buttonCancelImage # type: ignore # Prevent garbage collection
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
    buttonCancel.place(relx=0.5, rely=0.87, anchor="center")
    window.update()

# Function to draw the settings window
def openSettingsWindow(window):
    # Create and configure the settings window
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
    settingsWindow.iconbitmap(relativeToAssets("ikona.ico"))
    settingsWindow.lift()
    settingsWindow.attributes('-topmost',True)
    # Add the settings labels and inputs for setupPlanCollumns
    setupPlanCollumnsLabel = tk.Label(settingsWindow, text="Setup Plan Columns:", font=("Montserrat Medium", 10))
    setupPlanCollumnsLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    setupPlanCollumnsInput = tk.Entry(settingsWindow, width=50)
    setupPlanCollumnsInput.insert(0, ','.join(setupPlanCollumns))
    setupPlanCollumnsInput.pack(side="top", fill=tk.X, padx=20)
    # Add the settings labels and inputs for excelFileDirectory
    excelFileDirectoryLabel = tk.Label(settingsWindow, text="Excel File Directory:", font=("Montserrat Medium", 10))
    excelFileDirectoryLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelFileDirectoryInput = tk.Entry(settingsWindow, width=50)
    excelFileDirectoryInput.insert(0, str(excelFileDirectory))
    excelFileDirectoryInput.pack(side="top", fill=tk.X, padx=20)
    # Add the settings labels and inputs for excelSourceName
    excelSourceNameLabel = tk.Label(settingsWindow, text="Excel Source Name:", font=("Montserrat Medium", 10))
    excelSourceNameLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelSourceNameInput = tk.Entry(settingsWindow, width=50)
    excelSourceNameInput.insert(0, excelSourceName)
    excelSourceNameInput.pack(side="top", fill=tk.X, padx=20)
    # Add the settings labels and inputs for excelWorkbookName
    excelWorkbookNameLabel = tk.Label(settingsWindow, text="Excel Workbook Name:", font=("Montserrat Medium", 10))
    excelWorkbookNameLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelWorkbookNameInput = tk.Entry(settingsWindow, width=50)
    excelWorkbookNameInput.insert(0, excelWorkbookName)
    excelWorkbookNameInput.pack(side="top", fill=tk.X, padx=20)
    # Add the settings labels and inputs for excelOuputDir
    excelOuputDirLabel = tk.Label(settingsWindow, text="Excel Ouput Directory:", font=("Montserrat Medium", 10))
    excelOuputDirLabel.pack(side="top", fill=tk.X, padx=20, pady=10)
    excelOuputDirInput = tk.Entry(settingsWindow, width=50)
    excelOuputDirInput.insert(0, excelOutputDir)
    excelOuputDirInput.pack(side="top", fill=tk.X, padx=20)
    # Add the settings window buttons for close and save
    close_button = tk.Button(settingsWindow, text="Zapri", command=lambda:settingsWindow.destroy())
    close_button.place(relx=0.78, rely=0.94, anchor="center")
    save_button = tk.Button(settingsWindow, text="Shrani", command=lambda:saveSettings(setupPlanCollumnsInput, excelFileDirectoryInput, excelSourceNameInput, excelWorkbookNameInput, excelOuputDirInput))
    save_button.place(relx=0.9, rely=0.94, anchor="center")

# Main function to draw the main window
if __name__ == "__main__":
    mainWindow()