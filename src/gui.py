import tkinter as tk
import tkinter.font as font
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
from PIL import Image, ImageTk
from PIL import Image, ImageTk, ImageFilter
import ctypes
import math

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("../assets")
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
        file_paths = event.data.split('\n')
        file_paths = [path.strip() for path in file_paths if path.strip()]
        for file_path in file_paths:
            dropped_files.append(file_path)
        if len(dropped_files) > 0:
            waitWindow()
    
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
        window.after(20, rotate_image, angle - 10)
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
    
def resultsWindow():
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
    window.update()

if __name__ == "__main__":
    mainWindow()