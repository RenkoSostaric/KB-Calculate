from pathlib import Path
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter.font as font
import ctypes

ctypes.windll.shcore.SetProcessDpiAwareness(1)

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")

dropped_files = []

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

def close_window(window):
    window.destroy()
    
# Define a function to handle the drop event
def handle_drop(event):
    # Get the list of file paths from the drop event
    file_paths = event.data.split('\n')
    file_paths = [path.strip() for path in file_paths if path.strip()]

    # Do something with the file paths
    for file_path in file_paths:
        dropped_files.append(file_path)
        
def main_gui():
    window = TkinterDnD.Tk()
    window.title("KB Calculate")
    window_width = 800
    window_height = 530
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    window.overrideredirect(True)

    nav_bar = tk.Frame(
        window,
        bg = "#EAEAEA",
        height = 30,
        width = 800,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
    )

    nav_bar.place(x = 0, y = 0)

    
    x_font = font.Font(family="FontAwesome", size=14)
    x_button = tk.Button(
        nav_bar,
        text = "\uf00d",
        font = x_font,
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        command = lambda: close_window(window)
    )

    x_button.place(x = 770, y = 0)
    
    title_font = font.Font(family="Montserrat Medium", size=12)
    title = tk.Label(
        nav_bar,
        text = "KB Calculate",
        bg = "#EAEAEA",
        fg = "#333333",
        bd = 0,
        highlightthickness = 0,
        font = title_font
    )
    title.place(relx = 0.5, rely = 0.5, anchor = "center")

    canvas = tk.Canvas(
        window,
        bg = "#FFFFFF",
        height = 500,
        width = 800,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
    )

    canvas.place(x = 0, y = 30)

    image_1 = tk.PhotoImage(
        file=relative_to_assets("image_1.png")
    )

    canvas.create_image(
        527.4769287109375,
        249.0,
        image=image_1
    )

    canvas.create_rectangle(
        260.0,
        0.0,
        262.0,
        500.0,
        fill="#D9D9D9",
        outline=""
    )

    drop_zone = tk.Canvas(
        window,
        bg = "#FFFFFF",
        height = 500,
        width = 260,
        bd = 0,
        highlightthickness = 0,
        relief = "ridge"
    )

    drop_zone.place(x = 0, y = 30)

    image_2 = tk.PhotoImage(
        file=relative_to_assets("image_2.png")
    )

    drop_zone.create_image(
        130.0,
        250.0,
        image=image_2,
        tags = "image_2"
    )

    drop_zone.create_text(
        26.0,
        354.0,
        anchor="nw",
        text="Povleci in spusti datoteke",
        fill="#333333",
        font=("Montserrat Medium", 16 * -1)
    )

    drop_zone.drop_target_register(DND_FILES) # type: ignore
    drop_zone.dnd_bind('<<Drop>>', handle_drop) # type: ignore

    window.resizable(False, False)
    window.mainloop()
    
if __name__ == "__main__":
    main_gui()