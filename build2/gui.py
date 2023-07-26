from pathlib import Path
import tkinter as tk
import tkinter.font as font
import ctypes

ctypes.windll.shcore.SetProcessDpiAwareness(1)


OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

def close_window(window):
    window.destroy()

window = tk.Tk()
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
    file=relative_to_assets("image_1.png"))
canvas.create_image(
    527.4769287109375,
    249.0,
    image=image_1
)

canvas.create_rectangle(
    0.0,
    0.0,
    260.0,
    500.0,
    fill="#FFFFFF",
    outline="")

canvas.create_text(
    86.0,
    354.0,
    anchor="nw",
    text="Nalaganje...",
    fill="#333333",
    font=("Montserrat Medium", 16 * -1)
)

canvas.create_rectangle(
    260.0,
    0.0,
    262.0,
    500.0,
    fill="#D9D9D9",
    outline="")

image_2 = tk.PhotoImage(
    file=relative_to_assets("image_2.png"))
canvas.create_image(
    129.0,
    249.99996948242188,
    image=image_2
)

import math

from PIL import Image, ImageTk

from PIL import Image, ImageTk, ImageFilter

def rotate_image(angle=0):
    image_3_rotated = ImageTk.PhotoImage(
        Image.open(relative_to_assets("throbber.png")).rotate(angle, resample=Image.BICUBIC)
    )
    canvas.itemconfig(image_item, image=image_3_rotated)
    canvas.image = image_3_rotated  # type: ignore # prevent garbage collection
    window.after(20, rotate_image, angle - 10)


image_3 = tk.PhotoImage(
    file=relative_to_assets("throbber.png")
)

image_item = canvas.create_image(
    130.0,
    250.0,
    image=image_3
)

rotate_image()

window.resizable(False, False)
window.mainloop()