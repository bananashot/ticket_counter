import tkinter
from tkinter import *

import keyboard

from sheet_logic import *

if __name__ == "__main__":
    keyboard.add_hotkey('alt+s', lambda: add_ticket())
    keyboard.add_hotkey('alt+c', lambda: get_today_ticket_count())
    keyboard.add_hotkey('alt+o', lambda: open_sheet_copy())

    root = tkinter.Tk()
    root.resizable(False, False)
    root.title("Ticket counter")
    root.geometry("500x200")

    canvas = Canvas(root, bg="#444444", height=500, width=800,
                    bd=0, highlightthickness=0, relief="ridge")
    canvas.place(x=0, y=0)

    texts = """
    Instructions:

    Add ticket - Alt + S
    Get ticket count - Alt + C
    Open sheet copy - Alt + O
    """
    header = canvas.create_text(250, 100, text=texts, fill="#ECE8EF", font=(
        "Terminal", int(20.0)), justify=CENTER)
    canvas.pack()

    root.mainloop()
