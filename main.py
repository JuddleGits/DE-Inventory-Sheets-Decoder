import os
import tkinter as tk
from GSheetsApp import SheetsApp

root = tk.Tk()
root.geometry("400x400")
root.title("DE Inventory Decoder")
root.iconphoto(False, tk.PhotoImage(file="DEIcon.png"))
app = SheetsApp(master=root)
app.mainloop()