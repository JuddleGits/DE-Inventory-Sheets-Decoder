import os
import tkinter as tk
from GSheetsApp import SheetsApp

root = tk.Tk()
root.geometry("300x200")
root.title("DE Inventory Decoder")
app = SheetsApp(master=root)
app.mainloop()
