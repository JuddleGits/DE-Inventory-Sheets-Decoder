import os
import tkinter as tk
from GSheetsApp import SheetsApp

root = tk.Tk()
root.geometry("800x600")
app = SheetsApp(master=root)
app.mainloop()