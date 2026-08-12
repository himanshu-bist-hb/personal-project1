import datetime
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)
import time
import os
import sqlite3
import numpy as np
import pandas as pd
import os.path
import xlwings as xw
import datetime
import tabulate
import xlsxwriter
from openpyxl.styles import Font
from openpyxl import load_workbook
from openpyxl import Workbook
from ttkthemes import ThemedStyle, ThemedTk

start_time = time.time()

pd.set_option('display.max_columns', None)
pd.options.display.width = None

import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog as fd


def callbackRatePages():
    exec(open("CP_Rate_Pages.py").read())

def callbackHexValidate():
    exec(open("CP_Hex_Validate.py").read())

def NGIC_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NGICRatebook
    NGICRatebook = fd.askopenfilename(
        title='Open a file',
        initialdir='M:/Actshare/Com/Annual_Rate_Reviews',
        filetypes=filetypes)
    label_file_explorerNGIC.configure(text=os.path.basename(NGICRatebook))

def MM_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global MMRatebook
    MMRatebook = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    label_file_explorerMM.configure(text=os.path.basename(MMRatebook))

def NACO_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NACORatebook
    NACORatebook = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    label_file_explorerNACO.configure(text=os.path.basename(NACORatebook))

def NAFF_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NAFFRatebook
    NAFFRatebook = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    label_file_explorerNAFF.configure(text=os.path.basename(NAFFRatebook))

def NICOF_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NICOFRatebook
    NICOFRatebook = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    label_file_explorerNICOF.configure(text=os.path.basename(NICOFRatebook))

def HICNJ_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global HICNJRatebook
    HICNJRatebook = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    label_file_explorerHICNJ.configure(text=os.path.basename(HICNJRatebook))

def TerrAdj_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global TerrAdjRatebook
    TerrAdjRatebook = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    label_file_explorerTerrAdj.configure(text=os.path.basename(TerrAdjRatebook))

def CW_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )
    global CWRatebook
    CWRatebook = fd.askopenfilename(
        title='Open a file',
        filetypes=filetypes)
    label_file_explorerCW.configure(text=os.path.basename(CWRatebook))

def getFolderPath():
    global folder_selected
    folder_selected = fd.askdirectory()
    folderPath.set(folder_selected)
    label_file_explorerSave.configure(text=folder_selected)


def hexValidate():
    def getFolderPathHex():
        global hexpath_selected
        hexpath_selected = fd.askdirectory()
        folderPath2.set(hexpath_selected)
        label_file_explorerSave2.configure(text=hexpath_selected)

    def Hexbook_File():
        filetypes = (
            ('All files', '*.*'),
            ('text files', '*.txt')
        )

        global HexbookFile
        HexbookFile = fd.askopenfilename(
            title='Open Hexbook',
            initialdir='M:/Actshare/Com/Property/Territory Defs (Independent)/3.0/Development/Hexbook Repository',
            filetypes=filetypes)
        label_file_explorerHex.configure(text=os.path.basename(HexbookFile))

    def SMPages_File():
        filetypes = (
            ('All files', '*.*'),
            ('text files', '*.txt')
        )

        global SMPagesFile
        SMPagesFile = fd.askopenfilename(
            title='Open Excel Rate Pages',
            filetypes=filetypes)
        label_file_explorerSMPages.configure(text=os.path.basename(SMPagesFile))

    def MMPages_File():
        filetypes = (
            ('All files', '*.*'),
            ('text files', '*.txt')
        )

        global MMPagesFile
        MMPagesFile = fd.askopenfilename(
            title='Open Excel Rate Pages',
            filetypes=filetypes)
        label_file_explorerMMPages.configure(text=os.path.basename(MMPagesFile))


    hexwindow = Toplevel(window)
    hexstyle = ThemedStyle(window)
    folderPath2 = StringVar()
    label_file_explorerSave2 = Label(hexwindow, text="", fg=fg, bg=bg)
    label_file_explorerHex = Label(hexwindow, text="", fg=fg, bg=bg)
    label_file_explorerSMPages = Label(hexwindow, text="", fg=fg, bg=bg)
    label_file_explorerMMPages = Label(hexwindow, text="", fg=fg, bg=bg)
    hexstyle.theme_use("black")
    hexwindow.configure(bg=hexstyle.lookup('TLabel','background'))
    hexwindow.title("Hex Factor Validation")
    # hexwindow.geometry("200x400")
    Label(hexwindow,text="Validate Hex Factors", fg=fg, bg=bg).grid(column=1,row=2,padx=25,pady=5)
    Button(hexwindow,text='Select Hexbook', fg=fg, command=Hexbook_File, bg=bg).grid(column=1,row=3,padx=25,pady=5)
    label_file_explorerHex.grid(column=1,row=4,pady=5)
    Button(hexwindow,text='Select SM Excel Page File',fg=fg,command=SMPages_File,bg=bg).grid(column=1,row=5,padx=25,pady=5)
    label_file_explorerSMPages.grid(column=1,row=6,pady=5)
    Button(hexwindow,text='Select MM Excel Page File',fg=fg,command=MMPages_File,bg=bg).grid(column=1,row=7,padx=25,pady=5)
    label_file_explorerMMPages.grid(column=1,row=8,pady=5)
    Button(hexwindow,text='Pick Save Location',fg=fg,command=getFolderPathHex,bg=bg).grid(column=1,row=9,padx=25,pady=5)
    label_file_explorerSave2.grid(column=1, row=10, pady=5)
    Button(hexwindow,text='Create Validation File',fg=fg,command=callbackHexValidate,bg=bg).grid(column=1,row=11,padx=5,pady=25)


window = Tk()

style = ThemedStyle(window)
style.theme_use("black")
bg = style.lookup('TLabel', 'background')
fg = style.lookup('TLabel', 'foreground')
window.configure(bg=bg)
# -----------------------------
# Scrollable container
# -----------------------------
main_container = Frame(window, bg=bg)
main_container.pack(fill="both", expand=True)

canvas = Canvas(main_container,bg=bg, highlightthickness=0)

v_scrollbar = ttk.Scrollbar(
    main_container,
    orient="vertical",
    command=canvas.yview
)

h_scrollbar = ttk.Scrollbar(
    main_container,
    orient="horizontal",
    command=canvas.xview
)

scrollable_frame = Frame(canvas,bg=bg)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window(
    (0, 0),
    window=scrollable_frame,
    anchor="nw"
)

canvas.configure(
    yscrollcommand=v_scrollbar.set,
    xscrollcommand=h_scrollbar.set
)

canvas.pack(side="left", fill="both", expand=True)

v_scrollbar.pack(side="right", fill="y")
h_scrollbar.pack(side="bottom", fill="x")

window.configure(bg=style.lookup('TLabel', 'background'))
# Make columns responsive
for col in range(11):
    window.grid_columnconfigure(col, weight=1)

# Make rows responsive
for row in range(20):
    window.grid_rowconfigure(row, weight=1)
#lb_tasks.configure(bg=bg, fg=fg)

window.title('CP Rate Page User Inputs')
window.update_idletasks()

try:
    window.state('zoomed')
except:
    w = window.winfo_screenwidth()
    h = window.winfo_screenheight()
    window.geometry(f"{w}x{h}+0+0")

label_file_explorerNGIC = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerNGIC.grid(column=1, row=3, padx=2, pady=5)

label_file_explorerMM = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerMM.grid(column=2, row=3, padx=2, pady=5)

label_file_explorerNACO = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerNACO.grid(column=1, row=6, padx=2, pady=5)

label_file_explorerNAFF = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerNAFF.grid(column=2, row=6, padx=2, pady=5)

label_file_explorerNICOF = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerNICOF.grid(column=1, row=9, padx=2, pady=5)

label_file_explorerHICNJ = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerHICNJ.grid(column=2, row=9, padx=2, pady=5)

label_file_explorerTerrAdj = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerTerrAdj.grid(column=1, row=12, pady=5)

label_file_explorerSave = Label(scrollable_frame, text="", fg=fg, bg=bg)
label_file_explorerSave.grid(column=6, row=11, pady=5)

SelectFileNGIC = Button(scrollable_frame, text='Select NGIC Ratebook', fg=fg, command=NGIC_Ratebook, bg=bg)
SelectFileNGIC.grid(column=1, row=2, pady=2)

SelectFileMM = Button(scrollable_frame, text='Select MM Ratebook', fg=fg, command=MM_Ratebook, bg=bg)
SelectFileMM.grid(column=2, row=2, pady=2)

SelectFileNACO = Button(scrollable_frame, text='Select NACO Ratebook', fg=fg, command=NACO_Ratebook, bg=bg)
SelectFileNACO.grid(column=1, row=5, pady=2)

SelectFileNAFF = Button(scrollable_frame, text='Select NAFF Ratebook', fg=fg, command=NAFF_Ratebook, bg=bg)
SelectFileNAFF.grid(column=2, row=5, pady=2)

SelectFileNICOF = Button(scrollable_frame, text='Select NICOF Ratebook', fg=fg, command=NICOF_Ratebook, bg=bg)
SelectFileNICOF.grid(column=1, row=8, pady=2)

SelectFileHICNJ = Button(scrollable_frame, text='Select HICNJ Ratebook', fg=fg, command=HICNJ_Ratebook, bg=bg)
SelectFileHICNJ.grid(column=2, row=8, pady=2)

SelectFileTerrAdj = Button(scrollable_frame, text='Select Hex Ratebook', fg=fg, command=TerrAdj_Ratebook,bg=bg)
SelectFileTerrAdj.grid(column=1,row=10,pady=2)

v1 = IntVar()
v2 = IntVar()
v3 = IntVar()
v4 = IntVar()
v5 = IntVar()
v6 = IntVar()
v14 = IntVar()
C1 = Checkbutton(scrollable_frame, text="Include LCM", variable=v1, fg=fg, bg=bg, selectcolor="grey")
C2 = Checkbutton(scrollable_frame, text="Include PNF", variable=v2, fg=fg, bg=bg, selectcolor="grey")
C3 = Checkbutton(scrollable_frame, text="Include Territory Adjustment", variable=v3, fg=fg, bg=bg, selectcolor ="grey")
C4 = Checkbutton(scrollable_frame, text="Include Capping Range", variable=v4, fg=fg, bg=bg, selectcolor ="grey")
C5 = Checkbutton(scrollable_frame, text="SM Rate Pages", variable=v5, fg=fg, bg=bg, selectcolor ="grey")
C6 = Checkbutton(scrollable_frame, text="Include Full Rate Manual", variable = v6, fg=fg, bg=bg, selectcolor = "grey")
C14 = Checkbutton(scrollable_frame, text="Include Deductible", variable = v14, fg=fg, bg=bg, selectcolor = "grey")
C1.grid(column=4, row=3, padx=2)
C2.grid(column=6, row=3, padx=2)
C3.grid(column=4, row=5, padx=2)
C4.grid(column=6, row=5, padx=2)
C5.grid(column=5, row=2, padx=2)
C6.grid(column=4, row=6, padx=2)
C14.grid(column=6, row=6, padx=2)

global SMMinCapRange
SMMinCapRange = DoubleVar()
concbox = Entry(scrollable_frame, justify=CENTER, textvariable=SMMinCapRange, fg=fg, bg=bg)
concbox.grid(column=5, row=11)

global SMMaxCapRange
SMMaxCapRange = DoubleVar()
concbox = Entry(scrollable_frame, justify=CENTER, textvariable=SMMaxCapRange, fg=fg, bg=bg)
concbox.grid(column=6, row=11)

global MMMinCapRange
MMMinCapRange = DoubleVar()
concbox = Entry(scrollable_frame, justify=CENTER, textvariable=MMMinCapRange, fg=fg, bg=bg)
concbox.grid(column=9, row=11)

global MMMaxCapRange
MMMaxCapRange = DoubleVar()
concbox = Entry(scrollable_frame, justify=CENTER, textvariable=MMMaxCapRange, fg=fg, bg=bg)
concbox.grid(column=10, row=11)

v7 = IntVar()
v8 = IntVar()
v9 = IntVar()
v10 = IntVar()
v11 = IntVar()
v12 = IntVar()
v13 = IntVar()
v15 = IntVar()
C7 = Checkbutton(scrollable_frame, text="Include LCM", variable=v7, fg=fg, bg=bg, selectcolor ="grey")
C8 = Checkbutton(scrollable_frame, text="Include PMF", variable=v8, fg=fg, bg=bg, selectcolor ="grey")
C9 = Checkbutton(scrollable_frame, text="Include Territory Adjustment", variable=v9, fg=fg, bg=bg, selectcolor ="grey")
C10 = Checkbutton(scrollable_frame, text="MM Rate Pages", variable=v10, fg=fg, bg=bg, selectcolor ="grey")
C11 = Checkbutton(scrollable_frame, text="Include Capping Range", variable=v11, fg=fg, bg=bg, selectcolor ="grey")
C12 = Checkbutton(scrollable_frame, text="Include Segmentation Removal", variable=v12,fg=fg,bg=bg, selectcolor="grey")
C13 = Checkbutton(scrollable_frame, text="Include Full Rate Manual", variable=v13,fg=fg,bg=bg, selectcolor="grey")
C15 = Checkbutton(scrollable_frame, text="Include Deductible", variable=v15,fg=fg,bg=bg, selectcolor="grey")
C7.grid(column=8, row=3, padx=2)
C8.grid(column=10, row=3, padx=2)
C9.grid(column=8, row=5, padx=2)
C10.grid(column=9, row=2, padx=2)
C11.grid(column=10, row=5, padx=2)
C12.grid(column=8,row=6,padx=2)
C13.grid(column=10,row=6,padx=2)
C15.grid(column=8, row=7, padx=2)

folderPath = StringVar()

SaveButton = Button(scrollable_frame, text="Select Save Location", fg=fg, command=getFolderPath, bg=bg)
SaveButton.grid(column=5, row=14, pady=5)
label_file_explorerSave.grid(column=5, row=15, pady=2)


btncancel = Button(scrollable_frame, text="Cancel", width=10, fg=fg, command=window.destroy, bg=bg)
btncancel.grid(column=8, row=14, padx=2, pady=2)

btnrunRatePages = Button(scrollable_frame, text="Create Rate Pages", width=20, fg=fg, command=callbackRatePages, bg=bg)
btnrunRatePages.grid(column=6, row=14, padx=2, pady=2)

btnrunHexValidate = Button(scrollable_frame, text="Validate Hex Factors", width=20, fg=fg, command=hexValidate)
btnrunHexValidate.grid(column=7, row=14, padx=2,pady=2)

lbl1 = Label(scrollable_frame, text="Select items to include in small market output", fg='#65C7E3', font=("Arial", 12), bg=bg)
lbl1.grid(column=4, row=1, pady=5, columnspan=3)

lbl2 = Label(scrollable_frame, text="Select items to include in middle market output", fg='#65C7E3', font=("Arial", 12), bg=bg)
lbl2.grid(column=8, row=1, pady=5, columnspan=3)

lbl3 = Label(scrollable_frame, text="Select Proposed Ratebooks", fg='#65C7E3', font=("Arial", 12), bg=bg)
lbl3.grid(column=1, row=1, pady=5, columnspan=2)

lbl4 = Label(scrollable_frame, text="Input Min Cap Range:", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl4.grid(column=5, row=10)

lbl5 = Label(scrollable_frame, text="Input Max Cap Range:", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl5.grid(column=6, row=10)

lbl4 = Label(scrollable_frame, text="Input Min Cap Range:", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl4.grid(column=9, row=10)

lbl5 = Label(scrollable_frame, text="Input Max Cap Range:", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl5.grid(column=10, row=10)

lbl2 = Label(scrollable_frame, text="Input Decimal i.e. -0.5", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl2.grid(column=5, row=12)

lbl2 = Label(scrollable_frame, text="Input Decimal i.e. 0.5", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl2.grid(column=6, row=12)

lbl2 = Label(scrollable_frame, text="Input Decimal i.e. -0.5", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl2.grid(column=9, row=12)

lbl2 = Label(scrollable_frame, text="Input Decimal i.e. 0.5", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl2.grid(column=10, row=12)


tk.ttk.Separator(scrollable_frame, orient=VERTICAL).grid(column=3, row=0, padx=15, rowspan=10, sticky='ns')
tk.ttk.Separator(scrollable_frame, orient=VERTICAL).grid(column=7, row=0, padx=15, rowspan=10, sticky='ns')
window.mainloop()