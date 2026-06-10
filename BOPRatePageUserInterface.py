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
from ttkthemes import ThemedStyle

initialTime = time.perf_counter() # starting the clock

pd.set_option('display.max_columns', None)
pd.options.display.width = None

import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog as fd
#class BOPUserInterface:

    #def _init_(self, state, rateTables, nEffective, rEffective, NGICRatebook, NAFFRatebook, NACORatebook, NICOFRatebook, RatingPlansApplies, CommonRulesApplies, AdditionalRulesApplies, OptionalCoveragesApplies, IRPMCredit, IRPMDebit, folderPath) -> None:
     #   self.state = state
     #   self.rateTables = rateTables
     #   self.nEffective = nEffective
     #   self.rEffective = rEffective
     #   self.NGICRatebook = NGICRatebook
     #   self.NAFFRatebook = NAFFRatebook
     #   self.NACORatebook = NACORatebook
     #   self.NICOFRatebook = NICOFRatebook
     #   self.RatingPlansApplies = RatingPlansApplies
     #   self.CommonRulesApplies = CommonRulesApplies
     #   self.AdditionalRulesApplies = AdditionalRulesApplies
     #   self.OptionalCoveragesApplies = OptionalCoveragesApplies
     #   self.IRPMCredit = IRPMCredit
     #   self.IRPMDebit = IRPMDebit
     #   self.folderPath = folderPath

def callbackBOPRatePages():
    import tkinter
    proceed = tkinter.messagebox.askyesno(
        "",
        "This is the BOP-Apt-2.0 Branch.\n\nDo you want to proceed?"
    )
    if proceed:
        exec(open("StatePageGenerator.py").read())

def callbackBOPCurrentRatePages():
    exec(open("StatePageGeneratorCurrent.py").read())

def NGIC_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NGICRatebook
    NGICRatebook = fd.askopenfilename(
        title='Select Proposed NGIC Ratebook',
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
        title='Select Proposed MM Ratebook',
        #initialdir='M:/Actshare/Com/Annual_Rate_Reviews',
        filetypes=filetypes)
    label_file_explorerMM.configure(text=os.path.basename(MMRatebook))

def NACO_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NACORatebook
    NACORatebook = fd.askopenfilename(
        title='Select Proposed NACO Ratebook',
        #initialdir='M:/Actshare/Com/Annual_Rate_Reviews',
        filetypes=filetypes)
    label_file_explorerNACO.configure(text=os.path.basename(NACORatebook))

def NAFF_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NAFFRatebook
    NAFFRatebook = fd.askopenfilename(
        title='Select Proposed NAFF Ratebook',
        #initialdir='M:/Actshare/Com/Annual_Rate_Reviews',
        filetypes=filetypes)
    label_file_explorerNAFF.configure(text=os.path.basename(NAFFRatebook))

def NICOF_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global NICOFRatebook
    NICOFRatebook = fd.askopenfilename(
        title='Select Proposed NICOF Ratebook',
        #initialdir='M:/Actshare/Com/Annual_Rate_Reviews',
        filetypes=filetypes)
    label_file_explorerNICOF.configure(text=os.path.basename(NICOFRatebook))

def HICNJ_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global HICNJRatebook
    HICNJRatebook = fd.askopenfilename(
        title='Select Proposed HICNJ Ratebook',
        #initialdir='M:/Actshare/Com/Annual_Rate_Reviews',
        filetypes=filetypes)
    label_file_explorerHICNJ.configure(text=os.path.basename(HICNJRatebook))


def CW_Ratebook():
    filetypes = (
        ('All files', '*.*'),
        ('text files', '*.txt')
    )

    global CWRatebook
    CWRatebook = fd.askopenfilename(
        title='Select CW Ratebook',
        #initialdir='M:/Actshare/Com/Annual_Rate_Reviews',
        filetypes=filetypes)
    label_file_explorerCW.configure(text=os.path.basename(CWRatebook))

def getFolderPath():
    global folder_selected
    folder_selected = fd.askdirectory()
    folderPath.set(folder_selected)
    label_file_explorerSave.configure(text=folder_selected)


#def LaunchUserInterface():
window = Toplevel()
style = ThemedStyle(window)
style.theme_use("black")
bg = style.lookup('TLabel', 'background')
fg = style.lookup('TLabel', 'foreground')
window.configure(bg=style.lookup('TLabel', 'background'))
#lb_tasks.configure(bg=bg, fg=fg)

window.title('BOP Rate Page User Inputs')
window.geometry("+400+400")


label_file_explorerNGIC = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerNGIC.grid(column=1, row=3, padx=2, pady=5)

label_file_explorerMM = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerMM.grid(column=3, row=3, padx=2, pady=5)

label_file_explorerNACO = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerNACO.grid(column=1, row=7, padx=2, pady=5)

label_file_explorerNAFF = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerNAFF.grid(column=3, row=7, padx=2, pady=5)

label_file_explorerNICOF = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerNICOF.grid(column=1, row=10, padx=2, pady=5)

label_file_explorerHICNJ = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerHICNJ.grid(column=3, row=10, padx=2, pady=5)

label_file_explorerCW = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerCW.grid(column=2, row=13, padx=2, pady=5)

label_file_explorerSave = Label(window, text="", fg='#65C7E3', bg=bg)
label_file_explorerSave.grid(column=2, row=22, pady=5)


SelectFileNGIC = Button(window, text='Select NGIC Ratebook', fg=fg, command=NGIC_Ratebook, bg=bg)
SelectFileNGIC.grid(column=1, row=2, pady=2)

SelectFileMM = Button(window, text='Select MM Ratebook', fg=fg, command=MM_Ratebook, bg=bg)
SelectFileMM.grid(column=3, row=2, pady=2)

SelectFileNACO = Button(window, text='Select NACO Ratebook', fg=fg, command=NACO_Ratebook, bg=bg)
SelectFileNACO.grid(column=1, row=6, pady=2)

SelectFileNAFF = Button(window, text='Select NAFF Ratebook', fg=fg, command=NAFF_Ratebook, bg=bg)
SelectFileNAFF.grid(column=3, row=6, pady=2)

SelectFileNICOF = Button(window, text='Select NICOF Ratebook', fg=fg, command=NICOF_Ratebook, bg=bg)
SelectFileNICOF.grid(column=1, row=9, pady=2)

SelectFileHICNJ = Button(window, text='Select HICNJ Ratebook', fg=fg, command=HICNJ_Ratebook, bg=bg)
SelectFileHICNJ.grid(column=3, row=9, pady=2)

SelectFileCW = Button(window, text='Select CW Ratebook', fg=fg, command=CW_Ratebook, bg=bg)
SelectFileCW.grid(column=2, row=12, pady=2)


RatingPlansApplies = IntVar()
CommonRulesApplies = IntVar()
AdditionalRulesApplies = IntVar()
OptionalCoveragesApplies = IntVar()
ClassApplies = IntVar()
IndProgramsApplies = IntVar(value=1) #on by default
AllPerilApplies = IntVar()
AllProgramApplies = IntVar()

#v5 = IntVar()
C1 = Checkbutton(window, text="Include Rating Plans", variable=RatingPlansApplies, fg=fg, bg=bg, selectcolor ="grey")
C2 = Checkbutton(window, text="Include Common Rules", variable=CommonRulesApplies, fg=fg, bg=bg, selectcolor ="grey")
C3 = Checkbutton(window, text="Include Additional Rules", variable=AdditionalRulesApplies, fg=fg, bg=bg, selectcolor ="grey")
C4 = Checkbutton(window, text="Include Optional Coverages", variable=OptionalCoveragesApplies, fg=fg, bg=bg, selectcolor ="grey")
C5 = Checkbutton(window, text="Include Class", variable=ClassApplies, fg=fg, bg=bg, selectcolor ="grey")
C6 = Checkbutton(window, text="Include All Peril", variable=AllPerilApplies, fg=fg, bg=bg, selectcolor ="grey")
C7 = Checkbutton(window, text="Include All Programs", variable=AllProgramApplies, fg=fg, bg=bg, selectcolor ="grey")
C8 = Checkbutton(window, text="Include Individual Programs", variable=IndProgramsApplies, fg=fg, bg=bg, selectcolor ="grey")

#C5 = Checkbutton(window, text="SM Rate Pages", variable=v5, fg=fg, bg=bg, selectcolor ="grey")
C1.grid(column=1, row=15, padx=2, sticky='w')
C2.grid(column=1, row=17, padx=2, sticky='w')
C3.grid(column=1, row=18, padx=2, sticky='w')
C4.grid(column=1, row=19, padx=2, sticky='w')
C5.grid(column=1, row=20, padx=2, sticky='w')
C6.grid(column=1, row=21, padx=2, sticky='w')
C7.grid(column=1, row=22, padx=2, sticky='w')
C8.grid(column=1, row=23, padx=2, sticky='w')
#C5.grid(column=1, row=2, padx=2)


global IRPMCredit
IRPMCredit = DoubleVar()
concbox = Entry(window, justify=CENTER, textvariable=IRPMCredit, fg=fg, bg=bg)
concbox.grid(column=3, row=15)
#IRPMCredit = float(IRPMCredit)

global IRPMDebit
IRPMDebit = DoubleVar()
concbox = Entry(window, justify=CENTER, textvariable=IRPMDebit, fg=fg, bg=bg)
concbox.grid(column=3, row=16)
#IRPMDebit = float(IRPMDebit)

folderPath = StringVar()
SaveButton = Button(window, text="Select Save Location", fg=fg, command=getFolderPath, bg=bg)
SaveButton.grid(column=2, row=25, pady=5)
label_file_explorerSave.grid(column=2, row=26, pady=2)

btncancel = Button(window, text="Cancel", width=10, fg=fg, command=window.destroy, bg=bg)
btncancel.grid(column=3, row=28, padx=2, pady=5)

btnrunRatePages = Button(window, text="Create BP2.0 Rate Pages", width=25, fg=fg, command=callbackBOPRatePages, bg=bg)
btnrunRatePages.grid(column=1, row=28, padx=2, pady=5)

btnrunRatePages = Button(window, text="Create Pre 2.0 Rate Pages", width=25, fg=fg, command=callbackBOPCurrentRatePages, bg=bg)
btnrunRatePages.grid(column=2, row=28, padx=2, pady=5)


lbl1 = Label(window, text="Select Proposed Ratebooks",
            fg='#65C7E3', font=("Arial", 12), bg=bg)
lbl1.grid(column=1, row=1, pady=5, columnspan=3)

#lbl2 = Label(window, text="Misc Items",
#             fg='#65C7E3', font=("Arial", 12), bg=bg)
#lbl2.grid(column=1, row=15, pady=5, columnspan=3)

lbl4 = Label(window, text="Input IRPM Credit Percentage:", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl4.grid(column=2, row=15)

lbl5 = Label(window, text="Input IRPM Debit Percentage:", fg='#65C7E3', font=("Arial", 10), bg=bg)
lbl5.grid(column=2, row=16)

tk.ttk.Separator(window, orient=HORIZONTAL).grid(column=0, row=14, pady=10, columnspan=4, sticky='ew')
tk.ttk.Separator(window, orient=HORIZONTAL).grid(column=0, row=24, pady=10, columnspan=4, sticky='ew')
tk.ttk.Separator(window, orient=HORIZONTAL).grid(column=0, row=27, pady=10, columnspan=4, sticky='ew')


window.mainloop()


endTime = time.perf_counter() # stopping the clock
print(f'This program ran in {endTime - initialTime:0.4f} seconds')

        #LaunchUserInterface()

