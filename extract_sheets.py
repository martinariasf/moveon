print("Python is running with openpyxl")
import openpyxl
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import messagebox


# create the root window
messagebox.showinfo("Importing files", "First select the MoveOn file [AcademicMoves (date)] and then the Ranking sheet [KWYEAR_WEEK_MASTER_TYP]")

def moveon_file():
    filetypes = (
        ('Excel files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='MOVEON file',
        initialdir='W:\Faculty-GS\Office-K\Ablage',
        filetypes=filetypes)

    showinfo(
        title='Selected MOVEONE File. Now select the Ranking file',
        message=filename[-36:]
    )
    return filename

moveon_file_dir=moveon_file()
moveon=openpyxl.load_workbook(moveon_file_dir)

def rank_file():
    filetypes = (
        ('Excel files', '*.xlsx;*.xlsm'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='RANKING file',
        initialdir="W:\Faculty-GS\Office-K\Ablage",
        filetypes=filetypes)

    showinfo(
        title='Selected RANKING File. The new excel will be created in 5 seconds',
        message=filename[-22:]
    )
    return filename

rank_file_dir=rank_file()
rank=openpyxl.load_workbook(rank_file_dir)
