import datetime
import xlsxwriter

from tkinter import Tk
from tkinter.filedialog import askopenfilename

from dbs import dbs
from ocbc import ocbc
from utils import *


def main():
    Tk().withdraw()
    filename = askopenfilename(filetypes=[("PDF Files", "*.pdf")])

    if filename == '':
        print('No file selected! Program Terminating...')
        quit(0)

    t = input('Enter type of document (0 - OCBC, 1 - DBS): ')
    starttime = datetime.datetime.now()
    print(f"Process Started {starttime}.")

    sheetname = filename.replace("pdf", "xlsx")
    print(f'Creating Output Excel File ({sheetname})...', end=" ")
    workbook = xlsxwriter.Workbook(sheetname)

    sheetname = sheetname.split('/')[-1][:-5]
    if len(sheetname) > 30:
        sheetname = sheetname[:30]
    worksheet = workbook.add_worksheet(sheetname)

    for i in range(len(SHEETTOP)):
        worksheet.write(0, i, SHEETTOP[i][0])
        worksheet.set_column(i, i, SHEETTOP[i][1])

    print('Done!')
    print(f'Reading {filename}...')
    [ocbc, dbs][int(t)](filename, worksheet)

    print('Done! Closing and Saving...')
    workbook.close()

    endtime = datetime.datetime.now()
    print(f'Process Ended {endtime} ({(endtime - starttime).seconds}s)')

