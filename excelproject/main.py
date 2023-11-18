import datetime
import xlsxwriter

from tkinter import Tk
from tkinter.filedialog import askopenfilename

from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams, LTTextBox
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter


SPACE = 11.3375
SHEETTOP = [
    ("Index", 5),
    ("Transaction Date", 14),
    ("Value Date", 14),
    ("Description", 90),
    ("Withdrawal", 10),
    ("Deposit", 10),
    ("Balance", 10)
]


def find(l, y):
    for x in l:
        if x.y == y:
            return x


def rws(s: str, char: str = None):
    if char is None:
        char = '\n\t ,'
    for c in char:
        s = s.replace(c, '')
    return s


def main():
    data = []

    Tk().withdraw()
    filename = askopenfilename(filetypes=[("PDF Files", "*.pdf")])

    if filename == '':
        print('No file selected! Program Terminating...')

    fp = open(filename, 'rb')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pages = list(PDFPage.get_pages(fp))
    sheety = 1
    starttime = datetime.datetime.now()

    print(f"Process Started {starttime}.")

    for n in range(len(pages)):
        curr = []
        data.append([])
        print(f'Processing Page {n + 1} of {len(pages)}...')
        interpreter.process_page(pages[n])
        layout = device.get_result()
        prev = 0
        prevL = None
        info = None
        for lobj in layout:
            if isinstance(lobj, LTTextBox):
                x, y, text = lobj.bbox[0], lobj.bbox[3], lobj.get_text()

                if x == 136.92 and 'BALANCE' not in text:
                    info = text.split('\n')

                curr.append((x, y, text))

                if x == 46.2:
                    d = Entry(round(y, 5))
                    d.transDate = rws(text)

                    if prevL is None and len(d.transDate) == 5:
                        d.l = 0
                        prevL = 0
                        prev = d.y

                    if prevL is not None:
                        d.l = round((prev - d.y) / SPACE + prevL)
                        prevL = d.l
                        prev = d.y

                    data[-1].append(d)

                elif x == 91.56:
                    d = find(data[-1], round(y, 5))
                    if d is not None:
                        d.valueDate = rws(text)

                elif 300 <= x <= 350:
                    d = find(data[-1], round(y, 5))
                    if d is not None:
                        d.withdrawal = float(rws(text))

                elif 390 <= x <= 450:
                    d = find(data[-1], round(y, 5))
                    if d is not None:
                        d.deposit = float(rws(text))

                elif 490 <= x <= 550:
                    d = find(data[-1], round(y, 5))
                    if d is not None:
                        d.balance = float(rws(text))

        if data[-1] and info is not None:
            d = data[-1][0]

            for i in range(len(data[-1]) - 1):
                n = data[-1][i + 1]
                if n.l is None:
                    continue
                d.description = " ".join(info[d.l:n.l])
                d = n

            data[-1][-1].description = " ".join(info[d.l:])

    filename = filename.replace("pdf", "xlsx")
    print(f'Building Output Excel File ({filename})...')
    workbook = xlsxwriter.Workbook(filename)
    filename = filename.split('/')[-1][:-5]
    worksheet = workbook.add_worksheet(filename)

    for i in range(len(SHEETTOP)):
        worksheet.write(0, i, SHEETTOP[i][0])
        worksheet.set_column(i, i, SHEETTOP[i][1])

    for i in range(len(data)):
        for d in data[i]:
            if d.countNone() == 0:
                worksheet.write(sheety, 0, sheety)
                worksheet.write(sheety, 1, d.transDate)
                worksheet.write(sheety, 2, d.valueDate)
                worksheet.write(sheety, 3, d.description)
                worksheet.write(sheety, 4, d.withdrawal)
                worksheet.write(sheety, 5, d.deposit)
                worksheet.write(sheety, 6, d.balance)
                sheety += 1

    workbook.close()

    endtime = datetime.datetime.now()
    print(f'Process Ended {endtime} ({(endtime - starttime).seconds}s)')


class Entry:
    def __init__(self, y):
        self.transDate = None
        self.valueDate = None
        self.description = None
        self.withdrawal = None
        self.deposit = None
        self.balance = None
        self.y = y
        self.l = None

    def countNone(self):
        c = 0
        for x in [self.transDate, self.valueDate, self.description, self.balance, self.l]:
            if x is None:
                c += 1
        return c
