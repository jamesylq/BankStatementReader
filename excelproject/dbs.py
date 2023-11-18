from excelproject.utils import *

from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams, LTTextBox
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter


def dbs(filename, worksheet):
    fp = open(filename, 'rb')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pages = list(PDFPage.get_pages(fp))

    sheety = 1
    for n in range(len(pages)):
        curr = []
        print(f'Processing Page {n + 1} of {len(pages)}...')
        interpreter.process_page(pages[n])
        layout = device.get_result()
        prev = 0
        prevL = None
        info = []
        for lobj in layout:
            if isinstance(lobj, LTTextBox):
                x, y, text = lobj.bbox[0], lobj.bbox[3], lobj.get_text()

                if x == 173.0 and text not in DBSIGNORE:
                    count = 0
                    for row in rws(text, DBSIGNORE).split('\n'):
                        if row != '':
                            info.append(row)
                            count += 1
                    if 0 < count < 3:
                        for i in range(3 - count):
                            info.append('')

                if x == 62.0:
                    d = Entry(round(y, 5))
                    d.transDate = rws(text)

                    if prevL is None and len(d.transDate) == 9:
                        d.l = 0
                        prevL = 0
                        prev = d.y

                    if prevL is not None:
                        d.l = round((prev - d.y) / SPACE + prevL)
                        prevL = d.l
                        prev = d.y

                    curr.append(d)

                elif x == 128.0:
                    d = find(curr, round(y, 5))
                    if d is not None:
                        d.valueDate = rws(text)

                elif 360 <= x <= 410:
                    d = find(curr, round(y, 5))
                    if d is not None:
                        d.withdrawal = float(rws(text))

                elif 440 <= x <= 490:
                    d = find(curr, round(y, 5))
                    if d is not None:
                        d.deposit = float(rws(text))

                elif 520 <= x <= 600:
                    d = find(curr, round(y, 5))
                    if d is not None:
                        d.balance = float(rws(text))

            if curr and info is not None:
                d = curr[0]

                for i in range(len(curr) - 1):
                    n = curr[i + 1]
                    if n.l is None:
                        continue
                    d.description = " ".join(info[d.l:n.l])
                    d = n

                curr[-1].description = " ".join(info[d.l:])

        for d in curr:
            if d.countNone() == 0:
                worksheet.write(sheety, 0, sheety)
                worksheet.write(sheety, 1, d.transDate)
                worksheet.write(sheety, 2, d.valueDate)
                worksheet.write(sheety, 3, d.description)
                worksheet.write(sheety, 4, d.withdrawal)
                worksheet.write(sheety, 5, d.deposit)
                worksheet.write(sheety, 6, d.balance)
                sheety += 1
