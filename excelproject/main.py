import os
import sys
import glob
import pygame
import datetime
import xlsxwriter

from tkinter import Tk
from tkinter.filedialog import askopenfilename

try:
    from excelproject.utils import *
except ModuleNotFoundError:
    from utils import *

from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams, LTTextBox
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter


# class Button:
#     def __init__(self, x: Union[int, float], y: Union[int, float], w: Union[int, float], h: Union[int, float],
#                  color: Tuple[int, int, int], *, border: Union[None, Tuple[int, int, int]]):
#         self.x = x
#         self.y = y
#         self.w = w
#         self.h = h
#         self.rect = (x, y, w, h)
#         self.color = color
#         self.border = bool(border)
#         self.borderColor = border
#
#         global buttons
#         buttons.append(self)
#
#     def __del__(self):
#         buttons.remove(self)
#
#     def update(self):
#         global screen
#         if self.border:
#             pygame.draw.rect(screen, self.borderColor, self.rect, 3)
#         pygame.draw.rect(screen, self.color, self.rect)


# buttons: List[Button] = []
screen = pygame.display.set_mode((1200, 700))
pygame.init()
pygame.font.init()
clock = pygame.time.Clock()

curr_path = os.path.dirname(__file__)
resource_path = os.path.join(curr_path, 'resources')

tinyFont = pygame.font.Font(os.path.join(resource_path, 'fonts', 'UbuntuMono-Regular.ttf'), 17)
font = pygame.font.Font(os.path.join(resource_path, 'fonts', 'UbuntuMono-Regular.ttf'), 20)
mediumFont = pygame.font.Font(os.path.join(resource_path, 'fonts', 'UbuntuMono-Regular.ttf'), 30)
largeFont = pygame.font.Font(os.path.join(resource_path, 'fonts', 'UbuntuMono-Regular.ttf'), 75)


def leftAlignPrint(f: pygame.font.Font, text: str, pos: Tuple[int, int], color: Tuple[int, int, int] = (0, 0, 0)) -> None:
    textObj = f.render(text, True, color)
    screen.blit(textObj, textObj.get_rect(center=[pos[0] + f.size(text)[0] / 2, pos[1]]))


def centredPrint(f: pygame.font.Font, text: str, pos: Tuple[int, int], color: Tuple[int, int, int] = (0, 0, 0)) -> None:
    textObj = f.render(text, True, color)
    screen.blit(textObj, textObj.get_rect(center=pos))


def rightAlignPrint(f: pygame.font.Font, text: str, pos: Tuple[int, int], color: Tuple[int, int, int] = (0, 0, 0)) -> None:
    textObj = f.render(text, True, color)
    screen.blit(textObj, textObj.get_rect(center=[pos[0] - f.size(text)[0] / 2, pos[1]]))


def save():
    pass


def findFiles(path: str, filetypes: Iterable[str]):
    toreturn = []
    for filetype in filetypes:
        toreturn += glob.glob(os.path.join(path, filetype))
    return toreturn


def fileSelection(path: str, filetypes: Union[Iterable[str], None] = None) -> str:
    if filetypes is None:
        filetypes = ["*.pdf"]

    textfiles = findFiles(path, filetypes)
    subdirectories = glob.glob(os.path.join(path, '*', ''))
    displayedFiles = textfiles + subdirectories

    scroll = 0
    mode = 'select'

    illegalFiles = [os.path.join(os.path.dirname(curr_path), 'save.txt'), os.path.join(os.path.dirname(curr_path), 'game.txt')]
    while True:
        screen.fill(BGCOLOR)

        mx, my = pygame.mouse.get_pos()
        maxScroll = max(30 * len(displayedFiles) - 490, 0)

        if len(displayedFiles) == 0:
            centredPrint(font, 'No suitable files detected here!', (600, 350))

        else:
            n = 0
            for pathToFile in displayedFiles:
                if pathToFile in textfiles or (pathToFile in subdirectories and mode == 'select'):
                    if pathToFile in illegalFiles:
                        pygame.draw.rect(screen, (255, 128, 128), (25, 60 + 30 * n - scroll, 1150, 25))
                        pygame.draw.rect(screen, (0, 0, 0), (25, 60 + 30 * n - scroll, 1150, 25), 3)
                        pygame.draw.circle(screen, (255, 0, 0), (1070, 72 + 30 * n - scroll), 10, 2)
                        pygame.draw.line(screen, (255, 0, 0), (1064, 78 + 30 * n - scroll), (1076, 66 + 30 * n - scroll), 4)

                    else:
                        pygame.draw.rect(screen, (100, 100, 100), (25, 60 + 30 * n - scroll, 1150, 25))
                        if 25 <= mx <= 1175 and 60 + 30 * n - scroll <= my <= 85 + 30 * n - scroll and 60 <= my <= 550:
                            pygame.draw.rect(screen, (128, 128, 128), (25, 60 + 30 * n - scroll, 1150, 25), 5)
                        else:
                            pygame.draw.rect(screen, (0, 0, 0), (25, 60 + 30 * n - scroll, 1150, 25), 3)

                else:
                    pygame.draw.rect(screen, (150, 150, 150), (25, 60 + 30 * n - scroll, 950, 25))
                    pygame.draw.rect(screen, (64, 64, 64), (25, 60 + 30 * n - scroll, 950, 25), 3)

                if len(pathToFile) > 85:
                    pathToFileText = '...' + pathToFile[-82:]
                else:
                    pathToFileText = pathToFile
                leftAlignPrint(font, pathToFileText, (30, 72 + 30 * n - scroll))

                if pathToFile in textfiles:
                    size = os.path.getsize(pathToFile)
                    byteUnits = ['B', 'KB', 'MB', 'GB', 'TB']

                    m = 0
                    while True:
                        size = size / 1000
                        if size < 1:
                            size = size * 1000
                            break

                        if m == len(byteUnits):
                            m -= 1
                            size = size * 1000
                            break

                        m += 1

                    rightAlignPrint(font, f'{round(size, 1)} {byteUnits[m]}', (1170, 72 + 30 * n - scroll))

                n += 1

        pygame.draw.rect(screen, BGCOLOR, (0, 0, 1000, 60))
        centredPrint(mediumFont, 'File Selection', (600, 30))

        pygame.draw.rect(screen, BGCOLOR, (0, 550, 1000, 50))

        pygame.draw.rect(screen, (255, 0, 0), (30, 660, 100, 30))
        centredPrint(font, 'Cancel', (80, 675))
        if 30 <= mx <= 130 and 660 <= my <= 690:
            pygame.draw.rect(screen, (128, 128, 128), (30, 660, 100, 30), 5)
        else:
            pygame.draw.rect(screen, (0, 0, 0), (30, 660, 100, 30), 3)

        pygame.draw.rect(screen, (100, 100, 100), (150, 660, 200, 30))
        centredPrint(font, 'Parent Folder', (250, 675))
        if 150 <= mx <= 350 and 660 <= my <= 690:
            pygame.draw.rect(screen, (128, 128, 128), (150, 660, 200, 30), 5)
        else:
            pygame.draw.rect(screen, (0, 0, 0), (150, 660, 200, 30), 3)

        if mode == 'select':
            pygame.draw.rect(screen, (255, 0, 0), (25, 15, 100, 30))
            centredPrint(font, 'Delete', (75, 30))
            if 25 <= mx <= 125 and 15 <= my <= 45:
                pygame.draw.rect(screen, (128, 128, 128), (25, 15, 100, 30), 5)
            else:
                pygame.draw.rect(screen, (0, 0, 0), (25, 15, 100, 30), 3)

            rightAlignPrint(tinyFont, 'Select the file you want to open!', (1190, 30))

        elif mode == 'delete':
            rightAlignPrint(tinyFont, 'Select the file you want to delete!', (1190, 30))

        if len(path) > 60:
            pathText = '...' + path[-57:]
        else:
            pathText = path
        rightAlignPrint(tinyFont, pathText, (1190, 690))

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                save()
                quit()

            elif event.type == pygame.MOUSEBUTTONDOWN:
                if event.button == 1:
                    if 30 <= mx <= 130 and 15 <= my <= 45:
                        mode = 'delete'

                    if 30 <= mx <= 130 and 660 <= my <= 690:
                        if mode == 'select':
                            return ''

                        elif mode == 'delete':
                            mode = 'select'

                    if 150 <= mx <= 350 and 660 <= my <= 690:
                        oldpath = path
                        path = os.path.dirname(path)

                        if oldpath == os.path.join(path, ''):
                            path = os.path.dirname(path)

                        scroll = 0
                        textfiles = findFiles(path, filetypes)
                        subdirectories = glob.glob(os.path.join(path, '*', ''))
                        displayedFiles = textfiles + subdirectories

                    n = 0
                    for pathToFile in displayedFiles:
                        if pathToFile in illegalFiles:
                            n += 1
                            continue

                        if 25 <= mx <= 1175 and 60 + 30 * n - scroll <= my <= 85 + 30 * n - scroll and 60 <= my <= 550:
                            if mode == 'select':
                                if pathToFile in subdirectories:
                                    scroll = 0
                                    path = pathToFile

                                    textfiles = findFiles(path, filetypes)
                                    subdirectories = glob.glob(os.path.join(path, '*', ''))
                                    displayedFiles = textfiles + subdirectories

                                if pathToFile in textfiles:
                                    return pathToFile

                            elif mode == 'delete':
                                if pathToFile in textfiles:
                                    try:
                                        os.remove(pathToFile)
                                        print(f'Deleted file {pathToFile}!')

                                        textfiles = findFiles(path, filetypes)
                                        subdirectories = glob.glob(os.path.join(path, '*', ''))
                                        displayedFiles = textfiles + subdirectories

                                        mode = 'select'

                                    except OSError as e:
                                        print(f'excel-project: An error occured when trying to delete {pathToFile}. See details: {e}')

                        n += 1

                if event.button == 4:
                    scroll = max(0, scroll - 5)

                if event.button == 5:
                    scroll = min(maxScroll, scroll + 5)

        pressed = pygame.key.get_pressed()
        if pressed[pygame.K_UP]:
            scroll = max(0, scroll - 3)
        if pressed[pygame.K_DOWN]:
            scroll = min(maxScroll, scroll - 3)

        clock.tick(MAXFPS)
        pygame.display.update()


def main():
    status = '.'
    filename = ''

    while True:
        mx, my = pygame.mouse.get_pos()

        match status:
            case '.':
                screen.fill(BGCOLOR)

                centredPrint(mediumFont, "Welcome to AccoBuddy!", (600, 50))

                pygame.draw.rect(screen, (100, 100, 100), (30, 100, 1140, 40))
                leftAlignPrint(font, "Bank Statement Reader (DBS)", (50, 120))
                if 30 <= mx <= 1170 and 100 <= my <= 140:
                    pygame.draw.rect(screen, (128, 128, 128), (30, 100, 1140, 40), 3)
                else:
                    pygame.draw.rect(screen, (0, 0, 0), (30, 100, 1140, 40), 3)

                pygame.draw.rect(screen, (100, 100, 100), (30, 150, 1140, 40))
                leftAlignPrint(font, "Bank Statement Reader (OCBC)", (50, 170))
                if 30 <= mx <= 1170 and 150 <= my <= 190:
                    pygame.draw.rect(screen, (128, 128, 128), (30, 150, 1140, 40), 3)
                else:
                    pygame.draw.rect(screen, (0, 0, 0), (30, 150, 1140, 40), 3)

                pygame.display.update()

                for event in pygame.event.get():
                    if event.type == pygame.QUIT:
                        quit()

                    elif event.type == pygame.MOUSEBUTTONDOWN:
                        if event.button == 1:
                            if 30 <= mx <= 1170:
                                if 100 <= my <= 140:
                                    if sys.platform == "darwin":
                                        filename = fileSelection(curr_path)
                                    else:
                                        Tk().withdraw()
                                        filename = askopenfilename(filetypes=[("PDF Files", "*.pdf")])
                                    if filename != '':
                                        status = 'DBS'

                                elif 150 <= my <= 190:
                                    if sys.platform == "darwin":
                                        filename = fileSelection(curr_path)
                                    else:
                                        Tk().withdraw()
                                        filename = askopenfilename(filetypes=[("PDF Files", "*.pdf")])
                                    if filename != '':
                                        status = 'OCBC'

            case 'DBS':
                screen.fill(BGCOLOR)
                starttime = datetime.datetime.now()
                centredPrint(font, f"({filename})", (600, 50))
                centredPrint(mediumFont, "Processing Document...", (600, 120))
                centredPrint(tinyFont, f"Process Started {starttime}", (600, 150))

                sheetname = filename.replace("pdf", "xlsx")
                if len(sheetname) < 60:
                    shortname = sheetname
                else:
                    shortname = sheetname[-60:]
                centredPrint(tinyFont, f"Creating Output Excel File (...{shortname})", (600, 180))
                pygame.display.update()
                workbook = xlsxwriter.Workbook(sheetname)

                sheetname = sheetname.split('/')[-1][:-5]
                if len(sheetname) > 30:
                    sheetname = sheetname[:30]
                worksheet = workbook.add_worksheet(sheetname)

                for i in range(len(SHEETTOP)):
                    worksheet.write(0, i, SHEETTOP[i][0])
                    worksheet.set_column(i, i, SHEETTOP[i][1])

                if len(sheetname) < 30:
                    shortname = filename
                else:
                    shortname = filename[-30:]
                centredPrint(tinyFont, f"Created Output Excel File! Reading {shortname}...", (600, 210))
                pygame.display.update()

                fp = open(filename, 'rb')
                rsrcmgr = PDFResourceManager()
                laparams = LAParams()
                device = PDFPageAggregator(rsrcmgr, laparams=laparams)
                interpreter = PDFPageInterpreter(rsrcmgr, device)
                pages = list(PDFPage.get_pages(fp))

                sheety = 1
                for n in range(len(pages)):
                    screen.fill(BGCOLOR)
                    centredPrint(mediumFont, "Processing File...", (600, 100))
                    pygame.draw.rect(screen, (0, 255, 0), (30, 300, 1140 * n / len(pages), 100))
                    pygame.draw.rect(screen, (0, 0, 0), (30, 300, 1140, 100), 3)
                    centredPrint(font, f"Currently Processing Page {n + 1} of {len(pages)}...", (600, 500))
                    pygame.display.update()

                    curr = []
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

                workbook.close()
                status = '.'

            case 'OCBC':
                screen.fill(BGCOLOR)
                starttime = datetime.datetime.now()
                centredPrint(font, f"({filename})", (600, 50))
                centredPrint(mediumFont, "Processing Document...", (600, 120))
                centredPrint(tinyFont, f"Process Started {starttime}", (600, 150))

                sheetname = filename.replace("pdf", "xlsx")
                if len(sheetname) < 60:
                    shortname = sheetname
                else:
                    shortname = sheetname[-60:]
                centredPrint(tinyFont, f"Creating Output Excel File (...{shortname})", (600, 180))
                pygame.display.update()
                workbook = xlsxwriter.Workbook(sheetname)

                sheetname = sheetname.split('/')[-1][:-5]
                if len(sheetname) > 30:
                    sheetname = sheetname[:30]
                worksheet = workbook.add_worksheet(sheetname)

                for i in range(len(SHEETTOP)):
                    worksheet.write(0, i, SHEETTOP[i][0])
                    worksheet.set_column(i, i, SHEETTOP[i][1])

                if len(sheetname) < 30:
                    shortname = filename
                else:
                    shortname = filename[-30:]
                centredPrint(tinyFont, f"Created Output Excel File! Reading {shortname}...", (600, 210))
                pygame.display.update()

                fp = open(filename, 'rb')
                rsrcmgr = PDFResourceManager()
                laparams = LAParams()
                device = PDFPageAggregator(rsrcmgr, laparams=laparams)
                interpreter = PDFPageInterpreter(rsrcmgr, device)
                pages = list(PDFPage.get_pages(fp))

                sheety = 1
                for n in range(len(pages)):
                    screen.fill(BGCOLOR)
                    centredPrint(mediumFont, "Processing File...", (600, 100))
                    pygame.draw.rect(screen, (0, 255, 0), (30, 300, 1140 * n / len(pages), 100))
                    pygame.draw.rect(screen, (0, 0, 0), (30, 300, 1140, 100), 3)
                    centredPrint(font, f"Currently Processing Page {n + 1} of {len(pages)} ({round(100 * n / len(pages), 1)}% Completed)...", (600, 500))
                    centredPrint(font, f"Time Elapsed: {(datetime.datetime.now() - starttime).seconds}s", (600, 535))
                    pygame.display.update()

                    curr = []
                    interpreter.process_page(pages[n])
                    layout = device.get_result()
                    prev = 0
                    prevL = None
                    info = None
                    for lobj in layout:
                        if isinstance(lobj, LTTextBox):
                            x, y, text = lobj.bbox[0], lobj.bbox[3], lobj.get_text()

                            if x == 136.92 and text not in OCBCIGNORE:
                                info = rws(text, OCBCIGNORE).split('\n')

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

                                curr.append(d)

                            elif x == 91.56:
                                d = find(curr, round(y, 5))
                                if d is not None:
                                    d.valueDate = rws(text)

                            elif 300 <= x <= 350:
                                d = find(curr, round(y, 5))
                                if d is not None:
                                    d.withdrawal = float(rws(text))

                            elif 390 <= x <= 450:
                                d = find(curr, round(y, 5))
                                if d is not None:
                                    d.deposit = float(rws(text))

                            elif 490 <= x <= 550:
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

                workbook.close()
                status = '.'

