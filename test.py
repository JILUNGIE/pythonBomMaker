import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import Workbook

month = {'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 'June': '06',
         'July': '07', 'August': '08', 'September': '09', 'October': '10', 'November': '11', 'December': '12'}


def fileExplorer():  # https://stackoverflow.com/questions/9319317/quick-and-easy-file-dialog-in-python
    filetypes = (('bom files', '*.bom'),)
    root = tk.Tk()
    root.withdraw()
    # https://www.pythontutorial.net/tkinter/tkinter-open-file-dialog/
    files = filedialog.askopenfilename(title='Select bom file', filetypes=filetypes)

    if files == '':
        messagebox.showwarning("Warning!!", "Please select a file")
    return files

def fileName(title):
    bomContents = title.split()

    bomTitle = bomContents[0]
    bomYear = bomContents[5]
    bomMonth = month[bomContents[3]]
    bomDate = bomContents[4].replace(',', '')

    bomBirthDay = "%s(%s.%s.%s)"%(bomTitle, bomYear, bomMonth, bomDate)

    return bomBirthDay


if __name__ == "__main__":

    bomFile = open(fileExplorer(), 'r')
    bomFileIntroLines = bomFile.readlines()[0]
    
    print(bomFileIntroLines)
    bomtitle = "%s.xlsx"%(fileName(bomFileIntroLines))
    bomFile.close()

    # wb = Workbook()
    # ws = wb.active

    # ws.title = "BOM"  # sheet 의 이름 변경
    # wb.save(bomtitle)
    # wb.close()