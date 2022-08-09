import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

month = {'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 'June': '06',
         'July': '07', 'August': '08', 'September': '09', 'October': '10', 'November': '11', 'December': '12'}


def fileExplorer():  # https://stackoverflow.com/questions/9319317/quick-and-easy-file-dialog-in-python
    filetypes = (('bom files', '*.bom'),)
    root = tk.Tk()
    root.withdraw()
    # https://www.pythontutorial.net/tkinter/tkinter-open-file-dialog/
    return filedialog.askopenfilename(title='Select bom file', filetypes=filetypes)


def fileNameMaker():
    bomContents = bomFileLines.split()

    bomTitle = bomContents[0]
    bomYear = bomContents[5]
    bomMonth = month[bomContents[3]]
    bomDate = bomContents[4].replace(',', '')

    bomBirthDay = bomYear+'.'+bomMonth+'.'+bomDate
    return bomTitle, bomBirthDay


if __name__ == "__main__":
    wb = Workbook()
    ws = wb.active
    try:
        bomFile = open(fileExplorer(), 'r')
        bomFileLines = bomFile.readline()
        bomFile.close()
    except:
        raise Exception("Select File.")

    fileTitle, fileDate = fileNameMaker()
    bomTitle = fileTitle+"("+fileDate+")"

    ws.title = "BOM"  # sheet 의 이름 변경
    wb.save(bomTitle+".xlsx")
    wb.close()
