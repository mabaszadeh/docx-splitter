import os
from sys import exit
from signal import signal, SIGINT
from win32com import client


# CONFIGURATIONS
############################################
print("DocxTop [Version 1.3]\n\n")
FOLDER_PATH = input("Enter Folder Path: ")
TOP = int(input("\nHow many TOP pages to split? "))
SAVE_TO = r"{}\out.xlsx".format(FOLDER_PATH)
VISIBLE = True


def error(x):
    return f"\033[91mERROR: {x}\033[0m"
#############################################


if not os.path.exists(FOLDER_PATH):
    exit(error("THE SPECIFIED PATH DOESN'T EXIST; CHECK THE SPELLING AND TRY AGAIN."))


# SIGINT Handler
#############################################
def SIGINT_handler(signal, frame):
    if not (globals().get("word") is None):
        globals()["word"].Application.Quit()
        del globals()["word"]

    if not (globals().get("excel") is None):
        globals()["excel"].ActiveWorkbook.Close(SaveChanges=False)
        globals()["excel"].Application.Quit()
        del globals()["excel"]

    exit(0)


signal(SIGINT, SIGINT_handler)
#############################################


excel = client.dynamic.Dispatch('Excel.Application')
word = client.dynamic.Dispatch('Word.Application')

excel.Visible = VISIBLE
excel.DisplayAlerts = False
word.Visible = False
word.DisplayAlerts = False


workbook = excel.Workbooks.Add()
sheet1 = workbook.ActiveSheet
sheet1.Range("A1:B1").Value = ("NAME", f"FIRST {TOP} PAGES")


# HEAD
#############################################
def head(x, row):
    rng = word.ActiveDocument.GoTo(
        Count=1, What=client.constants.wdGoToPage, Which=client.constants.wdGoToAbsolute)

    word.Selection.GoTo(Count=x, What=client.constants.wdGoToPage,
                        Which=client.constants.wdGoToAbsolute)

    rng.End = word.Selection.Bookmarks("\\Page").Range.End

    sheet1.Cells(row, 2).Value = rng
#############################################


row = 4


for file in os.listdir(FOLDER_PATH):

    extension = os.path.splitext(file)[1]
    if not (extension == ".docx" or extension == ".doc"):
        continue

    try:
        filepath = FOLDER_PATH + "\\" + file
        word.Documents.Open(filepath)
        sheet1.Cells(row, 1).Value = file

        pages_count = int(word.ActiveDocument.Range().Information(
            client.constants.wdNumberOfPagesInDocument))

        if pages_count >= TOP:
            head(TOP, row)
        else:
            head(pages_count, row)

        word.ActiveDocument.Close()
        row += 1

    except Exception:
        # if file is corrupted then:
        word.Application.Quit()
        del word
        word = client.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False


word.Application.Quit()
del word
workbook.SaveAs(Filename=SAVE_TO)
excel.Application.Quit()
del excel
