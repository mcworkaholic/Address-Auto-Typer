import pandas as pd
import keyboard
from docx import Document
from docx.shared import Pt
from termcolor import colored
import os

# def print_test():                     #uncomment this function and line 76 to test if the information produced is correct
#    print(table)
#    print(addressList)
#    print("address one = " + addressList[0])

os.system('color')
string_one = "Address Auto Typer found the following spreadsheet files..."
string_two = "Which file would you like to prepare? (insert # only) "
string_four = "Press enter to continue."

keyboard.press('F11')

print('\n')
print(string_one.center(150))
print('\n')


fileList = []                         #lines 25-32 traverses through file system to find spreadsheet files and shows them to the user
for root, dirs, files in os.walk(r'C:\Users\wevan\OneDrive'):
    for file in files:
        if file.endswith(".xlsx") or file.endswith(".GSHEET"):
            fileList.append(file)
for i, value in enumerate(fileList):
    print(i, "-", fileList[i])
print('\n')

fileSelection = int(input(string_two.center(55)))
fileName = fileList[fileSelection]
table = pd.read_excel("C:\\Users\wevan\OneDrive\Desktop\\" + fileName)
addressList = table["address"].tolist()


def convert(string):                       #function that finds the number within "work order 1.xslx" for document_format()
    fileList2 = []
    orderNum = []
    fileList2[:0] = string
    for i, value in enumerate(fileName):
        if fileList2[i].isdigit():
            orderNum.append(fileList2[i])
    global order
    order = (''.join(str(x) for x in orderNum))
    return order


def document_format():                          # function formats the word document to specification for Snowpros LLC
    document = Document()
    styleEventStamp = document.styles['Normal']
    font = styleEventStamp.font
    font.name = "Bahnschrift"
    font.size = Pt(40)

    for i, value in enumerate(addressList):
        eventStampBefore = document.add_paragraph('Before             wo #' + order)
        address = document.add_paragraph(addressList[i])
        address.style = document.styles['Normal']
        eventStampBefore.styleEventStamp = document.styles['Normal']
        run = address.add_run()
        run.add_break()
        document.add_paragraph('During             wo #' + order)
        address = document.add_paragraph(addressList[i])
        run = address.add_run()
        run.add_break()
        document.add_paragraph('After                wo #' + order)
        document.add_paragraph(addressList[i])
        document.add_page_break()
    document.save('C:\\Users\wevan\OneDrive\Desktop\\workorder.docx')


# print_test()                  # uncomment this line and lines 8-11 to test and make sure the information produced is correct
convert(fileName)
document_format()

print('\n')
print('\n')
input(colored(string_four.center(150), 'green'))
os.system('start C:\\Users\wevan\OneDrive\Desktop\\workorder.docx')             # opens word document after the press of enter
os.system("TASKKILL /F /IM main.exe")                           # closes program window after the word document is opened

