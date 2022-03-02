import pandas as pd
import keyboard
from docx import Document
from docx.shared import Pt
from termcolor import colored
import os
import time


# def print_test():
#    print(table)
#    print(addressList)
#    print("address one = " + addressList[0])

os.system('color')

string_one = "Address Auto Typer found the following spreadsheet files..."
string_two = "Which file would you like to prepare? (insert # only) "
string_three = "Document is prepared. Please search for " '"workorder.docx"' " in the search bar outside of this window or your file system to find the completed document."
string_four = "You may now press enter to exit the application."

keyboard.press('F11')

def document_format():
    document = Document()
    styleEventStamp = document.styles['Normal']
    font = styleEventStamp.font
    font.name = "Bahnschrift"
    font.size = Pt(40)

    for i, value in enumerate(addressList):
        eventStampBefore = document.add_paragraph('Before             wo #1')
        address = document.add_paragraph(addressList[i])
        address.style = document.styles['Normal']
        eventStampBefore.styleEventStamp = document.styles['Normal']
        run = address.add_run()
        run.add_break()
        document.add_paragraph('During             wo #1')
        address = document.add_paragraph(addressList[i])
        run = address.add_run()
        run.add_break()
        document.add_paragraph('After             wo #1')
        document.add_paragraph(addressList[i])
        document.add_page_break()
    document.save('workorder.docx')



print('\n')
print(string_one.center(150))
print('\n')

time.sleep(1.5)
fileList = []
for root, dirs, files in os.walk(r'C:\Users\wevan\OneDrive'):
    for file in files:
        if file.endswith(".xlsx") or file.endswith(".GSHEET"):
            fileList.append(file)
for i, value in enumerate(fileList):
    print(i, "-", fileList[i])
time.sleep(1.5)
print('\n')

fileSelection = int(input(string_two.center(55)))


table = pd.read_excel("C:\\Users\wevan\OneDrive\Desktop\\" + (fileList[fileSelection]))
addressList = table["address"].tolist()

# print_test()
document_format()

print('\n')
print(colored( string_three.center(150),
    'green'))
print('\n')
input(colored(string_four.center(150), 'green'))
