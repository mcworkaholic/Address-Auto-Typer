import pandas as pd
import keyboard
from docx import Document
from docx.shared import Pt
from termcolor import colored
import os
import time


# def print_test():                # Uncomment lines 10-13 and line 70 to print out the table from the excel file you wish to work with for testing
#    print(table)
#    print(addressList)
#    print("address one = " + addressList[0])

os.system('color')     # forces the .exe window to use termcolor and color the text in the window

string_one = "Address Auto Typer found the following spreadsheet files..."
string_two = "Which file would you like to prepare? (insert # only) "
string_three = "Document is prepared. Please search for " '"workorder.docx"' " in the search bar outside of this window or your file system to find the completed document."
string_four = "You may now press enter to exit the application."

keyboard.press('F11') # forces the window to maximize when ran

def document_format(): # function that formats and saves workorder.docx based on how Snowpros LLC operates
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



print('\n') # line break
print(string_one.center(150)) # centers text
print('\n')

time.sleep(1.5) # artificial slow down for user
fileList = [] # declares empty list to append to 
for root, dirs, files in os.walk(r'C:\path'): # runs through all files to find .xlsx or .GSHEET, make sure to specify path such as "onedrive"
    for file in files:
        if file.endswith(".xlsx") or file.endswith(".GSHEET"):
            fileList.append(file) # adds each file to empty list declared above 
for i, value in enumerate(fileList): # formats the display of the files for the user
    print(i, "-", fileList[i])
time.sleep(1.5)
print('\n')

fileSelection = int(input(string_two.center(55))) # gets input from user for file selection


table = pd.read_excel("C:\path\\" + (fileList[fileSelection])) # make sure to specify path to which excel file you would like to use
addressList = table["address"].tolist()

# print_test() # uncomment before print and lines 10-13 to test
document_format()

print('\n')
print(colored( string_three.center(160), 'green'))
print('\n')
input(colored(string_four.center(150), 'green'))
