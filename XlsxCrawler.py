##reading existing xlsx file

from openpyxl import load_workbook
#for handling file
import os


path = './source'

def openFile(keyword):
#save files as list
    files = os.listdir(path)
    for file in files:
        if file.startswith(keyword):
            print (file)


wb = load_workbook(filename="Bom_Compare.xlsx")

ws = wb.active
part_name = ws['C']

# print the content
for x in range(len(part_name)):
    print(part_name[x].value)
    openFile(part_name[x].value)
