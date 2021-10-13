##reading existing xlsx file

from openpyxl import load_workbook
#for handling file
import os


path = './source'


#Secondary workbook loader
def checkXlsx(book):
    wbs = load_workbook(filename="./source/" + book)
    wss = wbs.active
    row=wss['9']
    value = 0
    for x in row:
        if x.value is None:
            value = 1
        else:
            value = 0
            break
    
    return value



def openFile(keyword):
#save files as list
    files = os.listdir(path)
    
    for file in files:
        if file.startswith(keyword):
            print (file)
            print(checkXlsx(file))
            return checkXlsx(file)


#Primary workbook loader
wb = load_workbook(filename="Bom_Compare.xlsx")

ws = wb.active
part_name = ws['C']

# print the content
for x in range(len(part_name)):
    print(part_name[x].value)
    if openFile(part_name[x].value) is 1:
        ws['E'+str(x+1)]='No Difference in Windchill and QAD'
        print ('E'+str(x+1))
    

wb.save('Result.xlsx')