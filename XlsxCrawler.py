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
        print(x.value)
        if x.value is None:
            value = 1
            

        elif x.value == 'No Values Present in Windchill':
            value = 2
            

                        
        elif x.value == 'Part is not Effective Windchill':
            value = 3
            

        else :
            continue
    
    

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

    if openFile(part_name[x].value) is 2:
        ws['E'+str(x+1)]='No Values Present in Windchill'
        print ('E'+str(x+1))

    if openFile(part_name[x].value) is 3:
        ws['E'+str(x+1)]='Part is not Effective Windchill'
        print ('E'+str(x+1))
    

wb.save('Result.xlsx')