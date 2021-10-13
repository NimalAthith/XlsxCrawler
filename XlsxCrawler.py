##reading existing xlsx file

from openpyxl import load_workbook

wb = load_workbook(filename="Bom_Compare.xlsx")

ws = wb.active
part_name = ws['C']

# pritn the content
for x in range(len(part_name)):
    print(part_name[x].value)
