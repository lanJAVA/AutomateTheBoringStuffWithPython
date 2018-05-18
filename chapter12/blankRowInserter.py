#! python3
# blankRowInserter.py - 12.13.2 空行插入

import sys, openpyxl

if len(sys.argv) != 4:
    print("Usage: blankRowInserter.py row blankNum fileName")
    sys.exit()

row = int(sys.argv[1])
blankNum = int(sys.argv[2])
fileName = sys.argv[3]

wb = openpyxl.load_workbook(fileName)
sheet = wb.active
wb1 = openpyxl.Workbook()
sheet1 = wb1.active
for i in range(1, row):
    for j in range(1, sheet.max_column + 1):
        sheet1.cell(row=i, column=j).value = sheet.cell(row=i, column=j).value
for i in range(row, sheet.max_row + 1):
    for j in range(1, sheet.max_column + 1):
        sheet1.cell(row=i+blankNum, column=j).value = sheet.cell(row=i, column=j).value

wb1.save('copy' + fileName)