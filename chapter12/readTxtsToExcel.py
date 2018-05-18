#! python3
# readTxtsToExcel.py - 12.13.4 文本文件到电子表格

import sys, openpyxl

if len(sys.argv) < 2:
    print('Usage: readTxtsToExcel.py txtFile1 [txtFile2...]')
    sys.exit()

txtFiles = sys.argv[1:]

wb = openpyxl.Workbook()
sheet = wb.active

for i in range(1, len(txtFiles) + 1):
    txt = open(txtFiles[i-1], encoding='utf8')
    lines = txt.readlines()
    for j in range(1, len(lines) + 1):
        sheet.cell(row=j, column=i).value = lines[j-1]
    txt.close()

wb.save('txtFiles.xlsx')