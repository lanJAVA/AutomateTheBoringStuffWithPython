#! python3
# excelToTxts.py - 12.13.5 电子表格到文本文件

import sys, openpyxl

if len(sys.argv) != 2:
    print('Usage: excelToTxts.py excelFileName')
    sys.exit()

fileName = sys.argv[1]

wb = openpyxl.load_workbook(fileName)
sheet = wb.active
columns = list(sheet.columns)
for i in range(len(columns)):
    txtName = fileName.replace('.xlsx', '') + '_' + str(i + 1) + '.txt'
    txt = open(txtName, 'w', encoding='utf8')
    for cell in columns[i]:
        if cell.value:
            txt.write(cell.value)
    txt.close()
