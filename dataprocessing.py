#!/usr/bin/python
# -*- coding: gbk -*-

# project to recalculate the numbers

from openpyxl import load_workbook

wb = load_workbook(filename='./data/ivol1.xlsx', read_only=True)
ws = wb['Sheet1'] # ws is now an IterableWorksheet

i=0
ivols=[]

for row in ws.rows:
    ivols.append([row[0].value,row[1].value,row[2].value,row[3].value,row[4].value])
    for cell in row:
        print(cell.value)
        i=i+1
    if i > 100:
        break

for list in ivols:
    print ivols

