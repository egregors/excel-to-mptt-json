from openpyxl import load_workbook
from functools import reduce

wb = load_workbook(filename='1.xlsx')
ws = wb.active

# replace all none cells with upper cell, by column
for col in ws.iter_cols():
    tmp = ''
    for cell in col:
        if not cell.value:
            cell.value = tmp
        else:
            tmp = cell.value

# read to check
for row in ws.iter_rows(min_row=3, max_col=4, max_row=3):  # debug on one row
    c = []
    for cell in row:
        c.append(cell.value)
    print(c)
    d = reduce(lambda x, y: {y: x}, reversed(c + ['']))
    print(d)
