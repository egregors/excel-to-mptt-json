from openpyxl import load_workbook

wb = load_workbook(filename='1.xlsx')
ws = wb.active

c = {}
for y, row in enumerate(ws['A3':'D10']):
    tmp = ''
    for x,cell in enumerate(row):
        if not cell.value:
            tmp += '='
        # tmp = cell.value
        if cell.value:
            print(x, y,tmp, cell.value)

#print(ws.calculate_dimension())
#print(ws.max_row,ws.max_column)