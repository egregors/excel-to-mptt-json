from openpyxl import load_workbook

wb = load_workbook(filename='1.xlsx', read_only=True)
ws = wb.active

for row in ws.rows:
    for cell in row:
        if cell.value:
            print(cell.value)
