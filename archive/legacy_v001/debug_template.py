import openpyxl

path = '経歴書_Updated_20251124.xlsx'
wb = openpyxl.load_workbook(path)
ws = wb['_Template']

print("Merges in _Template (Rows 1-5):")
for mr in ws.merged_cells.ranges:
    if mr.min_row <= 5:
        print(mr)
