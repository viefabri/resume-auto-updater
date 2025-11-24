import openpyxl

path = '経歴書_Updated_20251124.xlsx'
try:
    wb = openpyxl.load_workbook(path)
    print(f"Sheets in {path}: {wb.sheetnames}")
except Exception as e:
    print(f"Error: {e}")
