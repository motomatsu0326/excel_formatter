import openpyxl
import os
import re
path = "."
files = os.listdir(path)
excel_files = [file for file in files if os.path.isfile(os.path.join(path, file)) and re.compile(r"^(.*).(xls|xlsx)").match(file)]
for file in excel_files:
    wb = openpyxl.load_workbook(file)
    sheet_names = wb.sheetnames
    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        sheet.sheet_view.selection[0].activeCell = 'A1'
        sheet.sheet_view.selection[0].sqref = 'A1'
        sheet.sheet_view.tabSelected = False
        sheet.sheet_view.zoomScale = 100
        sheet.sheet_view.zoomScaleNormal = 100
        sheet.sheet_state = 'visible'
    wb.active = 0
    wb.save(file)
