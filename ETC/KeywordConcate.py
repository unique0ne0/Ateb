# -*- coding: utf-8 -*-

import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('D:\\Dropbox\\A-TEB\\keyword.xlsx')
ws = wb.ActiveSheet
kw = ""
print(int(ws.Cells(2, 1).Value))
ws.Cells(2, 3).Value = ""
for i in range(int(ws.Cells(2, 1).Value)):
    kw = kw + str(ws.Cells(i+2, 2).Value) + " "
ws.Cells(2, 3).Value = kw
# excel.Quit()
