import os

import win32com.client as win32

default_path = 'D:\\Documents\\업무\\업적평가\\2016\\하반기\\팀원\\업적평가서\\임시\\'
src = os.path.join(default_path, '업적평가서_AMG-Linux_임영현_2016.xlsx')
dst = os.path.join(default_path, "test.xlsx")

excel = win32.gencache.EnsureDispatch('Excel.Application')
wbS = excel.Workbooks.Open(src)
wbD = excel.Workbooks.Add()
wsS = wbS.Worksheets(r"2)업적보고서(공통)")
wsS.Name = "임영현"
wsS.Copy(wbD.Worksheets(1))
wbS.Close(False)
excel.DisplayAlerts = False
wbD.SaveAs(dst)
excel.DisplayAlerts = True
wbD.Close(True)
excel.Quit()
del excel

"""

2)업적보고서(공통)
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('D:\\Documents\\업무\\업적평가\\2016\\하반기\\팀원\\업적평가서\\업적평가서_AMG-Linux_임영현_2016.xlsx')
dst_wb = excel.Workbooks.Add()
#wb.Worksheets("2)업적보고서(공통)").Copy(dst_wb.dst_wb.ws)
dst_wb.SaveAs(dst)
excel.Quit()
"""
