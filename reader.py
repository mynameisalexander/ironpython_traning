import clr
import os
from model.group import Group
import pytest


clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel


ex = Excel.ApplicationClass()   
#ex.Visible = True
ex.DisplayAlerts = False


path  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data\\")
file = path + "{0}".format(os.listdir(path)[0])


workbook = ex.Workbooks.Open(file)
ws = workbook.Worksheets[1]


testdata = []


for i in range(1, 7):
    if ws.Rows[i].Value2[0,0] == None:
        testdata.append(Group(""))
        continue
    testdata.append(Group(ws.Rows[i].Value2[0,0])) # сразу создаю ряд объектов которые потом смогу использовать для теста по созданию групп