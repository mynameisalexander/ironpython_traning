import clr
import os
from model.group import Group
import pytest


clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel


ex = Excel.ApplicationClass()   
#ex.Visible = True
ex.DisplayAlerts = False

path = "C:/Devel/ironpython_traning/data/"

file = path + "{0}".format(os.listdir(path)[0])

workbook = ex.Workbooks.Open(file)
ws = workbook.Worksheets[1]


testdata = []

for i in range(1, 7):
    testdata = ws.Rows[i].Value2[0,0] # - не наполняется массив данными из excel
    #testdata = Group(ws.Rows[i].Value2[0,0])
    print(type(testdata))


@pytest.mark.parametrize("group", testdata)
def test_create_groups(group):
    pass

