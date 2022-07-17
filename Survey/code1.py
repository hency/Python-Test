# import xlwings as xw
# import os
# @xw.func
# def cwf():
#     return os.getcwd()
import xlwings as xw
def test_vba():
    wb = xw.Book.caller()
    sht = wb.sheets[0]
    sht.range('A1').value = 'python知识学堂'
# import xlwings as xw
# @xw.sub
# def my_macro():
#  wb = xw.Book.caller()
#  wb.sheets[0].range('A1').value = wb.name