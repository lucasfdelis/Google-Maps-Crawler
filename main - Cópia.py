import xlwings as xw

wb = xw.Book('Example.xlsx')
sht1 = wb.sheets['Plan1']
sht1.range('B2').value = 46
