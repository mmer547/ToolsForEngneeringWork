import openpyxl
import xlwings as xw


wb = openpyxl.load_workbook(filename="test.xlsm", read_only=False, keep_vba=True)
ws = wb['Sheet1']
cell = ws['B2']
cell.value = 150
wb.save("test.xlsm") 
wb.close()

app = xw.apps.add()
wb = app.books.open("test.xlsm", update_links=False)
macro = wb.macro("test")
macro() 
wb.save() 
wb.close()
app.kill() 