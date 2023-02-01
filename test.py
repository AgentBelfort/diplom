from openpyxl import Workbook
wb = Workbook()
ws = wb.active

for i in range(1,10):
    for j in range(1,10):
        ws.cell(row=i,column=j).value = i
wb.save("sample.xlsx")
wb.close()