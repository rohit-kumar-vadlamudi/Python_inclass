import openpyxl 
wb = openpyxl.load_workbook(r'D:\AIML-PGDM\Term 1\Python-1204\Data files\realestatedata.xlsx')

ws = wb.sheetnames

print('Print all sheets in work book')
for x in ws:
	print(x)

sheet = wb['peter']

print('Sheet type')
print(type(sheet))
print('Sheet name')
print(sheet.title)