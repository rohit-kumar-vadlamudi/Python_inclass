import openpyxl
wb = openpyxl.load_workbook('realestatedata.xlsx')
mysheets = wb.sheetnames

print("print all sheets in Excel Workbook")
for x in mysheets:
  print(x)

sheet = wb['peter']

print("print type sheet")
print(type(sheet))
print("sheet name") 
print(sheet.title)
