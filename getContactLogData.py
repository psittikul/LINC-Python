import openpyxl
wb = openpyxl.load_workbook('LINC Demographics for LVAIC.xlsx')
print(type(wb))
contactLogSheet = wb.get_sheet_by_name('Contact Log Info')
print(contactLogSheet['A2'].value)

# Cycle through all the rows of our contact log data
for row in contactLogSheet['A2:E1320']:
