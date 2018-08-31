from openpyxl import Workbook

# Reference: https://openpyxl.readthedocs.io/en/stable/tutorial.html#create-a-workbook

# A workbook is always created with at least one worksheet.
wb = Workbook()

# You can get it by using the
# openpyxl.workbook.Workbook.active()
ws = wb.active

# Create new worksheets
ws1 = wb.create_sheet("My sheet1")      # insert at the end (default)
ws2 = wb.create_sheet("My sheet2", 0)   # insert at first position

# Change sheet's title

# Saving to a file
wb.save('sample.xlsx')
