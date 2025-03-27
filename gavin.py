# this is gavin's part

import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font

gradebook = Workbook()

gradebook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")

currSheet = gradebook.active

currSheet.insert_cols(3, 3)
currSheet["C1"] = "Last Name"
currSheet["D1"] = "First Name"
currSheet["E1"] = "Student ID"

gradebook.create_sheet("Algebra")
currSheet["A5"] = "Test"

# calculate the last row 
lastRow = max((cell.row for cell in currSheet["B"] if cell.value), default=1)

for row_index, (value, ) in enumerate(currSheet.iter_rows(min_row = 2, max_row = lastRow, min_col = 2, max_col = 2, values_only = True), start=2) :

    parts = value.split("_")
    print(parts)
        
    if len(parts) == 3 :
        currSheet.cell(row=row_index, column = 3, value = parts[0])
        currSheet.cell(row=row_index, column = 4, value = parts[1])
        currSheet.cell(row=row_index, column = 5, value = parts[2])

gradebook.save(filename = "Working_Book.xlsx")
gradebook.close()