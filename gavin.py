# this is gavin's part

import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font

gradebook = Workbook()

gradebook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")

currSheet = gradebook.active

currSheet.insert_cols(3, 3)
currSheet["C1"] = "Student ID"

currSheet["C2"] = currSheet["B1"]

gradebook.create_sheet("Algebra")
currSheet["A5"] = "Test"

# calculate the last row 
lastRow = max((cell.row for cell in gradebook["B"] if cell.value), default = 1)

for row in gradebook.iter_rows(min_row = 2, max_row = lastRow, min_col = 2, max_col = 2, values_only = True) :
    pass

# gradebook.active = gradebook["Algebra"]

gradebook.save(filename = "Working_Book.xlsx")