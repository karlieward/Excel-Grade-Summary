# Karlie Ward, Rachel Pinkney, Sabrina Wong, 
# Spencer B, Mason, Gavin

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font


myWorkbook = Workbook()
myWorkbook.create_sheet("Algebra")
myWorkbook.create_sheet("Trigonometry")
myWorkbook.create_sheet("Geometry")
myWorkbook.create_sheet("Calculus")
myWorkbook.create_sheet("Statistics")

myWorkbook.save(filename="formatted_grades.xlsx")
myWorkbook.close()