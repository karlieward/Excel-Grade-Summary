# Karlie Ward, Rachel Pinkney, Sabrina Wong, 
# Spencer B, Mason, Gavin

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

# Creates a student class to hold all of the information
class Student:
    
    def __init__(self, className, grade, last, first, studentID):
        self.id = studentID
        self.first = first
        self.last = last
        self.grade = grade
        self.className = className

# Open the excel sheet to extract information
wbOriginal = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")
wsOriginal = wbOriginal.active

# List of studnet objects
lstStudents = []

# Sets active row as the second row to avoid the column headers
activeRow = 2

# Runs through the original sheet as long as there are values in the column
while True:
    activeCell = wsOriginal.cell(row=activeRow, column=1).value  # Get value of the current cell in column A
    if activeCell is None:  # Stop if cell is empty
        break

    class_name = activeCell
    studentInfo = wsOriginal.cell(row=activeRow, column=2).value
    last, first, studentID = wsOriginal.cell(row=activeRow, column=2).value.split('_')
    studentGrade = wsOriginal.cell(row=activeRow, column=3).value

    oStudent = Student(class_name, studentGrade, last, first, studentID) # Create new student 
    lstStudents.append(oStudent)  # Append student to the list
    activeRow += 1

# for item in lstStudents:
#     print(item.get_info())


# Set up the beginning columns for each new sheet
myWorkbook = Workbook()
wsAlg = myWorkbook.create_sheet("Algebra")
wsAlg["A1"] = "Last Name"
wsAlg["B1"] = "First Name"
wsAlg["C1"] = "Student ID"
wsAlg["D1"] = "Grade"

wsTrig = myWorkbook.create_sheet("Trigonometry")
wsTrig["A1"] = "Last Name"
wsTrig["B1"] = "First Name"
wsTrig["C1"] = "Student ID"
wsTrig["D1"] = "Grade"

wsGeo = myWorkbook.create_sheet("Geometry")
wsGeo["A1"] = "Last Name"
wsGeo["B1"] = "First Name"
wsGeo["C1"] = "Student ID"
wsGeo["D1"] = "Grade"

wsCalc = myWorkbook.create_sheet("Calculus")
wsCalc["A1"] = "Last Name"
wsCalc["B1"] = "First Name"
wsCalc["C1"] = "Student ID"
wsCalc["D1"] = "Grade"

wsStats = myWorkbook.create_sheet("Statistics")
wsStats["A1"] = "Last Name"
wsStats["B1"] = "First Name"
wsStats["C1"] = "Student ID"
wsStats["D1"] = "Grade"




myWorkbook.save(filename="formatted_grades.xlsx")
myWorkbook.close()