# Karlie Ward, Rachel Pinkney, Sabrina Wong, 
# Spencer Bigelow, Mason Zarges, Gavin Smith

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
wbOriginal = openpyxl.load_workbook("Poorly_Organized_Data_2.xlsx")
wsOriginal = wbOriginal.active

# List of student objects
lstStudents = []

# Class Dictionary
dictionaryClass = {}

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

    if class_name in dictionaryClass:
        dictionaryClass[class_name].append(oStudent)
    else:
        dictionaryClass[class_name] = []
        dictionaryClass[class_name].append(oStudent)

    activeRow += 1

# Set up the beginning columns for each new sheet
myWorkbook = Workbook()

for key in dictionaryClass:
    newWS = myWorkbook.create_sheet(key)
    newWS["A1"] = "Last Name"
    newWS["B1"] = "First Name"
    newWS["C1"] = "Student ID"
    newWS["D1"] = "Grade"

    count = 2
    for student in dictionaryClass[key]:
        newWS.cell(row=count, column=1).value = student.last
        newWS.cell(row=count, column=2).value = student.first
        newWS.cell(row=count, column=3).value = student.id
        newWS.cell(row=count, column=4).value = student.grade
        
        count += 1

myWorkbook.save(filename="formatted_grades.xlsx")
myWorkbook.close()