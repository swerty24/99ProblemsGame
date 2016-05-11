# -*- coding: utf-8 -*-
"""
5-5-2016 

This is a python program to generate the 99 Problems Game as an excel spreadsheet
The end result should be able to (1) generate a spreadsheet of all the problems and
solutions as a printable excel file. It should also be able to (2) add and remove problems
or solutions to the complete list through a text command line.

TODO:
complete problems dict
complete solutions dict


Modify for length of problems for both Row's and Columns.
Determine the length of problems for a 4 column Sheet.


QR codes
Borders around cells so you know where to cut!
Enable Re-sizing of text size if over a certain number of characters!

Part 2 or requirements
"""

#Imports
import math
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.cell import get_column_letter, Cell
from openpyxl.styles import Fill, Color, PatternFill, Border, Side, Alignment, Font
from openpyxl.styles import colors
from string import ascii_uppercase

#Create Workbook using open Pyxl package
wb = Workbook()


#List of Problems (currently just a sample for testing)
problems = {1: 'Pollution', 
            2: 'termites',
            3: 'jobs',
            4: 'Ant Farms', 
            5: 'Forest Fires',
            6: 'Traffic',
            7: 'e-waste', 
            8: 'black mirrors',
            9: 'AI',
            10: 'Miscommunication', 
            11: 'parking lots',
            12: 'people',
            13: 'War'}

# Set Workbook File Name and Sheet Name
dest_filename = 'TestOutput.xlsx'
ws1 = wb.active
ws1.title = "To Print"

#Only use 4 Columns with width 30 
ws1.column_dimensions["A"].width = 30
ws1.column_dimensions["B"].width = 30
ws1.column_dimensions["C"].width = 30
ws1.column_dimensions["D"].width = 30

#Set the Color of Text
rgb=[255,0,0]
color_string="".join([str(hex(i))[2:].upper().rjust(2, "0") for i in rgb])


#Loop through all the columns less than the letter E, and set each cell to red and set size of text for cell
for ProblemColumnLetter in ascii_uppercase:
    for i in range(1,5):
        
        ws1[ProblemColumnLetter+str(i)].fill=PatternFill(fill_type="solid", start_color='FF' + color_string, end_color='FF' + color_string) 
        ws1[ProblemColumnLetter+str(i)].font = Font(size = 23)
        print ProblemColumnLetter + str(i)
    if ProblemColumnLetter == "E":
        break
       
#Calculate the number of rows needed (based on 4 columns) and store into x
x = int(math.ceil(len(problems)/4.0))+1
print x

#print each of the problems into the cells
i = 1
for row in range(1, x):
    for col in range(1, 5):
            _ = ws1.cell(column=col, row=row,value="%s" % problems[i])
            if i< len(problems):
                i+= 1
            else:     
                break
            
#The following is example code for working with openpyxl            
"""               
ws2 = wb.create_sheet(title="Pi")

ws2['F5'] = 3.14

ws3 = wb.create_sheet(title="Data")
for row in range(10, 20):
    for col in range(27, 54):
        _ = ws3.cell(column=col, row=row, value="%s" % get_column_letter(col))
print(ws3['AA10'].value)

wb.save(filename = dest_filename)

print(problems[2])
"""
