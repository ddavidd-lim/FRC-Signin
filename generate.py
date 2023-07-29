import csv
import openpyxl
from openpyxl.styles import PatternFill
import os
import datetime



#first row has 5 dates, tues-sat
#second row is under first row

workbook = openpyxl.Workbook()

# Select the active sheet
sheet = workbook.active
sheet.title = "Week 1"

year = 2023
month = 7
day = 18

start_date = datetime.datetime(year, month, day)
inc_delta = datetime.timedelta(days=1)
week_delta = datetime.timedelta(days=7)


headers = []    #"Tuesday", "","", "", 
                #"Wednesday", "","", "",
                #"Thursday", "","","",  
                #"Friday", "","", "",   
                #"Saturday", "","", "", 

for i in range(5):
    headers.append(start_date.strftime("%A, %B %d, %Y"))
    start_date += inc_delta

# Data for the 3 columns beneath each cell of the main column
sub_column_data = ["Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason"]

# Write headers
column = 1
header_num = 0
# Fill in color for each column
for i in range(5):
    sheet.cell(row=1, column=column, value=headers[header_num]).fill = PatternFill(start_color='2273D1'
                                                                                   , end_color='2273D1', fill_type='solid')
    column += 4
    header_num += 1

start_cell = "A1"
end_cell = "D1"


# Merge Day columns
for i in range(5):
        sheet.merge_cells(f"{start_cell}:{end_cell}")
        start_cell = chr(ord(start_cell[0])+4) + "1"
        end_cell = chr(ord(end_cell[0])+4) + "1"

sheet.append(sub_column_data)

# Set width of name columns
sheet.column_dimensions['B'].width = 22
sheet.column_dimensions['F'].width = 22
sheet.column_dimensions['J'].width = 22
sheet.column_dimensions['N'].width = 22
sheet.column_dimensions['R'].width = 22

cd = os.getcwd()
filename = "Fall2024.xlsx"
file_path = os.path.join(cd, filename)

workbook.save(file_path)

