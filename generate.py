import csv
import openpyxl
import os
import datetime



#first row has 5 dates, tues-sat
#second row is under first row

workbook = openpyxl.Workbook()

# Select the active sheet
sheet = workbook.active
sheet.title = "Week 1"
sheet = workbook["Week 1"]

def merge_cells_in_sheet(sheet_name, start_cell, end_cell):
    # Select the desired sheet
    sheet = workbook[sheet_name]

    # Merge the cells
    sheet.merge_cells(f"{start_cell}:{end_cell}")

# Data for the main column (3 cells wide)
main_column_data = [
    "Tuesday", "","", "",  # Two empty cells for Tuesday
    "Wednesday", "","", "", # Two empty cells for Wednesday
    "Thursday", "","","",   # Two empty cells for Thursday
    "Friday", "","", "",    # Two empty cells for Friday
    "Saturday", "","", "", # Two empty cells for Saturday
]

# Data for the 3 columns beneath each cell of the main column
sub_column_data = ["Arrival Time", "Name", "# in Party", "Reason","Arrival Time", "Name", "# in Party", "Reason","Arrival Time", "Name", "# in Party", "Reason","Arrival Time", "Name", "# in Party", "Reason","Arrival Time", "Name", "# in Party", "Reason"]

# Write main column data
for col_idx, data in enumerate(main_column_data, start=1):
    sheet.cell(row=1, column=col_idx, value=data)

start_cell = "A1"
end_cell = "D1"



for i in range(5):
        merge_cells_in_sheet("Week 1", start_cell, end_cell)
        start_cell = chr(ord(start_cell[0])+4) + "1"
        end_cell = chr(ord(end_cell[0])+4) + "1"

sheet.append(sub_column_data)

        

cd = os.getcwd()
filename = "output.xlsx"
file_path = os.path.join(cd, filename)

workbook.save(file_path)

