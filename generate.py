import csv
import openpyxl
from openpyxl.styles import PatternFill
import os
import datetime



#first row has 5 dates, tues-sat
#second row is under first row

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Week 1"

inc_delta = datetime.timedelta(days=1)
week_delta = datetime.timedelta(days=7)

def writeDayHeaders(start_date):
    headers = []
    # store dates from start time
    for i in range(5):
        headers.append(start_date.strftime("%A, %B %d, %Y"))
        start_date += inc_delta
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
            
    # Set width of name columns
    sheet.column_dimensions['B'].width = 22
    sheet.column_dimensions['F'].width = 22
    sheet.column_dimensions['J'].width = 22
    sheet.column_dimensions['N'].width = 22
    sheet.column_dimensions['R'].width = 22

def writeSubheaders():
    sub_column_data = ["Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason",
                   "Arrival", "Name", "# in Party", "Reason"]

    sheet.append(sub_column_data)

def saveWorkbook(filename):
    cd = os.getcwd()
    filename = filename
    file_path = os.path.join(cd, filename)

    workbook.save(file_path)
    return 1

if __name__ == "__main__":
    # Date inputs
    # 7/25/2023
    while True:
        try:
            year = int(input(">>Enter Year: "))
            month = int(input(">>Enter Month: "))
            day = int(input(">>Enter a Tuesday: "))

            start_date = datetime.datetime(year, month, day)
            break
        except ValueError as v:
            print("!! Invalid input")
    
    writeDayHeaders(start_date)
    writeSubheaders()
    
    filename = input(">>Enter the name of your file: ")
    filename += ".xlsx"
    
    if saveWorkbook(filename) != 1:
        print(">>Failed to save file")
        exit()
    print(">>Success")

