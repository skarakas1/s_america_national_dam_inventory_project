# [description]
# This script prepares raw dam register data from the PROAGRO Inventario Nacional de Presas for Bolivia to then format for GIS-importable format.
# Authors: Sam Karakas
# Contact: skarakas21@g.ucla.edu
# Last update: 07/29/2022

infile1 = "Bolivia/presas-inventario_a_clip.xlsx"
infile2 = "Bolivia/inventario_b_clip.xlsx"
outfile = "Bolivia/bolivia_main_output_1.xlsx"

# Important packages and functions
import sys, re, datetime, xlsxwriter, openpyxl

# Open the first input spreadsheet
infile_obj1 = openpyxl.load_workbook(infile1, keep_vba=True) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet1 = infile_obj1.active #active spreadsheet of this excel
# Openn the second spreadsheet
infile_obj2 = openpyxl.load_workbook(infile2, keep_vba=True)
infile_sheet2 = infile_obj2.active
# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("output_1") #any name you want to give to the output spreadsheet

# Loop through each record in selected range
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.
row_number_write = 0 #initiate an integer to indicate the row index in the output spreadsheet.

#read infile, unmerge, skip empty cells
for each_row in infile_sheet1:
    print('looping row#: ' + str(row_number))
    row = []
    for cell in each_row:
        if cell.value:
            row.append(cell.value)

        outfile_sheet.write_row(row_number_write, 0, row)
    row_number = row_number + 1
    row_number_write = row_number_write + 1

#let's do it again for the second sheet
for each_row in infile_sheet2:
    print('looping row#: ' + str(row_number))
    row = []
    for cell in each_row:
        if cell.value:
            row.append(cell.value)

        outfile_sheet.write_row(row_number_write, 0, row)
    row_number = row_number + 1
    row_number_write = row_number_write + 1

outfile_obj.close()

print('Script ends at: ' + str(datetime.datetime.now())) 
