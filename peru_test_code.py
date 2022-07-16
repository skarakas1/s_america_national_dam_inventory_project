# INPUT
# Define the input excel
infile = "Peru_pdf_clip.xlsx"

# Important packages and functions
import sys, datetime, xlsxwriter, openpyxl

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Loop through each record in selected range
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.

#printing cells A1 through BP8 to see what kind of values python is reading from this heavily formatted spreadsheet

def print_row_num():
    print('looping row#: ' + str(row_number))
#define cell range
test_range = infile_sheet['A1':'BP8']
#for loop to read infile
for each_row in test_range:
    
    cell_number = 0
    
    print_row_num()

    for each_cell in each_row:
        if each_cell.value:
            print('row' + str(row_number) + ' cell: ' + str(cell_number) + ' = ' + str(each_cell.value))

            cell_number = cell_number + 1
            #count which cell we are reading

    row_number = row_number + 1 #count which row we are reading from the input spreadsheet.
