# INPUT
# Define the input excel
infile = r"Peru_pdf_clip.xlsx"

# OUTPUT
# Define the output excel
outfile = r"peru_output.xlsx"

#========================================================

# Important packages and functions
import sys, datetime, xlsxwriter, openpyxl

# Print script starting time
print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("natl_reg_peru")

# Loop through each record in the intput spreadsheet
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.
row_number_write = 0 #initiate an integer to indicate the row index in the output spreadsheet.

#unmerge all cells?

#for loop to read infile
for each_row in infile_sheet:
    print('looping row#: ' + str(row_number))
    
    #here we will need to
    # check if the row has a value
    # if that value is 'código'
    # then skip to the next row (to avoid including header values) at the same column
    # then write that cell's value to the output file
    # then continue for the next 8 cells (9 per row in total)
    # then, if next row cell with same column value as 'código' is NOT empty, move to reading next row for 9 cells
        # if it is empty, start loop again to find next cell with value 'código'
