infile = "Bolivia/Bolivia_Summary.xlsx"
outfile = "Bolivia/bolivia_output_1.xlsx"

# Important packages and functions
import sys, re, datetime, xlsxwriter, openpyxl

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile, keep_vba=True) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("output_1") #any name you want to give to the output spreadsheet

# Loop through each record in selected range
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.
row_number_write = 0 #initiate an integer to indicate the row index in the output spreadsheet. 

#read infile, unmerge, skip empty cells

for each_row in infile_sheet:
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
