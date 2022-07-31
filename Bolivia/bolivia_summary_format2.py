infile = "Bolivia/bolivia_output_1_summary.xlsx"
outfile = "Bolivia/bolivia_output_2_summary.xlsx"

# Important packages and functions
import sys, re, datetime, xlsxwriter, openpyxl

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile, keep_vba=True) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("output_1") #any name you want to give to the output spreadsheet

outfile_sheet.write_row(0,0, tuple(['Code','Name','Area']))

# Loop through each record in selected range
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.
row_number_write = 0 #initiate an integer to indicate the row index in the output spreadsheet. 


#for loop to read cleaned outfile
for each_row in infile_sheet:

    print('looping row#: ' + str(row_number))
    row = []

    #loop to iterate through cells
    for cell in each_row:
        # check for and skip empty cells
        if cell.value:
            row.append(cell.value) # pop cell w value

    for this_cell in row:

        cell_value = str(this_cell) # read cell value as string
        code_check = re.search(r"[A-Z][A-Z]\-[A-Z]\-\d\d\d", cell_value) # reg ex to check for 'code' values which should have a format like 'AB-X-123'
       
        if code_check: # check each cell for this regex
            
            output_row = [row[row.index(cell_value)], row[row.index(cell_value) + 1], row[row.index(cell_value) + 4]] # create a list with index 0 being the code then grabbing the next cell, the name, as index 1, finally grab cell value at index 4 which should be the area
            # we can use the area attribute as a key value to QC merging the summary data with the main data

            outfile_sheet.write_row(row_number_write, 0, output_row) # write an output file with column 1 - codes and column 2 - names

            print(output_row)

    row_number = row_number + 1
    row_number_write = row_number_write + 1

outfile_obj.close()

print('Script ends at: ' + str(datetime.datetime.now())) 