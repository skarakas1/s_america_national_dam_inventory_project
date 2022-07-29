infile = "Bolivia_Summary.xlsx"
outfile = "bolivia_output_1.xlsx"

# Important packages and functions
import sys, datetime, xlsxwriter, openpyxl

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("output_1") #any name you want to give to the output spreadsheet

# Loop through each record in selected range
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.
row_number_write = 0 #initiate an integer to indicate the row index in the output spreadsheet. 

def print_row_num():
    print('looping row#: ' + str(row_number))

#for loop to read infile
for each_row in infile_sheet:
    
    #create cell # variable w/in row loop to start cell numbering at 0 for each row
    cell_number = 0

    print_row_num()

    row = []

    #loop to iterate through cells
    for cell in each_row:
        
        #skip empty cells
        if cell.value:

            row.append(cell.value)

            outfile_sheet.write_row(row_number_write, 0, row)
            #row_number_write = row_number_write + 1

            #output string
            #print('row ' + str(row_number) + ' cell: ' + str(cell_number) + ' = ' + str(cell.value))

            #cell_number = cell_number + 1 #count which cell we are reading

    row_number = row_number + 1 #count which row we are reading from the input spreadsheet.
    row_number_write = row_number_write + 1

outfile_obj.close() #close output file

print('Script ends at: ' + str(datetime.datetime.now())) #print end time
