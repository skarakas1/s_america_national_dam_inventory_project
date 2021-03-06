# Define the input & output excel docs
infile = "Peru_pdf_clip.xlsx"
outfile = "Peru_test_output.xlsx"

# Important packages and functions
import sys, datetime, xlsxwriter, openpyxl

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("dams_subset") #add worksheet to xlsx file

outfile_sheet.write(0,0, 'Code')
outfile_sheet.write(0,1, 'East')
outfile_sheet.write(0,2, 'North')
outfile_sheet.write(0,3, 'Dam Name')
outfile_sheet.write(0,4, 'Administrative Authority')
outfile_sheet.write(0,5, 'Local Administration')
outfile_sheet.write(0,6, 'River')
outfile_sheet.write(0,7, 'District')
outfile_sheet.write(0,8, 'Province')

# Loop through each record in selected range
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.
row_number_write = 1 #initiate an integer to indicate the row index in the output spreadsheet.
#initiated at '1' to not write over previous manually written header row

#for loop to read infile
for each_row in infile_sheet:

    print('looping row#: ' + str(row_number))

    row = []

    #loop to iterate through cells
    for cell in each_row:
        
        # check for and skip empty cells
        if cell.value:

            row.append(cell.value) # pop cell w value

    if int(len(row)) > 1: #this ignores rows that just contain page number

        if str(row[0]) != "Código": #this ignores header rows

            for each_cell in row:

                outfile_sheet.write_row(row_number_write, 0, row)#write relevant rows to output

            row_number_write = row_number_write + 1 # index and write only relevant (non-empty, non-header) rows

        else:
            print("header row")
    else:
        print("empty row")

    row_number = row_number + 1 #count which row we are reading from the input spreadsheet.


outfile_obj.close() #close output file

print('Script ends at: ' + str(datetime.datetime.now())) #print end time