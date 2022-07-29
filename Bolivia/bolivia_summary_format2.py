infile = "bolivia_output_1.xlsx"
outfile = "bolivia_output_2_summary.xlsx"

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


#for loop to read cleaned outfile
for each_row in outfile_sheet:

    print('looping row#: ' + str(row_number))

    row = []

    for cell in each_row:

        if cell.value:

            cell_value = str(cell.value)
            code_check = re.search(r"[A-Z][A-Z][A-Z]\-[A-Z]\-\d\d\d", cell_value)
       
            if code_check:
                print(str(cell.value))

                output_row = [cell.value, next(iter(each_row)).value]# i don't think this is going to work while there are empty cells resulting from column merges in the document

                outfile_sheet.write_row(row_number_write, 0, output_row)

    row_number = row_number + 1
    row_number_write = row_number_write + 1