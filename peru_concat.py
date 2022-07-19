#define input & output docs
infile = "Peru_test_output.xlsx"
outfile = "peru_concat.xlsx"

import sys, datetime, xlsxwriter, openpyxl # import libraries

print('Script starts at: ' + str(datetime.datetime.now())) # print script start time

infile_obj = openpyxl.load_workbook(infile) # open input xlsx
infile_sheet = infile_obj.active # select active sheet

outfile_obj = xlsxwriter.Workbook(outfile) # create empty output xlsx
outfile_sheet = outfile_obj.add_worksheet("peru_dams_concat") # add worksheet to xlsx file

# initialize index variables for read and write
row_number = 0
row_number_write = 0

def concatenate(row_range): # function to concatenate multiple cell values in a row into one
    metadata_string = str('dam_name: ' + str(row_range[0]) + ' admin_auth: ' + str(row_range[1]) + ' local_admin: ' + str(row_range[2]) + ' river: ' + str(row_range[3]) + ' district: ' + str(row_range[4]) + ' province: ' + str(row_range[5]))
    return metadata_string


# for loop to read infile
for each_row in infile_sheet:

    print('looping row#: ' + str(row_number)) # print statement to track running script

    row = [] # create list object for spreadsheet rows
    row_range = [] # list for values to concatenate

    cell_number = 0

    #reg_string = concatenate(row[3],row[4],row[5],row[6],row[7],row[8]) # send row values as argument to concatenate function to return a concatenated string
    
    if row_number == 0: # index 0 = header row
        output_header = tuple(['reg_code', 'east', 'north', 'reg_string'])
        outfile_sheet.write_row(row_number_write, 0, output_header)
        row_number_write += 1

    else: # records below header
        for cell in each_row:
            if cell_number < 3:
                row.append(cell.value) # populate row list with input cell values
            else:
                row_range.append(cell.value) # populate list to concatenate
            cell_number += 1
            row.append(concatenate(row_range))

        outfile_sheet.write_row(row_number_write, 0, tuple([row[0],row[1],row[2],row[3]]))
        row_number_write += 1

    row_number += 1

outfile_obj.close() # close and save output file

# Print script ending time    
print('Script ends at: ' + str(datetime.datetime.now()))


