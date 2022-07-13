# Sample code by J Wang
# Function: 
# 1. iterate through every dam record (row) in the input spreadsheet,
# 2. write the subset of the dam records whose storage capacity > 1 cmc (cubic m3) into a new excel spreadsheet, and
# 3. in the output spreadsheet, only keep the attributes of "dam_name", "capacity_mcm", "Latitude_dec" and "Longitude_dec".

#======================================================================================
# INPUT
# Define the input excel
infile = r"D:\Research\Projects\SWOT\Dam_inventory_collection\sample_codes\sample_dams.xlsx"

# OUTPUT
# Define the output excel
outfile = r"D:\Research\Projects\SWOT\Dam_inventory_collection\sample_codes\large_dams.xlsx"
#======================================================================================


#======================================================================================
# CODE
# Note: you will need to first install xlsxwriter and openpyxl to Python.
# see: https://pypi.org/project/openpyxl; https://xlsxwriter.readthedocs.io/getting_started.html

# Important packages and functions
import sys, datetime, xlsxwriter, openpyxl

# Print script starting time
print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("dams_subset") #any name you want to give to the output spreadsheet
          
# Loop through each record in the intput spreadsheet
row_number = 0 #initiate an integer to indicate the row index in the input spreadsheet.
row_number_write = 0 #initiate an integer to indicate the row index in the output spreadsheet. 
for each_row in infile_sheet:
    print('looping row#: ' + str(row_number))

    #Read this row
    row = []
    for cell in each_row:
        row.append(cell.value) 
    
    if row_number == 0: #the first row is the header (attribute row)      
    
        #define the output header
        #'dam_name', 'capacity_mcm', 'Latitude_dec', 'Longitude_dec' correspond to the 1st, 6th, 7th, and 8th column in the input 
        output_header = [row[0], row[5], row[6], row[7]] #Note: Python is zero indexed (i.e., starting from 0 not 1). 
        #Alternative way, we can define the output header by simply typing the attribute names as below
        #output_header = tuple(['dam_name', 'capacity_mcm', 'Latitude_dec', 'Longitude_dec'])
        
        #Write the defined header to the output excel 
        outfile_sheet.write_row(row_number_write, 0, output_header) #see more: https://xlsxwriter.readthedocs.io/worksheet.html#write_row
        row_number_write = row_number_write + 1 #preparation: switching to the next row to write the next record later
    
    else: #records below the header
        
        #retrieve the storage capacity of this record 
        this_capacity = row[5] #the 6th column
        
        #check if this storage capacity exceeds 1.0
        if this_capacity > 1: #if so, write this record to the output spreadsheet
            outfile_sheet.write_row(row_number_write, 0, tuple([row[0], row[5], row[6], row[7]]))
            row_number_write = row_number_write + 1 #preparation: switching to the next row to write the next record later
            
    row_number = row_number + 1 #count which row we are reading from the input spreadsheet.

outfile_obj.close() #we need to close the output file in order to save what we have written
    
# Print script ending time    
print('Script ends at: ' + str(datetime.datetime.now()))
