# [description]
# This script formatted the raw dam register from the PROAGRO Inventario Nacional de Presas for Bolivia to GIS-importable format.
# Authors: Sam Karakas and Jida Wang
# Contact: skarakas21@g.ucla.edu
# Last update: 07/31/2022

# define input & output docs
infile = "Bolivia/bolivia_main_output_1.xlsx"
outfile = "Bolivia/bolivia_main_output_2.xlsx"

# import packages
import sys, datetime, xlsxwriter, openpyxl, re

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile, keep_vba=True)
#infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("dams_subset") #add worksheet to xlsx file

# create a header row from the main data:
outfile_sheet.write_row(0,0, tuple(['Code','Name','AreaQC','Type','Use','Area','Municipality','Height','Latitude','Longitude','Crown_length','Capacity','Watershed','Crown Dimension','River']))
# we will take 'code' and 'name' and 'AreaQC' from the summary file later

row_number = 0
dam_number = 0 # we can use this to track dam number since it is not equivalent to row number

output_list = [] # this list will hold 'output_row' lists to later be iterated into the output excel spreadsheet

# this list of header values allows us to keep the if statements more concise, while handling special headers separately
# 'tipo de presa' flags the start of a new dam table
# 'latitud\nLongitude' must be handled separately as it contains both values in one merged cell and we want them in separate columns
# 'Rio de la presa' represents the last relevant attr for each dam, so we will store our list after processing this value...
# and then reset the temporary 'output_row' list at the next hit on 'tipo de presa'
non_special_headers = ["Uso","Área de la cuenca","Municipio","Altura de la presa","Longitud coronamiento","Capacidad de embalse","Cuenca de influencia","Cota coronamiento"]

# this function takes the iterator and iterable of our for loop to select the next cell as we iterate
# we can flag the headers as they are consistent strings, but we need the actual attribute data from the immediate next cell
# within the if/elif statement within the loop, if a header cell is found then this function grabs the next cell
# and appends it to our running 'output_row' list
def append_output(a,b):
    try:
        output_row.append(a[a.index(b)+1]) # our if statements recognize cells with header information, by adding 1 to the index we skip these redundant headers and take the relevant attr data from the immediate next cell in the row
    except:
        # some dams are missing the attribute "cuenca de influencia" or watershed
        # this will throw an error without this try/except handling
#       print("EXPECTED CELL EMPTY AT " + str(row_number))
        output_row.append("n/a") # for these we will fill the cell as 'n/a'


for each_row in infile_sheet:

    # print('looping row#: ' + str(row_number)) # terminal QC
    
    row = []
    for cell in each_row:
        if cell.value:
            row.append(cell.value)
        # this stores the temporary values in a list, rather than the openpyxl object 'each_row'
        # 'each_row' does not have an index property that we can use to grab the next cell, rather than the flagged header cell

    for this_cell in row: # for loop through our list object of 'each_row'

        cell_value = str(this_cell) # make sure value is read as a string

        if cell_value == "Tipo de presa":
            output_row = [] # output_row will reset at the beginning of a new dam, flagged by header value 'Tipo de presa'
            append_output(row,cell_value) 

        elif cell_value in non_special_headers:
            append_output(row, cell_value)

        elif cell_value == "Latitud\nLongitud": #the lat/long attributes are stored in one excel cell with a line break
            # so they must be handled separately and split
            #print(cell_value)

            lat_long_str = row[row.index(cell_value) + 1] # same logic as append_output function for getting 'next' cell
            #print(repr(lat_long_str)) #repr is a built-in function which gives us the string literal, incl line break

            for i, char in enumerate(lat_long_str): # the lat/long cells are not all formatted exactly the same
                #so we can loop through the string to find the line break, then slice the string at the line break
                if char == '\n':
                    line_break = i

            # let's slice some strings
            lat = lat_long_str[:line_break]
            lat = lat[0:9] # this slices the string one more time, as some coords have an 'S' or an 'O' at the end
            long = lat_long_str[line_break + 1:]
            long = long[0:9]

            #and finallly append the output
            output_row.append(lat)
            output_row.append(long)

        elif cell_value == "Río de la presa":
            append_output(row, cell_value)
            output_list.append(output_row) # 'rio...' is the last relevant attr for each table, so our 'ouput_row' list will contain all relevant attributes at this point in the loop
            # so we will append that list (representing one dam's attributes) to 'output_list'
            #print(output_row) #terminal QC
            dam_number += 1 # increment dam number

    row_number += 1

row_number_write = 1 # set index at 1 to skip header row

for dam in output_list: # this for loop prints each dam's records out as one row in the output excel sheet
    outfile_sheet.write_row(row_number_write, 3, output_list[row_number_write - 1]) #start at column 4 to leave room for summary data
    row_number_write += 1

outfile_obj.close() # close output file

print('Script ends at: ' + str(datetime.datetime.now())) # print end time
print('# of dams processed: ' + str(dam_number) + ' out of 287') # we are looking for 287 dams total