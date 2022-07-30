# [description]
# This script formatted the raw dam register from the PROAGRO Inventario Nacional de Presas for Bolivia to GIS-importable format.
# Authors: Sam Karakas and Jida Wang
# Contact: skarakas21@g.ucla.edu
# Last update: 07/29/2022

# define input & output docs
infile = "Bolivia/bolivia_main_output_1.xlsx"
outfile = "Bolivia/bolivia_main_output_2.xlsx"
summary = "Bolivia/bolivia_output_2_summary.xlsx"
# !!!! still debugging bolivia_summary_format2.py which outputs the file we are calling for 'summary'

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

# create a header row
# from the main data, the columns will be:
# type, area, height, crown length, capacity, crown dimension (?), use, municipality, lat, long, watershed, river
outfile_sheet.write_row(0,0, tuple(['Code','Name','Type','Use','Area','Municipality','Height','Latitude','Longitude','Crown_length','Capacity','Watershed','Crown Dimension','River']))
# we will take 'code' and 'name' from the summary file later

row_number = 0
output_row = []
output_list = []

for each_row in infile:

    print('looping row#: ' + str(row_number))
    flag = re.search("^Comentario", )
    # !!! ^this needs to be filled in
    row = []

    for cell in each_row:
        if str(cell.value) == "Tipo de presa":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Uso":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Área de la cuenca":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Municipio":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Altura de la presa":
            output_row.append(each_row(next()).value)
        # elif str(cell.value) == "Latitud\nLongitud":
            #special case will need to parse string as this is a double row...
        elif str(cell.value) == "Longitud coronamiento":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Capacidad de embalse":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Cuenca de influencia":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Cota coronamiento":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Río de la presa":
            output_row.append(each_row(next()).value)
        #here we will check for a string starting with Comentarios: using regex
        elif re.search("^Comentario", str(cell.value)):
            output_list.append(output_row)
            output_row = [] # reset output_row
        #this will tell us to move to the next dam
        #then we will add our list, 'output_row' to the list 'output_list', storing the relevant values of each dam as a list
    row_number += 1



        #here we will write our output rows from the 'output_list'
for dam in output_list:
    row_number_write = 0
    # we can add in code and name here, grabbing them from the summary excel file
    outfile_sheet.write_row(row_number_write, 2, row)
    row_number_write += 1

outfile_obj.close() #close output file

print('Script ends at: ' + str(datetime.datetime.now())) #print end time
    






