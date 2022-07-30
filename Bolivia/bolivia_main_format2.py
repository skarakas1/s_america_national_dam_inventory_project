# [description]
# This script formatted the raw dam register from the PROAGRO Inventario Nacional de Presas for Bolivia to GIS-importable format.
# Authors: Sam Karakas and Jida Wang
# Contact: skarakas21@g.ucla.edu
# Last update: 07/29/2022

# define input & output docs
infile = "Bolivia/bolivia_main_output_1.xlsx"
outfile = "Bolivia/bolivia_main_output_2.xlsx"

# import packages
import sys, datetime, xlsxwriter, openpyxl

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile, keep_vba=True)
#infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("dams_subset") #add worksheet to xlsx file

# create a header row
# we will take 'code' and 'name from the summary file later
# from the main data, the columns will be:
# type, area, height, crown length, capacity, crown dimension (?), use, municipality, lat, long, watershed, river

output_list = []

for each_row in infile:

    row = []
    output_row = []

    for cell in each_row:
        if str(cell.value) == "Tipo de presa":
            output_row.append(each_row(next()).value)
        elif str(cell.value) == "Uso":
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
        #this will tell us to move to the next dam
            #then, if so, we will add our list, 'output_row' to the list 'output_list'
            #then, we will iterate through output list to write the output excel
        


