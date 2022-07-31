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

# create a header row from the main data:
outfile_sheet.write_row(0,0, tuple(['Code','Name','AreaQC','Type','Use','Area','Municipality','Height','Latitude','Longitude','Crown_length','Capacity','Watershed','Crown Dimension','River']))
# we will take 'code' and 'name' and 'AreaQC' from the summary file later

row_number = 0
output_list = []
def append_output(a,b):
    try:
        output_row.append(a[a.index(b)+1])
    except:
        print("ERROR: EXPECTED CELL EMPTY")
        output_row.append("n/a")
non_special_headers = ["Uso","Área de la cuenca","Municipio","Altura de la presa","Longitud coronamiento","Capacidad de embalse","Cuenca de influencia","Cota coronamiento"]


for each_row in infile_sheet:

    #each_row_iterator = iter(each_row) # define iterator in order to use 'next' method in our for cell in each_row loop (called in append_output function)
    print('looping row#: ' + str(row_number))
    #flag = re.search("^Comentario", )
    # !!! ^this needs to be filled in
    #row = []
    row = []
    for cell in each_row:
        if cell.value:
            row.append(cell.value)
    for this_cell in row:

        cell_value = str(this_cell)

        if cell_value == "Tipo de presa":
            output_row = []
            #output_row.append(row[row.index(cell_value)+1])
            append_output(row,cell_value)
        elif cell_value in non_special_headers:
            append_output(row, cell_value)
        #elif cell_value == "Latitud\nLongitud":
            #special case will need to parse string as this is a double row...
        elif cell_value == "Río de la presa":
            append_output(row, cell_value)
            output_list.append(output_row)
            print(output_row)
    row_number += 1

        #here we will write our output rows from the 'output_list' """

print(output_list)
row_number_write = 0
for dam in output_list:
    # we can add in code and name here, grabbing them from the summary excel file
    outfile_sheet.write_row(row_number_write + 1, 3, output_list[row_number_write])
    row_number_write += 1

outfile_obj.close() #close output file

print('Script ends at: ' + str(datetime.datetime.now())) #print end time
    






