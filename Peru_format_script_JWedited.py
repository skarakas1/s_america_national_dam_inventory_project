# [Description] ------------------------------
# This script formatted the raw dam register from Autoridad Nacional del Agua for Peru to GIS-importable format.
# Authors: xxx, xxx, xxx, and Jida Wang
# Contact: xxx@g.ucla.edu
# Last update: 07/26/2022

# Reference:
# Autoridad Nacional del Agua: Inventario de Presas en el Peru, Lima,
# available at: https://www.ana.gob.pe/etiquetas/inventario-de-presas (last access: 11 July 2022), 2016. 
#---------------------------------------------


# JW:
# There are two header formats in the original PDF:

# header format 1: (in most tables)
# 'Código','Este','Norte','Nombre de la presa','Autoridad Administrativa del Agua','Administración Local del Agua','Río / Quebrada','Distrito','Provincia'				

# header format 2: (in a few tables)
# 'N°','Código','Este','Norte','Nombre de la presa','Autoridad Administrativa del Agua','Administración Local del Agua,'Río / Quebrada','Unidad Hidrográfica', ...
# ... 'Distrito', 'Provincia', 'Dpto.'					

# In the final cleaned-up table, we can perhaps drop 'N°' and 'Dpto.' but keep 'Unidad Hidrográfica' (hydrographic unit, if available).

# Also, I noted some of the dam "codes" (IDs) start with '0' (e.g., '07007'). If we read/write them directly to excel, they will be treated as integers, ...
# ... and the initila '0' will be dropped (even if we converted the values to text). This is a little tricky. So I modified the script to make sure the ...
# ... leading 0 will NOT be dropped. 

# !! Another note: in the input excel (and the original PDF), there are a few pairs of dams whose records are stacked on each other in ONE row. 
# Cases include: dam IDs 59002 and 59004; IDs 61004 and 61005; IDs 69009 and 69011.

#^^^ done, see 'Peru_pdf_clip_QCed.xlsx - SK

# Please clean up the input excel (breaking them to different rows) before running the script. 
# The number of the final cleaned-up dams should be 743 (excluding the header).  

# Please see my edits/comments below (the lines starting with "JW").


# Define the input & output excel docs
infile = r"D:\Research\Projects\SWOT\Dam_inventory_collection\Dam_harmonization\Codes\Peru_pdf_clip.xlsx"
outfile = r"D:\Research\Projects\SWOT\Dam_inventory_collection\Dam_harmonization\Codes\Peru_formatted.xlsx"

#these can be edited back in if using GitHub project folder in VSCode
#infile = "Peru_pdf_clip_QCed.xlsx"
#outfile = "Peru_test_output.xlsx"

# Important packages and functions
import sys, datetime, xlsxwriter, openpyxl

print('Script starts at: ' + str(datetime.datetime.now()))

# Open the input spreadsheet
infile_obj = openpyxl.load_workbook(infile, keep_vba=True)
#infile_obj = openpyxl.load_workbook(infile) #see more: https://openpyxl.readthedocs.io/en/stable/
infile_sheet = infile_obj.active #active spreadsheet of this excel

# Create an empty output excel spreadsheet
outfile_obj = xlsxwriter.Workbook(outfile) #see more: https://xlsxwriter.readthedocs.io/
outfile_sheet = outfile_obj.add_worksheet("dams_subset") #add worksheet to xlsx file

# JW : This block (for writing the header) can be simplified to: 
outfile_sheet.write_row(0,0, tuple(['Code','East','North','Dam_name','Administrative_authority_of_water',\
                                    'Local_water_administration','River','District','Province','Hydrographic_unit',\
                                    'UTM_zone'])) #JW: Please manually add the UTM zone in the output excel. 
#outfile_sheet.write(0,0, 'Code') 
#outfile_sheet.write(0,1, 'East')
#outfile_sheet.write(0,2, 'North')
#outfile_sheet.write(0,3, 'Dam Name')
#outfile_sheet.write(0,4, 'Administrative Authority')
#outfile_sheet.write(0,5, 'Local Administration')
#outfile_sheet.write(0,6, 'River')
#outfile_sheet.write(0,7, 'District')
#outfile_sheet.write(0,8, 'Province')


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
        
        # JW: sometimes the first header can be "N°" (header format 2), so I added another logical statement here. 
        if str(row[0]) != "Código" and str(row[0]) != "N°": #this ignores header rows

            # JW: This FOR loop is not needed. By executing this loop with the same "row_number_write", ...
            # ... this same row will actually be overwritten multiple times (size(each_cell) times). 
            #for each_cell in row:
            #    outfile_sheet.write_row(row_number_write, 0, row)#write relevant rows to output
            #row_number_write = row_number_write + 1      
              
            # JW: Added discussion to deal with different header formats. 
            if len(row) == 9: #header format 1 (see comments above), which is the norm. 
                # header format 1:
                # 'Código','Este','Norte','Nombre de la presa','Autoridad Administrativa del Agua','Administración Local del Agua', ...
                # ... 'Río / Quebrada','Distrito','Provincia'	
                rearranged_row = row                        
                
            elif len(row) == 12: #header format 2, which is rare.
                # header format 2:
                # 'N°','Código','Este','Norte','Nombre de la presa','Autoridad Administrativa del Agua','Administración Local del Agua, ...
                # ... 'Río / Quebrada','Unidad Hidrográfica', 'Distrito', 'Provincia', 'Dpto.'	            
                # rearrange the row to drop 'N°' and 'Dpto' and relocate 'Unidad Hidrográfica' (hydrographic unit) to the end  
                rearranged_row = [row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[9],row[10],row[8]]
        
            else: #to verify this situation won't happen. 
                print("qualtiy control: this situation should not happen...........") #This has been verified. 
            
            # JW: To make sure the leading '0' is not dropped in the dam code/ID (which should always have 5 digits)
            if len(str(rearranged_row[0])) < 5: #meaning the leading 0 in the dam ID has been dropped after reading it. 
                rearranged_row[0] = '0' + str(rearranged_row[0]) #then add the leading 0 back and save the updated ID in a text format. 
            else:
                rearranged_row[0] = str(rearranged_row[0]) #convert the dam ID to text format
            
            # JW: write the re-arranged row to the output excel
            outfile_sheet.write_row(row_number_write, 0, rearranged_row)             
            row_number_write = row_number_write + 1
            
        #else:
        #    print("header row")
    #else:
        #print("empty row")

    row_number = row_number + 1 #count which row we are reading from the input spreadsheet.


outfile_obj.close() #close output file

print('Script ends at: ' + str(datetime.datetime.now())) #print end time
