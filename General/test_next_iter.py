infile = r'D:\Research\Projects\SWOT\Dam_inventory_collection\Dam_harmonization\Codes\test_table.xlsx'


import urllib, csv, openpyxl, unicodedata, re

# open infile
infile_obj = openpyxl.load_workbook(infile) 
infile_sheet = infile_obj.active

for each_row in infile_sheet:

    # JW: WE NEED TO FIRST DEFINE THE ITERATOR. THE FIRST TIME TO USE 'NEXT' WILL ALWAYS START FROM THE 0TH INDEX.
    each_row_iterator = iter(each_row)

    for cell in each_row:

        cell_value = str(cell.value) # read cell value as string
        code_check = re.search(r"[A-Z][A-Z]\-[A-Z]\-\d\d\d", cell_value) # reg ex to check for 'code' values which should have a format like 'AB-X-123'

        # JW: this is equivalent to cell_value.
        this_cell = next(each_row_iterator).value
        
        if code_check: # check each cell for this regex
        
            output_row = [cell.value, next(each_row_iterator).value] #JW: here the next() will read the next element. 
            
            print(output_row)#check in terminal to see if working

            break



# JW: Alternatively, we can do the following:
print('Alternatively...')

for each_row in infile_sheet:

    row = []
    #loop to iterate through cells
    for cell in each_row:
        # check for and skip empty cells
        if cell.value:
            row.append(cell.value) # pop cell w value
            
    for this_cell in row:

        cell_value = str(this_cell)
        code_check = re.search(r"[A-Z][A-Z]\-[A-Z]\-\d\d\d", cell_value)
        
        if code_check: # check each cell for this regex
        
            output_row = [cell_value, row[row.index(cell_value) + 1]] #JW: read the value at the next index
            
            print(output_row)#check in terminal to see if working

infile_obj.close()
