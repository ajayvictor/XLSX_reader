import xlrd
from xlrd.sheet import ctype_text   


xl_workbook = xlrd.open_workbook("validation.xlsx")

sheet_names = xl_workbook.sheet_names()

print('Sheet Names', sheet_names)

xl_sheet = xl_workbook.sheet_by_name(sheet_names[3])

print(xl_sheet.name)

#first column
checker_column = 18

#comparison column with first column
tbm_column = 22

#value to be taken
og_value = 24

#starting row
starter_row = 5

y_value = 0

num_cols = xl_sheet.ncols   # Number of columns
for row_idx in range(starter_row, xl_sheet.nrows):    # Iterate through rows
    #print ('Row: %s' % row_idx)   # Print row number
    checker_value = xl_sheet.cell(row_idx, tbm_column)
    
    #print(checker_value.value)
    #print(xl_sheet.cell(row_idx, og_value))
    #print(xl_sheet.cell(row_idx, og_value).value)

    #for col_idx in columns_needed:  # Iterate through columns
    
    matching_value = xl_sheet.cell(row_idx, checker_column)

    checker_value = float(matching_value.value) - float(checker_value.value)
    for row_idx_d1 in range(row_idx, xl_sheet.nrows):    # Iterate through rows
        cell_obj = xl_sheet.cell(row_idx_d1, tbm_column)  # Get cell object by row, col
        #cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
        #print("Next One: ", cell_type_str)
        # print("Printing Cell Object Here: ", cell_obj.value, row_idx_d1)
        if cell_obj.value == "" or matching_value.value == "":
          #print("break happened with values", cell_obj.value, " and ", matching_value.value)
          break
        #value = float(matching_value.value) - float(cell_obj.value)
        value = abs(float(cell_obj.value) - float(matching_value.value))
        #print("matching value: ", matching_value.value, "checker value: ", checker_value)
        #print(float(matching_value.value) - float(cell_obj.value))
        #print("value is ", value)
        #print("checker_value is ", checker_value)
        if value <= abs(checker_value):
                y_value = xl_sheet.cell(row_idx_d1, og_value)
                checker_value = value
                #checker_value = cell_obj.value
                #print("y value is changing: ", y_value, checker_value)
        #print("y value is ====> ", y_value)
        #print(float(y_value.value))
        #print(matching_value.value)
    #print("Columns Matching: DEYANOVA X--> " + str(matching_value.value) + " PRESENT ANALYSIS Y--> " + str(y_value.value))
    print(str(y_value.value))
