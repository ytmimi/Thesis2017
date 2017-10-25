#imports to get the files
import os
import sys
import openpyxl
import pandas as pd
from test_path import test_path, test_path3

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')

#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
#imported from one line becasue its path is already in sys.path
import update_excel_workbooks as uxlw

#update_sheet with the BDP function
#uxlw.update_sheet_with_BDP_description(workbook_path =test_path)

#adds a new sheet for each option contract listed in the Options Chain sheet
#uxlw.update_option_contract_sheets(workbook_path=test_path, sheet_name='Options Chain', sheet_end_date_cell='B6')

#deletes all worksheets in the workbook except the first worksheet
#uxlw.delet_workbook_sheets(test_path)

#update the index for each sheet in relation to the announcement date
#uxlw.update_workbook_data_index(workbook_path =test_path)

#test the find_column_index_by_header() function
# wb = openpyxl.load_workbook(test_path3)
# data_dict= uxlw.find_column_index_by_header(reference_wb =wb, column_header='PX_LAST', header_row=8)
# for index, key in enumerate(data_dict):
# 	print(key, data_dict[key])


#test the update_workbook_average_column() function
uxlw.update_workbook_average_column(reference_wb_path = test_path3, column_header='PX_LAST', header_row=8, data_start_row=9)