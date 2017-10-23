#imports to get the files
import os
import sys
import pandas as pd
from test_path import test_path

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

#update the index for each sheet with data in the workbook
#uxlw.update_workbook_data_index(workbook_path =test_path)


