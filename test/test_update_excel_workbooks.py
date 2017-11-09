#imports to get the files
import os
import sys
import datetime as dt
import openpyxl
import pandas as pd
from test_path import test_path,test_path2, test_path3, NextEra_test_path, test_stock_price, Allegran_path

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')


#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
#imported from one line becasue its path is already in sys.path
import update_excel_workbooks as uxlw

#update_sheet with the BDP function
#uxlw.update_sheet_with_BDP_description(workbook_path =test_path)

#test the update_options_contract_sheets function.
#adds a new sheet for each option contract listed in the Options Chain sheet and pulls bloomberg data for each field listed in 

# uxlw.update_option_contract_sheets(workbook_path=test_path, 
# 									sheet_name='Options Chain',
# 									sheet_start_date_cell='B7',
# 									sheet_end_date_cell='B8',
# 									data_header_row=8,
# 									data_table_index=['INDEX','DATE'],
# 									data_table_header=['PX_LAST','PX_BID','PX_ASK','PX_VOLUME','OPEN_INT', 'IVOL'],
# 									BDH_optional_arg=['Days', 'Fill'],
# 									BDH_optional_val=['T','0'])

#deletes all worksheets in the workbook except the first worksheet
#uxlw.delet_workbook_sheets(NextEra_test_path) #change back to test_path after testing NextEra sheet

#update the index for each sheet in relation to the announcement date
uxlw.update_workbook_data_index(workbook_path =test_path, data_start_row=9, index_column='A')

#test the find_column_index_by_header() function
# wb = openpyxl.load_workbook(test_path3)
# data_dict= uxlw.find_column_index_by_header(reference_wb =wb, column_header='PX_LAST', header_row=8)
# for index, key in enumerate(data_dict):
# 	print(key, data_dict[key])

#test the update_workbook_average_column() function
#uxlw.update_workbook_average_column(reference_wb_path = test_path3, column_header='PX_LAST', header_row=8, data_start_row=9, ignore_sheet_list=['Stock Price'])


#test the update_stock_price_sheet()
'''
uxlw.update_stock_price_sheet(	workbook_path =NextEra_test_path, #change back to test_path after testing NextEra sheet
							sheet_name='Options Chain',
							stock_sheet_index = 1,
							sheet_start_date_cell='B7',
							sheet_end_date_cell='B8',  
							data_header_row=8, 
							data_table_index=['INDEX','DATE'], 
							data_table_header=['PX_LAST'], 
							BDH_optional_arg=None, 
							BDH_optional_val=None )
'''

#test
# stock_data = uxlw.stock_data_to_list(reference_wb_path=test_path2, price_column_header='PX_LAST', header_start_row=8)
# print(stock_data)
# average=uxlw.stock_price_average(reference_wb_path=test_path2, price_column_header='PX_LAST', header_start_row=8)
# print(average)

# st_dev=uxlw.stock_price_standard_dev(reference_wb_path=test_stock_price, price_column_header='PX_LAST', header_start_row=8)
# print(st_dev)

# print(average+st_dev)
# print(average-st_dev)

#test mean and std functions:
# wb = openpyxl.load_workbook(Allegran_path)
# sheet = wb.get_sheet_by_name('Options Chain')
# announcement_date =dt.datetime.strptime(str(sheet['B8'].value),'%Y%m%d')
# print(announcement_date)

# hm_std=uxlw.historic_stock_mean_and_std(reference_wb_path=Allegran_path, price_column_header='PX_LAST', header_start_row=8, date_0=announcement_date)
# print(hm_std)
# print('\n')
# mm_std=uxlw.merger_stock_mean_and_std(reference_wb_path=Allegran_path, price_column_header='PX_LAST', header_start_row=8, date_0=announcement_date)
# print(mm_std)

# #test the in_range() function
# print(uxlw.is_in_range(num=25, high=hm_std[0]+hm_std[1], low=hm_std[0]-hm_std[1]))
# print(uxlw.is_in_range(num=hm_std[0], high=hm_std[0]+hm_std[1], low=hm_std[0]-hm_std[1]))


# uxlw.fill_option_wb_empty_cells(reference_wb_path=Allegran_path, column_start=3, row_start=9, fill_value=0)






