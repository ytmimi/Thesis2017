#imports to get the files
import os
import sys
import openpyxl
import pandas as pd
from test_path import test_path, test_path2,test_path5, NextEra_test_path

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')

#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
#imported from one line becasue its path is already in sys.path
import generate_sorted_options_workbooks as gsow


#tests the convert_to_numbers() function
# letter_list = ['A', 'e', 'p', 'B', 'Z', 'AB']
# num_list=gsow.convert_to_numbers(letter_list)
# print(num_list)


#test the group_contracts_by_strik() function
test_wb = openpyxl.load_workbook(test_path)
contracts1 = gsow.group_contracts_by_strike(reference_wb = test_wb)
# print(contracts1['call'].keys())
# print(contracts1['put'].keys())
# print('\n')
# #prints a sample of the list values stored in each key 
# for (index, key) in enumerate(contracts1['call']):
# 	if index > 4:
# 		break
# 	print(key, len(contracts1['call'][key]))
# print('\n')
# for (index, key) in enumerate(contracts1['put']):
# 	if index > 4:
# 		break
# 	print(key, contracts1['put'][key])

contracts2 = gsow.group_contracts_by_expiration(reference_wb = test_wb)
print(list(contracts2['call'].keys()))
print(contracts2['put'].keys())
# print('\n')
# #prints a sample of the list values stored in each key
for (index, key) in enumerate(contracts2['call']):
	if index > 4:
		break
	print('Expiration date: {} Contracts sampled: {}'.format(key, len(contracts2['call'][key])))
print('\n')
# for (index, key) in enumerate(contracts2['put']):
# 	if index > 4:
# 		break
# 	print(key, contracts2['put'][key])


#test the create_sorted_workbooks function.
#NOTE: one of the helper functions for the given function is, generate_sorted_sheets(),
#and its fuction is to generate the formated sheets of the workbook
#this test also test that function.
# gsow.create_sorted_workbooks(reference_wb_path= test_path2, header_start_row=8,
# 						data_column=['C','E'], index_column=['A','B'],
# 						sort_by_strike=True, sort_by_expiration=True)

# gsow.create_sorted_workbooks(reference_wb_path= test_path2, header_start_row=8,
# 						data_column=['C','E'], index_column=['A','B'],
# 						sort_by_strike=True, sort_by_expiration=False)

# gsow.create_sorted_workbooks(reference_wb_path= test_path5, header_start_row=8,
# 						data_column=['C','E'], index_column=['A','B'],
# 						sort_by_strike=False, sort_by_expiration=True)

# gsow.create_sorted_workbooks(reference_wb_path= test_path2, header_start_row=8,
# 						data_column=['C','E'], index_column=['A','B'],
# 						sort_by_strike=False, sort_by_expiration=False)


#the sort_expiration_dates() function
# test_list = ['01-15-16', '09-18-15', '02-20-15', '03-20-15', '06-19-15','12-20-14', '01-17-15']
# print(test_list)
# expected_sort = ['12-20-14', '01-17-15', '02-20-15', '03-20-15', '06-19-15', '09-18-15', '01-15-16']
# sorted_list = gsow.sort_expiration_dates(test_list)
# print(sorted_list)
# print(expected_sort)

#the sort_strike_prices() function
# test_list = ['C22', 'C34', 'C5', 'C17', 'C50', 'C25', 'C3']
# print(test_list)
# expected_sort = ['C3','C5','C17','C22','C25','C34','C50']
# sorted_list = gsow.sort_strike_prices(test_list)
# print(sorted_list)
# print(expected_sort)






