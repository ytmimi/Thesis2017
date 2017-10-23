#imports to get the files
import os
import sys
import openpyxl
import pandas as pd
from test_path import test_path, test_path2

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')

#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
#imported from one line becasue its path is already in sys.path
import generate_sorted_options_workbooks as gsow


#tests the convert_to_numbers() function
letter_list = ['A', 'e', 'p', 'B', 'Z', 'AB']
num_list=gsow.convert_to_numbers(letter_list)
print(num_list)


#test the group_contracts_by_strik() function
test_wb = openpyxl.load_workbook(test_path)
contracts = gsow.group_contracts_by_strike(wb = test_wb)
print(contracts['call'].keys())
print(contracts['put'].keys())
print('\n')
#prints a sample of the list values stored in each key 
for (index, key) in enumerate(contracts['call']):
	if index > 4:
		break
	print(key, contracts['call'][key])
print('\n')
for (index, key) in enumerate(contracts['put']):
	if index > 4:
		break
	print(key, contracts['put'][key])


#test the create_sorted_workbooks function.
#NOTE: one of the helper functions for the given function is, generate_sorted_sheets(),
#and its fuction is to generate the formated sheets of the workbook
#this test also test the other function.
gsow.create_sorted_workbooks(reference_wb_path= test_path2,
						data_start_row= 8, 
						data_column=['C'], 
						index_column=['A','B'])






