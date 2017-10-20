#imports to get the files
import os
import sys

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')

#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
#imported from one line becasue its path is already in sys.path
import add_bloomberg_excel_functions as abxl

	
'''
add_BDS_OPT_CHAIN() EXAMPLE

Creates a string representing the Bloomberg BDS function for options chains to be used in an excel worksheet.
Note: the excel document needs to be run in order for the formulas to be calculated

Example: add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B6')
         function arguments are string cell references
         
Example: add_BDS_OPT_CHAIN(ticker_cell='AAPL US',type_cell='EQUITY', date_override_cell='20151231')
         function arguments are literal string commands excepted by the BDS function
'''

#test the function with both cell references and valid Bloomberg strings
print('\n')
print(abxl.add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B6'))
print(abxl.add_BDS_OPT_CHAIN(ticker_cell='AAPL US',type_cell='EQUITY', date_override_cell='20151231'))


'''
add_BDP_fuction() EXAMPLE

Creates a string representing the Bloomberg BDP function to be used in an excel worksheet.
Note: the excel document needs to be run in order for the formulas to be calculated

Example: add_BDP_fuction(security_cell='A3', field_cell='A4')
		 function arguments are string cell references

Example: add_BDP_fuction(security_cell='BBG00673J6L5 Equity', field_cell='SECURITY_DES')
    	 function arguments are literal string commands excepted by the BDS function
'''
#test the function with both cell references and valid Bloomberg strings
print('\n')
print(abxl.add_BDP_fuction(security_cell='A3', field_cell='A4'))
print(abxl.add_BDP_fuction(security_cell='BBG00673J6L5 Equity', field_cell='SECURITY_DES'))
print(abxl.add_BDP_fuction(security_cell='A3', field_cell='PX_LAST'))



'''
add_option_BDH() EXAMPLE

Creates a string representing the Bloomberg BDH function to be used in an excel worksheet.
Note: the excle document needs to be run in order for the formulas to be calculated

Example: add_option_BDH(security_name='B4', fields='C3', start_date='A1', end_date='A2', optional_arg = None, optional_val=None)
		 function arguments are string cell references 

Example: add_option_BDH(security_name='B4', fields='C3:C5', start_date='A1', end_date='A2', optional_arg = 'E1:E2', optional_val='F1:F2')
		 function arguments are a combination of cells and cell references

Example: add_option_BDH(security_name='IBM US EQUITY', fields='PX_BID, PX_ASK, PX_VOLUME', start_date='20150520', end_date='20170915', optional_arg = 'Days', optional_val='W')
		 function arguments are strings

Example: add_option_BDH(security_name='B4', fields=['PX_LAST','PX_BID'], start_date='20131230', end_date='A2', optional_arg = 'Days, Fill', optional_val=['W','0'])
		 function argumetns are strings, cell references and lists

'''
print('\n')
print(abxl.add_option_BDH(security_name='B4', fields='C3', start_date='A1', end_date='A2', optional_arg = None, optional_val=None))
print(abxl.add_option_BDH(security_name='B4', fields='C3:C5', start_date='A1', end_date='A2', optional_arg = 'E1:E2', optional_val='F1:F2'))
print(abxl.add_option_BDH(security_name='IBM US EQUITY', fields='PX_BID, PX_ASK, PX_VOLUME', start_date='20150520', end_date='20170915', optional_arg = 'Days', optional_val='W'))
print(abxl.add_option_BDH(security_name='B4', fields=['PX_LAST','PX_BID'], start_date='20131230', end_date='A2', optional_arg = 'Days, Fill', optional_val=['W','0']))





