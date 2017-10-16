#imports to get the files
import os
import sys

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')

#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
from ma_option_vol import add_bloomberg_excel_functions as abxl


'''
add_BDS_OPT_CHAIN() EXAMPLE

creates a string representing the Bloomberg BDS function for options chains to be inserted into an excel WorkSheets Cell
Note: the document needs to be run in order for the formulas to be calculated

Example: add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B6')
         Gets values from the cell cooridinates passed to the function
         
Example: add_BDS_OPT_CHAIN(ticker_cell='AAPL US',type_cell='EQUITY', date_override_cell='20151231')
         Gets data from bloomberg based on the literal commands excepted by the BDS function
'''

print(abxl.add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B6'))
print(abxl.add_BDS_OPT_CHAIN(ticker_cell='AAPL US',type_cell='EQUITY', date_override_cell='20151231'))

