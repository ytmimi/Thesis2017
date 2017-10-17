#imports to get the files
import os
import sys

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')

#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
import generate_company_workbooks as gcw


#harded coded, need to be flexible so these should be updated
sample_path = '/Users/yacintmimi/Documents/UVM/Fall 2017/ma_option_vol/company_data/sample/M&A List A-S&P500 T-US Sample Set.xlsx'
target_path = '/Users/yacintmimi/Documents/UVM/Fall 2017/ma_option_vol/company_data/target'
acquirer_path = '/Users/yacintmimi/Documents/UVM/Fall 2017/ma_option_vol/company_data/acquirer'


c = gcw.Create_Company_Workbooks(source_file = sample_path, target_path= target_path, acquirer_path = acquirer_path)
c.create_company_workbooks()

