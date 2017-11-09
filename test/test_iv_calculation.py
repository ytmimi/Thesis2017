#imports to get the files
import os
import sys
import datetime as dt
import openpyxl
from test_path import test_path,test_path2, test_path3,test_path4, NextEra_test_path, test_stock_price, Allegran_path

parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')


#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

#imports the add_bloomber_excel_functions module
#imported from one line becasue its path is already in sys.path
import iv_calculation as ivc

#test the ivc.find_starting_risk_free_rate_index() function
# date = dt.datetime(year=2013, month=7, day=15)
# print(ivc.find_starting_risk_free_rate_index(start_date=date, data_start_row=9))


# expiration_date = dt.datetime(year=2013,month=7,day=20)
# #test the days_till_expiration(start_date, expiration_date)
# print(ivc. days_till_expiration(start_date=date, expiration_date=expiration_date))

#test ivc.is_negative(num)
# print(ivc.is_negative(num=-1))


#test ivc.calculate_sheet_iv()
# wb = openpyxl.load_workbook(test_path4)
# stock_sheet = wb.get_sheet_by_name('PFE US EQUITY')
# option_sheet = wb.get_sheet_by_name('PFE US 12-20-14 C26')
# ivc.calculate_sheet_iv(stock_sheet=stock_sheet, option_sheet=option_sheet,sheet_date_column=2,
# 	sheet_price_column=3,data_start_row=9, three_month=True, six_month=False, twelve_month=False)


#test ivc.calculate_workbook_iv()
ivc.calculate_workbook_iv(workbook_path= test_path4,sheet_date_column=2,sheet_price_column=3,
	data_start_row=9,three_month=True, six_month=True, twelve_month=True)
