import re
import openpyxl
import sys
import os



#a regular expression for a formated option description where the strike is an integer
OPTION_SHEET_PATTERN_INT = re.compile(r'^\w+\s\w+\s\d{2}-\d{2}-\d{2}\s\w+$')

#a regular expression for a formated option description where the strike is a foat
OPTION_SHEET_PATTERN_FLOAT= re.compile(r'^\w+\s\w+\s\d{2}-\d{2}-\d{2}\s\w+\.\w+$')

#a regular expression pattern for the stock sheet
STOCK_SHEET_PATTERN =re.compile(r'^\w+\s\w+\s\w+$')

#a regular expression for a formated option description where the strike is an integer
OPTION_DESCRIPTION_PATTERN_INT= re.compile(r'^\w+\s\w+\s\d{2}/\d{2}/\d{2}\s\w+$')

#a regular expression for a formated option description where the strike is a foat
OPTION_DESCRIPTION_PATTERN_FLOAT = re.compile(r'^\w+\s\w+\s\d{2}/\d{2}/\d{2}\s\w+\.\w+$')

#a regular expression to designate whether an option description is a Call
CALL_DESIGNATION_PATTERN = re.compile(r'[C]\d+')

#a regular expression to designate whether an option description is a Put
PUT_DESIGNATION_PATTERN = re.compile(r'[P]\d+')






#USEFUL INFORMATINO FROM THE 'Treasury Rates.xlsx' file
TREASURY_WORKBOOK_PATH = '{}/{}/{}/{}'.format(os.path.abspath(os.pardir), 'company_data','sample', 'Treasury Rates.xlsx')
TREASURY_WORKSHEET= openpyxl.load_workbook(TREASURY_WORKBOOK_PATH).get_sheet_by_name('Intrest Rates')
TOTAL_TREASURY_SHEET_ROWS = TREASURY_WORKSHEET.max_row
DATE_COLUMN= 1
THREE_MONTH_COLUMN= 2
SIX_MONTH_COLUMN= 3 
TWELVE_MONTH_COLUMN= 4
