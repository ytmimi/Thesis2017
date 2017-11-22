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
TREASURY_WORKBOOK_PATH = '{}/{}/{}/{}'.format(os.path.abspath(os.pardir), 'company_data','sample', 'Week_Day Treasury Rates.xlsx')

TREASURY_WORKSHEET= openpyxl.load_workbook(TREASURY_WORKBOOK_PATH, data_only=True).get_sheet_by_name('Rates')

TOTAL_TREASURY_SHEET_ROWS = TREASURY_WORKSHEET.max_row

TREASURY_DATA_START_ROW = 2

DATE_COLUMN= 2

THREE_MONTH_COLUMN= 7

SIX_MONTH_COLUMN= 8 

TWELVE_MONTH_COLUMN= 9


#total number of minutes in a 365 day year
MINUTES_PER_YEAR= 525600
#total number of minutes in a 30 day period
MINUTES_PER_MONTH= 43200




