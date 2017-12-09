import re
import openpyxl
import sys
import os


#Root path to all the files
ROOT_PATH = os.path.abspath(os.pardir)
#path to acquirer folder
ACQUIRER_DIR = os.path.join(ROOT_PATH,'company_data', 'acquirer')

#path to the target folder
TARGET_DIR = os.path.join(ROOT_PATH,'company_data', 'target')

#path to the output folder
OUTPUT_DIR = os.path.join(ROOT_PATH, 'company_data', 'function_output')

#path to the sample folder
SAMPLE_DIR = os.path.join(ROOT_PATH, 'company_data', 'sample')

#sample file path
MERGER_SAMPLE = os.path.join(SAMPLE_DIR,'M&A List A-S&P500 T-US Sample Set.xlsx')


#a regular expression for a formated option description where the strike is an integer
OPTION_SHEET_PATTERN_INT = re.compile(r'^\w+\s\w+\s\d{2}-\d{2}-\d{2}\s\w+$')

#a regular expression for a formated option description where the strike is a float
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



#usefull information for the treasury sheet
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




