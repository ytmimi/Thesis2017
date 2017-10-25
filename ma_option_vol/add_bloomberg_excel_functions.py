#import regula exprssions library to help with finding patterns for the function arguments
import re

def add_BDS_OPT_CHAIN (ticker_cell, type_cell, date_override_cell):

	'''
	Creates a string representing the Bloomberg BDS function for options chains to be inserted into an excel worksheet.
	Note: the document needs to be opened in order for the formulas to be calculated
	'''
	#A cell reference follows the pattern 1 or more letters followed by 1 or more numbers.
	#the variable cell_ref is a regualr expression representing the above pattern

	cell_ref =re.compile(r'[A-Z]+\d+$')
	OPT_CHAIN = '"OPT_CHAIN"'
	OPTION_CHAIN_OVERRIDE = '"OPTION_CHAIN_OVERRIDE","M"'
	 
	#checks if the function arguments are cell references or other strings. A string needs to be wrapped in " "
	if re.match(cell_ref, ticker_cell):
	    TICKER ='{}'.format(ticker_cell)
	else:
	    TICKER = '"{}"'.format(ticker_cell)

	if re.match(cell_ref, type_cell):
	    TYPE = '{}'.format(type_cell)
	else:
	    TYPE = '"{}"'.format(type_cell)
	    
	if re.match(cell_ref, date_override_cell):
	    SINGLE_DATE_OVERRIDE = '{}'.format(date_override_cell)        
	else:
	    SINGLE_DATE_OVERRIDE = '"{}"'.format(date_override_cell)

	#formating the output of the function
	date_concat= 'CONCATENATE("SINGLE_DATE_OVERRIDE=",{})'.format(SINGLE_DATE_OVERRIDE)
	security_concat ='CONCATENATE({}, " ", {})'.format(TICKER, TYPE)
	#formated BDS function
	BDS = '=BDS({},{},{},{})'.format(security_concat,OPT_CHAIN,OPTION_CHAIN_OVERRIDE,date_concat)

	#return the BDS string
	return BDS


def add_BDP_fuction(security_cell, field_cell):
	'''
	Creates a string representing the Bloomberg BDP function to be used in an excel worksheet.
	The two arguments can either be cell references or strings with valid excel BDP function arguments
	'''
	#regular expression pattern for a cell reference.
	#The pattern is any number of of letters followed by any number of digit, and then ends
	cell_ref =re.compile(r'[A-Z]+\d+$')

	#checks to see if both arguments are cell references or not
	if re.match(cell_ref, security_cell) and re.match(cell_ref, field_cell): 
	    BDP = '=BDP({},{})'.format(security_cell, field_cell)
	#if onlyt the security_cell is a cell reference
	elif re.match(cell_ref, security_cell):
	    BDP = '=BDP({},"{}")'.format(security_cell, field_cell)
	#if only the field_cell is a cell reference    
	elif re.match(cell_ref, field_cell):
	    BDP = '=BDP("{}",{})'.format(security_cell, field_cell)
	#if neither the security_cell nor the field_cell is a cell reference
	else:
	    BDP = '=BDP("{}","{}")'.format(security_cell, field_cell)
	#return the BDP string
	return BDP


def add_option_BDH(security_name, fields, start_date, end_date, optional_arg = None, optional_val=None):
    '''
    Creates a string representing the Bloomberg BDH function to be used in an excel worksheet.
    
    Secuirty_name:  can either be a cell reference or a BDH accepted string
    
    fields:         can either be a single cell reference or a range of cells
                    can also be a single BDH string or list of BDH strings

    start_date:     can either be a single cell reference or a BDH accepted string in the form 'yyyymmdd'

    end_date:       can either be a single cell reference or a BDH accepted string in the form 'yyyymmdd'

    optional_arg:   can either be a single string or a list of optional arguments accepted by the BDH function

    optional_val:   if optional arguments are give, a value must be provided for each argument passed.
                    Arguments can either be a single sting,list of strings, cell reference or cell range.

    '''
    #regular expression pattern for a cell reference. The pattern is any number of of letters(capital or lowercase) followed by any number of digits.
    cell_ref = re.compile(r'^[A-Za-z]+\d+$')
    #regular expresson pattern for a cell range. the pattern included two cell reference patterns seperated by a :
    cell_range = re.compile(r'^[A-Za-z]+\d+:[A-Za-z]+\d+$')

    #regular_expression for a sheet cell reference
    sheet_cell_reference = re.compile(r"^'.+'![A-Za-z]+\d+$")

    #regular expression for a date
    formated_date = re.compile(r'\d{8}')

    #checks if security_name is  either a cell reference or a string
    if re.match(cell_ref, security_name):
        SECURITY_NAME = security_name
    else:
        SECURITY_NAME ='"{}"'.format(security_name)
    
    #checks if fields is a list
    if type(fields) == list:
        FIELDS = '"{}"'.format(', '.join(fields))
    #checks if fields is a single cell reference OR a range of cells
    elif re.match(cell_ref, fields) or re.match(cell_range, fields):
        FIELDS = fields
    #check if fields is a string
    elif type(fields) == str:
        FIELDS = '"{}"'.format(fields)

    #checks if the start_date is a cell reference or a cell reference from another sheet
    if re.match(cell_ref, start_date) or re.match(sheet_cell_reference, start_date):
        START_DATE = start_date
    #checks if the other input isn't true, check if the date was formated correctly
    elif re.match(formated_date, start_date):
        START_DATE = '"{}"'.format(start_date)
    else:
        return 'Error: start_date not formated correctly'

    #checks if the end_date is a cell reference or a cell reference from another sheet
    if re.match(cell_ref, end_date) or re.match(sheet_cell_reference, start_date):
        END_DATE = end_date
    #checks if the other input isn't true, check if the date was formated correctly    
    elif re.match(formated_date, end_date):
        END_DATE = '"{}"'.format(end_date)
    else:
        return 'Error: end_date not formated correctly'
    
    #if optional_arg was provided:
    if optional_arg != None:
        #optional_arg is a list
        if type(optional_arg) == list:
            OPTIONAL_ARG = '"{}"'.format(', '.join(optional_arg))
        #checks if optional_arg is a cell reference or range
        elif re.match(cell_ref, optional_arg) or re.match(cell_range, optional_arg):
            OPTIONAL_ARG = optional_arg
        #optional_arg is a string    
        elif type(optional_arg) == str:
            OPTIONAL_ARG = '"{}"'.format(optional_arg)
        
    #if optional_val was provided:
    if optional_val != None:
        #if the optional_val was a list
        if type(optional_val) == list: 
            OPTIONAL_VAL = '"{}"'.format(', '.join(optional_val))
        #if the optoinal_val was a cell reference or range
        elif re.match(cell_ref, optional_val) or re.match(cell_range, optional_val):
            OPTIONAL_VAL = optional_val
        #if the option_val was a string
        elif type(optional_val) == str:
            OPTIONAL_VAL = '"{}"'.format(optional_val)
        
    #checks if optional_arg was given but optioinal_val wasn't
    if (optional_arg != None) and (optional_val == None):
        return 'Please provid values for {}'.format(OPTIONAL_ARG)
    
    #if optional_arg is None don't include either OPTIONAL_ARG OR OPTIONAL_VAL
    if optional_arg == None:
        BDH = '=BDH({},{},{},{})'.format(SECURITY_NAME, FIELDS, START_DATE, END_DATE)
    #checks if optional_arg was given but optioinal_val wasn't
    elif (optional_arg != None) and (optional_val == None):
        return 'Please provid values for {}'.format(OPTIONAL_ARG)
    #both optoinal parameters were provided
    else:
        BDH = '=BDH({},{},{},{},{},{})'.format(SECURITY_NAME, FIELDS, START_DATE, END_DATE, OPTIONAL_ARG, OPTIONAL_VAL)
    #return the BDH string
    return BDH



