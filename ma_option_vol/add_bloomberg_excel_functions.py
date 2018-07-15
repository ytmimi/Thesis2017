import datetime as dt
import re

#regular expression that define cell and range patterns in excel
CELL_REF = re.compile(r'^[A-Za-z]+\d+$')
CELL_RANGE= re.compile(r'^[A-Za-z]+\d+:[A-Za-z]+\d+$')
SHEET_CELL_REF = re.compile(r"^'.+'![A-Za-z]+\d+$")
SHEET_CELL_RANGE = re.compile(r"^'.+'![A-Za-z]+\d+:[A-Za-z]+\d+$")

#regular exression for a formatted date
FORMAT_DATE = re.compile(r'^\d{8}')

def is_cell_ref(item):
	'''Checks if an input is a cell reference or cell range'''
	return (re.match(CELL_REF, item) or re.match(CELL_RANGE, item)
			or re.match(SHEET_CELL_REF, item) or re.match(SHEET_CELL_RANGE, item))

class Bloomberg_Excel():
	'''Bloomberg ad-in Excel functions. Note: Excel Workbooks must be opened before data can populate'''
	def BDP(self, security, field, type_=None):
		'''BDP (Bloomberg Data Point) returns static or real time current market data.
		security:  	Can either be a cell reference, or a string.
   		field:		Can either be a cell reference, cell range, list, or string.
   		type_:		If the security does not end with a propper identifier like: 
   					"Govt", "Corp", " Mtge", "M-Mkt", "Muni", "Pfd", "Equity", "Comdty", 
   					"Index", or "Crncy" please provide the type. Otherwise leave type=None.
   		'''
		security = self.security_type(security, type_)
		field = self.check_input(field)
		return'=BDP({},{})'.format(security, field)

	def BDS(self, security, field, type_=None, opt_args=None, opt_vals=None):
		'''BDS (Bloomberg Data Set) returns informational bulk data
		security:		Can either be a cell reference, or a string
		field:			Can either be a cell reference, cell range, list, or string
		type_:			If the security does not end with a propper identifier like: 
						"Govt", "Corp", " Mtge", "M-Mkt", "Muni", "Pfd", "Equity", "Comdty", 
						"Index", or "Crncy" please provide the type. Otherwise leave type=None.
		opt_args:		can either be a single string or a list of optional arguments
		opt_vals:		if optional arguments are give, a value must be provided for each argument passed.
						arguments can either be a single sting,list of strings, cell reference or cell range.
		'''
		security = self.security_type(security, type_)
		field = self.check_input(field)
		optional = self.check_optional(opt_args, optional_val)
		if optional != None:
			return 'BDS({},{},{},{})'.format(security, field, optional[0], optional[1])
		else:
			return '=BDS({},{})'.format(security, field)

	def BDS_OPT_CHAIN(self, security, opt_override, date_override, type_=None):
		'''Specific for BDS Option contract data			
		security:	  	Can either be a cell reference, or a string.
		opt_override:	Either "A"-all, "Q"-qurterly, "M"-monthly, and "W"-weekly option contracts
		date_override:	can either be a cell reference, an int, or a string in the form 'yyyymmdd'
		'''
		security = self.security_type(security, type_)
		date_override = self.check_input(date_override)
		return  BDS(security, 'OPT_CHAIN', opt_args=['OPTION_CHAIN_OVERRIDE', 'SINGLE_DATE_OVERRIDE'], opt_vals=[opt_override, date_override])

	def BDH(self, security, field, start_date, end_date, type_=None, opt_args=None, opt_vals=None):
		'''BDH (Bloomberg Data History)
		returns the historical data for a selected security or set of securities.
		security:  		can either be a cell reference, or a string
   		field:			can either be a cell reference, cell range, list, or string
    	start_date:		can either be a cell reference, an int, or a string in the form 'yyyymmdd'
		end_date:       can either be a cell reference, an int, or a string in the form 'yyyymmdd'
		type_:			If the security does not end with a propper identifier like: 
   						"Govt", "Corp", " Mtge", "M-Mkt", "Muni", "Pfd", "Equity", "Comdty", 
   						"Index", or "Crncy" please provide the type. Otherwise leave type=None.
    	opt_args:   	can either be a single string or a list of optional arguments accepted by the BDH function
    	opt_vals:   	if optional arguments are give, a value must be provided for each argument passed.
                    	arguments can either be a single sting,list of strings, cell reference or cell range.
		'''
		security = self.security_type(security, type_)
		field = self.check_input(field)
		start_date = self.check_input(start_date)
		end_date = self.check_input(end_date)
		optional = self.check_optional(opt_args, opt_vals)
		if optional != None:
			return '=BDH({},{},{},{},{},{})'.format(security, field, start_date, end_date, optional[0], optional[1])
		else:
			return '=BDH({},{},{},{})'.format(security, field, start_date, end_date,)

	def check_input(self, item):
		''' Checks the type of the input and returns a value Excel can work with'''
		if type(item) == list:
			return '"{}"'.format(', '.join(item))
		elif type(item) == int:
			return '"{}"'.format(item)
		elif type(item) == str:
			if is_cell_ref(item):
				return item
			else:
				return '"{}"'.format(item)
		elif type(item) == dt.datetime or type(item) == dt.date:
			return '"{}"'.format(item.strftime('%Y%m%d'))
		else:
			return '"{}"'.format(item)

	def security_type(self, security, type_=None):
		'''Checks the security input, and then appends the type if it is provided
		security:  		can either be a cell reference, or a string
		type_:			If the security does not end with a propper identifier like: 
   						"Govt", "Corp", " Mtge", "M-Mkt", "Muni", "Pfd", "Equity", "Comdty", 
   						"Index", or "Crncy" please provide the type. Otherwise leave type_=None.
		'''
		# security = 
		if type_ != None and is_cell_ref(security):
			security = 'CONCATENATE({}, " ", {})'.format(security, type_)
		elif type_ != None and not is_cell_ref(security):
			security = '{} {}'.format(security, type_)
		return self.check_input(security)

	def check_optional(self, arguments, values):
		if arguments == values == None:
			return None
		elif arguments == None and values !=None:
			raise ValueError('Arguments were not defined but values were')
		elif arguments != None and values == None:
			raise ValueError('Arguments were defined by values were not')
		else:
			return (self.check_input(arguments), self.check_input(values))

####################################################################### END CLASS

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
	DEPRECATED use Bloomberg_Excel().BDP()
	'''
	return Bloomberg_Excel().BDP(security_cell, field_cell)


def add_option_BDH(security_name, fields, start_date, end_date, optional_arg = None, optional_val=None):
	''' DEPRECATED use Bloomberg_Excel().BDH() '''
	return Bloomberg_Excel().BDH(security_name, fields, start_date, end_date, opt_args=optional_arg, opt_vals=optional_val)


