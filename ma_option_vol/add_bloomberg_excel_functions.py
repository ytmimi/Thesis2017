import openpyxl

def add_BDS_OPT_CHAIN(ticker_cell, type_cell, date_override_cell):
    '''
    creates a string representing the Bloomberg BDS function for options chains to be inserted into an excel WorkSheets Cell
    Note: the document needs to be run in order for the formulas to be calculated
    '''
    OPT_CHAIN = '"OPT_CHAIN"'
    OPTION_CHAIN_OVERRIDE = '"OPTION_CHAIN_OVERRIDE","M"'
     
    if len(ticker_cell) > 2:
        TICKER = '"{}"'.format(ticker_cell)
    else:
        TICKER ='{}'.format(ticker_cell)
    
    if len(type_cell) > 2:
        TYPE = '"{}"'.format(type_cell)
    else:
        TYPE = '{}'.format(type_cell)
    
    if len(date_override_cell) >2:
        SINGLE_DATE_OVERRIDE = '"{}"'.format(date_override_cell)
    else:
        SINGLE_DATE_OVERRIDE = '{}'.format(date_override_cell)
    
    date_concat= 'CONCATENATE("SINGLE_DATE_OVERRIDE=",{})'.format(SINGLE_DATE_OVERRIDE)
    security_concat ='CONCATENATE({}, " ", {})'.format(TICKER, TYPE)
    #formated BDS function
    BDS = '=BDS({},{},{},{})'.format(security_concat,OPT_CHAIN,OPTION_CHAIN_OVERRIDE,date_concat)
    return BDS


def add_BDP_fuction(security_cell, field_cell):
    '''
    Creates the string version of the Bloomberg BDP function to be used in an excel worksheet
    the security_cell, and field_cell can either be cell references or strings
    '''
    if len(security_cell) and len(field_cell) > 2:
        BDP = '=BDP("{}",CONCATENATE("{}"))'.format(security_cell, field_cell)
    elif len(security_cell) > 2: 
        BDP = '=BDP("{}",CONCATENATE({}))'.format(security_cell, field_cell)
    elif len(field_cell) > 2:
        BDP = '=BDP({},CONCATENATE("{}""))'.format(security_cell, field_cell) 
    else:
        BDP = '=BDP({},CONCATENATE({}))'.format(security_cell, field_cell)
    return BDP


def update_sheet_with_BDP_description(file_path):
    '''
    Note: The sheet needs to be opened at least onece so that the BDS function added by create_files() can populate
          Then make sure the notebook is saved after that.
    '''
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.get_active_sheet()
    
    for i in range(10, sheet.max_row+1):
        sheet['B{}'.format(i)] = add_BDP_fuction('A{}'.format(i), "SECURITY_DES")
    
    wb.save(file_path)


def add_option_BDH():
    security_name = 'B1'
    fields = 'C8:H8'
    start_date = "'Options Chain'!B4"
    end_date = "'Options Chain'!B6"
    optional_arg = '"Days, Fill"'
    optional_val = '"W,  0"'
    
    BDH = '=BDH({},{},{},{},{},{})'.format(security_name, fields, start_date, end_date, optional_arg,optional_val)
    return BDH