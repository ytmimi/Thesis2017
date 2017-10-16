#######Done

def create_company_workbooks(sample_file_path,target_path, acquirer_path):
    #saves the n
    wb = openpyxl.load_workbook(sample_file_path)
    sheet = wb.get_sheet_by_name('Filtered Sample Set')
    
    #iterates over all the rows of the worksheet
    for (i, row) in enumerate(sheet.rows):
        #skips the first row of the worksheet because it contains column titles
        if i < 1:
            continue
        else:
            #creates the new workbooks
            new_target_workbook(row, target_path)
            new_acquirer_workbook(sheet, acquirer_path)
            break
    print('done')


    def new_target_workbook(row_data, target_path):
    '''
    row_data is a tuple with the following indexes
    0)Deal Type 1)Announce Date 2)Completion/Termination Date
    3)Target Name 4)Target Ticker 5)EQY_OPT_AVAIL   
    6)Acquirer Name 7)Acquirer Ticker 8)EQY_OPT_AVAIL 9)Announced Total Value (mil.)
    10)Payment Type 11)TV/EBITDA 12)Deal Status 13)Stock Terms
    '''
    one_year = dt.timedelta(days=360)
    start_date = row_data[1].value - one_year
    
    #a list of data that will be added to each newly created worksheet
    data = [['Target Name', row_data[3].value], 
            ['Target Ticker', row_data[4].value],
            ['Type', 'Equity'],
            ['Start Date', start_date.date()],
            ['Announcement Date', row_data[1].value.date()],
            ['End Date', row_data[2].value.date()],
            ['Formated Start Date',str(start_date.date()).replace('-','')],
            ['Formated End Date',str(row_data[2].value.date()).replace('-','')]]
    
    #creates a new Workbook
    wb_target = openpyxl.Workbook()
    target_sheet = wb_target.get_active_sheet()
    target_sheet.title = 'Options Chain'
            

    #appends the data to the workbook        
    for (index, cell) in enumerate(target_sheet['A1:B8']):
        (cell[0].value, cell[1].value) = data[index]
    
    target_sheet['A10'] = add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B7')
          
    #save the worksheet
    #wb_target.save(target_path)
    
    
def new_acquirer_workbook(row_data, acquirer_path):
    '''
    row_data is a tuple with the following indexes
    0)Deal Type 1)Announce Date 2)Completion/Termination Date
    3)Target Name 4)Target Ticker 5)EQY_OPT_AVAIL   
    6)Acquirer Name 7)Acquirer Ticker 8)EQY_OPT_AVAIL 9)Announced Total Value (mil.)
    10)Payment Type 11)TV/EBITDA 12)Deal Status 13)Stock Terms
    '''
    one_year = dt.timedelta(days=360)
    start_date = row_data[1].value - one_year
    
    #a list of data that will be added to each newly created worksheet
    data = [['Acquirer Name', row_data[6].value], 
            ['Acquirer Ticker', row_data[7].value],
            ['Type', 'Equity'],
            ['Start Date', start_date.date()],
            ['Announcement Date', row_data[1].value.date()],
            ['End Date', row_data[2].value.date()],
            ['Formated Start Date',str(start_date.date()).replace('-','')],
            ['Formated End Date',str(row_data[2].value.date()).replace('-','')]]
    
    #creates a new Workbook
    wb_acquirer = openpyxl.Workbook()
    acquirer_sheet = wb_target.get_active_sheet()
    acquirer_sheet.title = 'Options Chain'
            
    #appends the data to the workbook        
    for (index, cell) in enumerate(acquirer_sheet['A1:B8']):
        (cell[0].value, cell[1].value) = data[index]
    
    acquirer_sheet['A10'] = add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B7')
          
    #save the worksheet
    #wb_acquirer.save(acquirer_path)


############################################### UPDATE #############################################################
def add_BDS_OPT_CHAIN(ticker_cell, type_cell, date_override_cell):
    '''
    creates the Bloomberg BDS function for options chains to be inserted into an excel WorkSheets Cell
    Note: the document needs to be run in order for the formulas to be calculated
    
    Example: add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B6')
             Gets values from specific cells
             
    Example: add_BDS_OPT_CHAIN(ticker_cell='AAPL US',type_cell='EQUITY', date_override_cell='20151231')
             Gets data from bloomberg based on the literal commands excepted by the BDS function
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
    #creates the string version of the Bloomberg BDP function to be used in an excel worksheet:
    #if a cell reference is passed the field_cell variable will most likely be a bloomberg field
    if len(field_cell) > 2: 
        BDP = '=BDP({},CONCATENATE("{}"))'.format(security_cell, field_cell)
    #if the field_cell <= 2, then chances are a cell reference was passed
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
    
#Once I implemented read_data_only, I didn't need to copy values since after that fuction is called, the workbook no
#longer contains function references, but just data. this function may find some other use though, but I don't think so
def copy_data_DE(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.get_sheet_by_name('Options Chain')
    
    max_rows = sheet.max_row
    
    for i in range(10, max_rows+1):
        sheet['D{}'.format(i)] = sheet['A{}'.format(i)].value
        sheet['E{}'.format(i)] = sheet['B{}'.format(i)].value
    
    wb.save(file_path)
        
    
def read_data_only(file_path):
    wb=openpyxl.load_workbook(file_path, data_only = True)
    wb.save(file_path)
    
    

def create_option_contract_tabs(file_path):
    '''
    NOTE: the data from Bloomberg should be copied into column D and E respecivly
    '''
    wb = openpyxl.load_workbook(file_path)
    #The active sheet is the first sheet, which in this case is "Options Chain"
    sheet = wb.get_sheet_by_name('Options Chain')
    
    
    #if there are already tabs get rid of the ones we don't want:
    if len(wb.get_sheet_names()) >1:
           for i,x in enumerate(wb.get_sheet_names()):
                if i > 0:
                    wb.remove_sheet(wb.get_sheet_by_name(x))

    # now add the sheets that we want
    #NOTE: THE SHEET IS SET UP SO THAT VALUES WE'RE INTERESTED IN START AT ROW 10
    for i in range(10, sheet.max_row+1):
        
        new_sheet = wb.create_sheet()
        
        option_description = sheet['B{}'.format(i)].value
        
        #just in case we weren't able to find options, the tab can still be created
        if option_description == None:
            option_description = 'No Options'
        security_name = sheet['A{}'.format(i)].value
      
        # '/' aren't allowed in excel tab names, so we replace them with '-' 

        tab_name = option_description.replace('/', '-')

        #will split the option_description into a list that looks like: ['PFE', 'US', '12/20/14', 'P18']
        #This information will be used to give values to the labels set up below
        description_list = option_description.split(' ')


        #Setting the labels for the sheet
        new_sheet.title =  tab_name
        new_sheet['A1'] ='Security Name'
        new_sheet['A2'] ='Description'
        new_sheet['A3'] = 'Type'
        new_sheet['A4'] = 'Expiration Date'
        new_sheet['A5'] = 'Strike'

        #Setting the values for the labels
        new_sheet['B1'] = security_name
        new_sheet['B2'] = option_description
        if description_list[-1][0] =='P':
            new_sheet['B3'] = 'Put'
        elif description_list[-1][0] == 'C':
            new_sheet['B3'] = 'Call'
        else:
            new_sheet['B3'] = '?'
        
        #we converted option_descrption to a string earlier, so we have to compare option_description to the string 'None'
        if option_description == 'No Options':
            break
        else:
            new_sheet['B4'] = description_list[2]
            new_sheet['B5'] = description_list[-1][1:]

        #ADDING DATA COLUMN LABELS:
        new_sheet['A8'] = 'INDEX'
        new_sheet['B8'] = 'DATE'
        new_sheet['C8'] = 'PX_LAST'
        new_sheet['D8'] = 'PX_BID'
        new_sheet['E8'] = 'PX_ASK'
        new_sheet['F8'] = 'PX_VOLUME'
        new_sheet['G8'] = 'OPEN_INT'
        new_sheet['H8'] = 'IVOL'

        #add the BDH formula to cell B9
        new_sheet['B9'] = add_option_BDH()
        #Get rid of break after done testing:
        wb.save(file_path)
    print('Done')
        #if i > 15:
        #break #<--- remove after testing.  Just want to make sure that we can create one tab the way that we want

def add_option_BDH():
    security_name = 'B1'
    fields = 'C8:H8'
    start_date = "'Options Chain'!B4"
    end_date = "'Options Chain'!B6"
    optional_arg = '"Days, Fill"'
    optional_val = '"W,  0"'
    
    BDH = '=BDH({},{},{},{},{},{})'.format(security_name, fields, start_date, end_date, optional_arg,optional_val)
    return BDH



def find_index_0(file_path):
    '''
    After the BDH function has been added to the file, then this function should be called
    '''
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.get_sheet_by_name('Options Chain')
    
    announcement_date = sheet['B5'].value
    
    #a list of all the tabs in the workbook
    tab_list = wb.get_sheet_names()
    
    reference_sheet = wb.get_sheet_by_name(tab_list[1]) #<--the first sheet after the 'Options Chain' Sheet
    #assignes the max number of rows in the spreadsheet to the variable 
    ref_sheet_max_rows = reference_sheet.max_row
    
    #Goes through each row in the sheet to check if we've found the announcement date
    for i in range(9,ref_sheet_max_rows+1):
        if reference_sheet['B{}'.format(i)].value == announcement_date:
            _0_index = i
        
    
    #adds value's to the INDEX column in the reference sheet
    for j in range(9, ref_sheet_max_rows+1):
        reference_sheet['A{}'.format(j)] = j - _0_index # will give a value of 0 once we reach the _0_index
    
    for k,x in enumerate(tab_list):
        #skip the first two tabs because they are already assigned to sheet and reference_sheet above
        if k>1:
            #create a new sheet given the sheeet we are iterating over
            new_sheet = wb.get_sheet_by_name(x)
            for row in range (9, ref_sheet_max_rows+1): #<----all tabs have the same rows as the reference_sheet
                #assignes the values in the A column for each new sheet created to match those from the reference_sheet
                new_sheet['A{}'.format(row)] = reference_sheet['A{}'.format(row)].value
    
    print('done')
    wb.save(file_path)