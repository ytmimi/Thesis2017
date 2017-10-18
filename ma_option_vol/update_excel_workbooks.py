import openpyxl
import add_bloomberg_excel_functions as abxl

def update_sheet_with_BDP_description(workbook_path, sheet_name):
    '''
    Given an excel workbook, The BDP function is added with the appropriate cell reference and description  
    Note: The excel workbook needs to be opened and saved so that bloomberg data can populate
    '''
    #opens the workbook
    wb = openpyxl.load_workbook(workbook_path)
    #gets the specified sheet from the workbook
    sheet = wb.get_sheet_by_name(sheet_name)
    
    #iterate over every row in column A and B starting at A10:B10 and ending at the last row of the worksheet
    for index, cell in enumerate(sheet['A10:B{}'.format(sheet.max_row)]):
        #cell[0] corresponds to cells in column A and cell[1] corresponds to cells in column B
        cell[1].value = abxl.add_BDP_fuction(cell[0].coordinate, "SECURITY_DES")
        print(cell[0].value, cell[1].value)
    #saves the workbook
    wb.save(workbook_path)

def update_option_contract_tabs(file_path):
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
        new_sheet['B9'] = abxl.add_option_BDH(  security_name = 'B1',
                                                fields = 'C8:H8', 
                                                start_date = "'Options Chain'!B4",
                                                end_date = "'Options Chain'!B6",
                                                optional_arg = '"Days, Fill"',
                                                optional_val = '"W,  0"')


        #Get rid of break after done testing:
        wb.save(file_path)
    print('Done')
        #if i > 15:
        #break #<--- remove after testing.  Just want to make sure that we can create one tab the way that we want



def update_index(file_path):
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


def update_read_data_only(file_path):
    wb=openpyxl.load_workbook(file_path, data_only = True)
    wb.save(file_path)