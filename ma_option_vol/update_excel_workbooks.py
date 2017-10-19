import openpyxl
import datetime as dt
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
    for (index, cell) in enumerate(sheet['A10:B{}'.format(sheet.max_row)]):
        #cell[0] corresponds to cells in column A and cell[1] corresponds to cells in column B
        cell[1].value = abxl.add_BDP_fuction(cell[0].coordinate, "SECURITY_DES")
        print(cell[0].value, cell[1].value)
    #saves the workbook
    #wb.save(workbook_path) <----uncomment this after testing

def update_option_contract_tabs(workbook_path, sheet_name, sheet_end_date_cell, test=False):
    '''
    Creates new sheets in the given excel workbook based on Option data stored in the given sheet.
    '''
    #data labels to be added to the new excel worksheet
    option_data_labels = ['Security Name', 'Description', 'Type', 'Expiration Date', 'Strike Price']

    #headers for the datatable to be added to the new worksheet
    data_table_header = ['INDEX','DATE','PX_LAST','PX_BID','PX_ASK','PX_VOLUME','OPEN_INT', 'IVOL']


    #given the full_file path an excel workbook is loaded.
    wb = openpyxl.load_workbook(workbook_path)
    
    #The sheet we want to get data from is set to the variable sheet
    data_sheet = wb.get_sheet_by_name(sheet_name)
    
    #The cell in the sheet that contains the completion/termination date
    completion_date= data_sheet[sheet_end_date_cell].value.date()

    #if we're running a test, do the following:
    if test:
        #if there are already tabs get rid of the ones we don't want:
        print(len(wb.get_sheet_names()))
        if len(wb.get_sheet_names()) >1:
               for i,x in enumerate(wb.get_sheet_names()):
                    if i > 0:
                        wb.remove_sheet(wb.get_sheet_by_name(x))


    total_rows = data_sheet.max_row
    #iterate through the rows of the data_sheet
    #NOTE: THE SHEET IS SET UP SO THAT VALUES WE'RE INTERESTED IN START AT ROW 10
    for (index, cell) in enumerate(data_sheet['A10:B{}'.format(total_rows)]):
        
        #if there is no option description, then break out of this loop
        if cell[1].value == None:
            print('No option description found. Could not create new workbook sheets')
            break

        #format_option_data() returns the following list:
        #[security_name, option_description, option_type, expiration_date, strike_price]
        option_data = format_option_data(cell[0].value, cell[1].value)

        #the number of days between the expiration and completion date. 
        date_diff = (option_data[3] - completion_date).days

        #if the expiration_date occurs 2 months after the completion_date, then stop creating sheets
        if date_diff >= 60:
            print('Found contract past {}. Saving the workbook with {} new tabs'.format(option_data[3], index-1))
            print('Potential new tabs: {}'.format(total_rows-10))
            wb.save(workbook_path)
            break

        #otherwise, keep creating sheets
        else:
            #creates a new sheet for the passed in workbook
            new_sheet = wb.create_sheet()
            #/' aren't allowed in excel sheet names, so we replace them with '-' 
            new_sheet.title = option_data[1].replace('/', '-')

            #zip creates a tuple for each item of the passed in lists. this tuple can then be appended to the sheet
            for data in zip(option_data_labels,option_data):
                new_sheet.append(data)

            #loop through every value of the data_table_header and add it to the worksheet A8:H8
            for index, value in enumerate(data_table_header, start= 1) :
                new_sheet.cell(row = 8,column = index ).value = value 

            #add the BDH formula to cell B9
            new_sheet['B9'] = abxl.add_option_BDH(  security_name = 'B1',
                                                    fields = 'C8:H8', 
                                                    start_date = "'Options Chain'!B4",
                                                    end_date = "'Options Chain'!B6",
                                                    optional_arg = ['Days', 'Fill'],
                                                    optional_val = ['W',  '0'])
            
    #if the loop ends without finding contracts 2 months past the completion/termination date, save the workbook      
    wb.save(workbook_path)  
                                                    
    #     new_sheet['A1'] ='Security Name'
    #     new_sheet['A2'] ='Description'
    #     new_sheet['A3'] = 'Type'
    #     new_sheet['A4'] = 'Expiration Date'
    #     new_sheet['A5'] = 'Strike'

    #     #Setting the values for the labels
    #     new_sheet['B1'] = security_name
    #     new_sheet['B2'] = option_description
    #     if description_list[-1][0] =='P':
    #         new_sheet['B3'] = 'Put'
    #     elif description_list[-1][0] == 'C':
    #         new_sheet['B3'] = 'Call'
        
    #     #we converted option_descrption to a string earlier, so we have to compare option_description to the string 'None'
    #     if option_description == 'No Options':
    #         break
    #     else:
    #         new_sheet['B4'] = description_list[2]
    #         new_sheet['B5'] = description_list[-1][1:]

    #     data_table_header = ['INDEX','DATE','PX_LAST','PX_BID','PX_ASK','PX_VOLUME','OPEN_INT', 'IVOL']
    #     #ADDING DATA COLUMN LABELS:
    #     new_sheet['A8'] = 'INDEX'
    #     new_sheet['B8'] = 'DATE'
    #     new_sheet['C8'] = 'PX_LAST'
    #     new_sheet['D8'] = 'PX_BID'
    #     new_sheet['E8'] = 'PX_ASK'
    #     new_sheet['F8'] = 'PX_VOLUME'
    #     new_sheet['G8'] = 'OPEN_INT'
    #     new_sheet['H8'] = 'IVOL'

    #     #add the BDH formula to cell B9
    #     new_sheet['B9'] = abxl.add_option_BDH(  security_name = 'B1',
    #                                             fields = 'C8:H8', 
    #                                             start_date = "'Options Chain'!B4",
    #                                             end_date = "'Options Chain'!B6",
    #                                             optional_arg = '"Days, Fill"',
    #                                             optional_val = '"W,  0"')


    #     #Get rid of break after done testing:
    #     #wb.save(file_path) <-------------------remove after done testing
    # print('Done')

def format_option_data(security_name, option_description):
    '''
    security_name should be a string that looks similar to 'BBG00673J6L5 Equity'
    option_description should be a string that looks similar to 'PFE US 12/20/14 P18'
    given the security_name and option_description a list of data will be output
    '''
    #will split the option_description by whitespace into a list that looks like: ['PFE', 'US', '12/20/14', 'P18']
    description_list = option_description.split(' ')

    #determins the option type based on description_list[-1][0] = 'P'
    if description_list[-1][0] =='P':
        option_type = 'Put'
    elif description_list[-1][0] == 'C':
        option_type = 'Call'

    #description_list[2] = 12/20/14 and convertis it into a datetime object
    expiration_date = dt.datetime.strptime(description_list[2],'%m/%d/%y').date()

    #description_list[-1][1:] = '18', and converts the string to an int
    strike_price = description_list[-1][1:]

    option_data_list = [security_name, option_description, option_type, expiration_date, strike_price]

    return option_data_list


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