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
    wb.save(workbook_path)


def update_option_contract_sheets(workbook_path, sheet_name, sheet_end_date_cell):
    '''
    Creates new sheets in the given excel workbook based on Option data stored in the given sheet.
    '''
    #data labels to be added to the new excel worksheet
    option_data_labels = ['Security Name', 'Description', 'Type', 'Expiration Date', 'Strike Price']

    #headers for the datatable to be added to the new worksheet
    data_table_header = ['INDEX','DATE','PX_LAST','PX_BID','PX_ASK','PX_VOLUME','OPEN_INT', 'IVOL']

    #given the file path, an excel workbook is loaded.
    wb = openpyxl.load_workbook(workbook_path)
    
    #The sheet we want to get data from is set to the variable sheet
    data_sheet = wb.get_sheet_by_name(sheet_name)
    
    #The cell in the sheet that contains the completion/termination date, as passed in by the function.
    completion_date= data_sheet[sheet_end_date_cell].value.date()

    total_rows = data_sheet.max_row
    #iterate through the rows of the data_sheet
    #NOTE: THE SHEET IS SET UP SO THAT VALUES WE'RE INTERESTED IN START AT ROW 10
    for (index, cell) in enumerate(data_sheet['A10:B{}'.format(total_rows)]):
        
        #if there is no option description, then break out of this loop
        if cell[1].value == None:
            print('No option description found. Could not create new workbook sheets')
            wb.save(workbook_path)
            break

        #format_option_description() returns the following list:
        #[security_name, option_description, option_type, expiration_date, strike_price]
        option_data = format_option_description(cell[0].value, cell[1].value)

        #the number of days between the expiration and completion date. 
        date_diff = (option_data[3] - completion_date).days

        #if the expiration_date occurs 2 months after the completion_date, then stop creating sheets
        if date_diff >= 60:
            print('Found contracts past {}. Saving the workbook with {} new tabs'.format(completion_date, index))
            wb.save(workbook_path)
            break

        #otherwise, keep creating sheets
        else:
            #creates a new sheet for the passed in workbook
            new_sheet = wb.create_sheet()
            #/' aren't allowed in excel sheet names, so we replace them with '-' if the name contains '/' 
            new_sheet.title = option_data[1].replace('/', '-')

            #zip creates a tuple pair for each item of the passed in lists. this tuple can then be appended to the sheet
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
 

def format_option_description(security_name, option_description):
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


def update_workbook_data_index(workbook_path):
    '''
    Given a workbook, loop through all the sheets of that workbook and update the index for each sheet.
    '''
    #loads an excel workbook given the file path to that workbook.
    wb = openpyxl.load_workbook(workbook_path)
    #gets a list of all the sheets in the workbook
    sheet_list = wb.get_sheet_names()

    #iterates through every sheet
    for (index, sheet_name) in enumerate(sheet_list):
        #gets the sheet given the name in the sheet list
        sheet = wb.get_sheet_by_name(sheet_name)       
        #skips the first sheet of the workbook, becasuse data isn't stored there.
        #indexing starts at 0.
        if index == 0:
            #get the announcement date from the first sheet
            announcement_date = sheet['B5'].value
        if index > 0:
            update_sheet_index(sheet_name= sheet, date=announcement_date, start_row= 9)
    wb.save(workbook_path)
    print('Saving workbook')


def update_sheet_index(sheet_name, date, start_row):
    '''
    Given an excel worksheet,a designated date, and a starting row,
    an index is added for each date relative to the specified date and row
    '''
    #gets the total number of rows in the worksheet
    total_rows = sheet_name.max_row
    #iterates over every cell in column

    index_0 =find_index_0(worksheet=sheet_name,start= start_row, end=total_rows, date_0= date)
    #iterates over every column in the given date_column from the start to the end of the sheet
    for index in range(start_row, total_rows+1):
        sheet_name.cell(row= index, column=1).value = index - index_0



def update_read_data_only(file_path):
    wb=openpyxl.load_workbook(file_path, data_only = True)
    wb.save(file_path)


def delet_workbook_sheets(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)
    start_sheet_num = len(wb.get_sheet_names())
    for (index,sheet) in enumerate(wb.get_sheet_names()):
        if index > 0:
            wb.remove_sheet(wb.get_sheet_by_name(sheet))
    end_sheet_num = len(wb.get_sheet_names())
    deleted_sheet_num = start_sheet_num - end_sheet_num 
    print('Deleted {} sheets from the Workbook'.format(deleted_sheet_num))
    wb.save(workbook_path)


def find_index_0(worksheet,start, end, date_0):
    '''
    binary search function to determine which row index of the worksheet
    contains the date we're looking for.
    '''
    #list comprehesion  for all the row indexes.
    index_list = [x for x in range(start,end+1)]
    start_index = index_list[0]
    end_index = index_list[-1]
    average_index = int((end_index + start_index)/2)
    #variable for the while loop
    
    #import pdb; pdb.set_trace()
    found = False
    while not found:
        #print(start_index, found)        
        #import pdb; pdb.set_trace()
        curr_date = worksheet['B{}'.format(average_index)].value
        if (date_0 == curr_date):
            found = True

        elif (date_0 > curr_date):
            start_index = average_index +1
            average_index = int((end_index + start_index)/2)

        elif (date_0 < curr_date):
            end_index = average_index -1
            average_index = int((end_index + start_index)/2)
  
    return average_index
    
    












