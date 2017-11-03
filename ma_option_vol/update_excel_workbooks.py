import openpyxl
import datetime as dt
import re
from statistics import mean, stdev
from math import ceil, floor
import add_bloomberg_excel_functions as abxl
from generate_sorted_options_workbooks import convert_to_numbers

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
    #saves the workbook
    wb.save(workbook_path)


def update_option_contract_sheets(workbook_path, sheet_name, sheet_start_date_cell, sheet_announce_date_cell, sheet_end_date_cell,  data_header_row, data_table_index, data_table_header, BDH_optional_arg=None, BDH_optional_val=None):
    '''
    Creates new sheets in the given excel workbook based on Option data stored in the given sheet_name.

    workbook_path       the full file_path to the specified workbook

    sheet_name          Specify the sheet_name that data will be referenced from

    sheet_start_date_cell Give the coordinates of the cell in 

    sheet_end_date_cell Specify the coordinates of the the cell in the specified sheet that contains the end date
    '''
    #a regular expression for a formated option description where the strike is an integer
    option_description_pattern_int = re.compile(r'^\w+\s\w+\s\d{2}/\d{2}/\d{2}\s\w+$')

    #a regular expression for a formated option description where the strike is a foat
    option_description_pattern_float = re.compile(r'^\w+\s\w+\s\d{2}/\d{2}/\d{2}\s\w+\.\w+$')

    #combine data_table_index and data_table_header
    total_data_headers = data_table_index+data_table_header


    #data labels to be added to the new excel worksheet
    option_data_labels = ['Security Name', 'Description', 'Type', 'Expiration Date', 'Strike Price']
    
    #given the file path, an excel workbook is loaded.
    wb = openpyxl.load_workbook(workbook_path)
    
    #The sheet we want to get data from is set to the variable data_sheet
    data_sheet = wb.get_sheet_by_name(sheet_name)
    

    #The cell in the sheet that contains the completion/termination date, as passed in by the function.
    if type(data_sheet[sheet_end_date_cell].value) == int:
        completion_date = dt.datetime.strptime(str(data_sheet[sheet_end_date_cell].value),'%Y%m%d').date()
    else:
        completion_date= data_sheet[sheet_end_date_cell].value.date()

    #The cell in the sheet that contains the announcement date, as passed in by the function.
    if type(data_sheet[sheet_announce_date_cell].value) == int:
        completion_date = dt.datetime.strptime(str(data_sheet[sheet_end_date_cell].value),'%Y%m%d').date()
    else:
        completion_date= data_sheet[sheet_end_date_cell].value.date()

    total_rows = data_sheet.max_row

    #counter to keep track of each sheet created
    sheet_count = 0
    #gets the average stock price and standard deviation of the stock price data for the historic and merger period:
    historic = historic_stock_mean_and_std(reference_wb_path=workbook_path, price_column_header='PX_LAST', header_start_row=data_header_row, date_0=dt.datetime.strptime(str(data_sheet[sheet_announce_date_cell].value),'%Y%m%d'))
    merger = merger_stock_mean_and_std(reference_wb_path=workbook_path, price_column_header='PX_LAST', header_start_row=data_header_row, date_0=dt.datetime.strptime(str(data_sheet[sheet_announce_date_cell].value),'%Y%m%d'))

    #iterate through the rows of the data_sheet
    #NOTE: THE SHEET IS SET UP SO THAT VALUES WE'RE INTERESTED IN START AT ROW 10
    for (index, cell) in enumerate(data_sheet['A10:B{}'.format(total_rows)]):
        
        #checks to see if the cell value follows the pattern of an option description
        if (re.match(option_description_pattern_int, cell[1].value) or re.match(option_description_pattern_float, cell[1].value)) :

            #format_option_description() returns the following list:
            #[security_name, option_description, option_type, expiration_date, strike_price]
            option_data = format_option_description(cell[0].value, cell[1].value)

            #the number of days between the expiration and completion date. 
            date_diff = (option_data[3] - completion_date).days

            #if the expiration_date occurs 3 months after the completion_date, then stop creating sheets
            if date_diff >= 90:
                wb.save(workbook_path)
                print('Found contracts past {}'.format(completion_date))
                break
                #otherwise, keep creating sheets
            else:
                #check to see if the stike is within one standard deviation of the historical and merger stock mean
                if ((is_in_range(num=option_data[-1], high=historic[0]+2*historic[1], low=historic[0]-2*historic[1])) or (is_in_range(num=option_data[-1], high=merger[0]+2*merger[1], low=merger[0]-2*merger[1]))):
                    #creates a new sheet for the passed in workbook
                    new_sheet = wb.create_sheet()
                    #increment the sheet count by 1
                    sheet_count +=1
                    #/' aren't allowed in excel sheet names, so we replace them with '-' if the name contains '/' 
                    new_sheet.title = option_data[1].replace('/', '-')

                    #zip creates a tuple pair for each item of the passed in lists. this tuple can then be appended to the sheet
                    for data in zip(option_data_labels,option_data):
                        new_sheet.append(data)

                    #loop through every value of total_data_headers and add it to the worksheet at the specified data_header_row
                    for (index, value) in enumerate(total_data_headers, start= 1) :
                        new_sheet.cell(row = data_header_row,column = index ).value = value 

                    #add the BDH formula to the sheet
                    new_sheet['B{}'.format(data_header_row+1)] = abxl.add_option_BDH( security_name = option_data[0],
                                                                                  fields = data_table_header, 
                                                                                  start_date = data_sheet[sheet_start_date_cell].value,
                                                                                  end_date = 'B4',
                                                                                  optional_arg = BDH_optional_arg,
                                                                                  optional_val = BDH_optional_val)
        else:
            print('Not a valid option description. Could not create new workbook sheets for {}'.format(cell[1].value))
            continue
    
    #if the loop ends without finding contracts 2 months past the completion/termination date, save the workbook      
    wb.save(workbook_path) 
    print('Saving the workbook with {} new tabs'.format(sheet_count)) 
 

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
    try:
        strike_price = int(description_list[-1][1:])
    #if the string was a floating point number like 18.5, convert it to a float
    except:
        strike_price = float(description_list[-1][1:])

    option_data_list = [security_name, option_description, option_type, expiration_date, strike_price]

    return option_data_list


def update_workbook_data_index(workbook_path, data_start_row, index_column):
    '''
    Given a workbook, loop through all the sheets of that workbook and update the index for each sheet.
    '''
    #a regular expression for a formated option description where the strike is an integer
    option_sheet_pattern_int = re.compile(r'^\w+\s\w+\s\d{2}-\d{2}-\d{2}\s\w+$')

    #a regular expression for a formated option description where the strike is a foat
    option_sheet_pattern_float = re.compile(r'^\w+\s\w+\s\d{2}-\d{2}-\d{2}\s\w+\.\w+$')

    #a regular expression pattern for the stock sheet
    stock_sheet_pattern =re.compile(r'^\w+\s\w+\s\w+$')


    #loads an excel workbook given the file path to that workbook.
    wb = openpyxl.load_workbook(workbook_path)
    #gets a list of all the sheets in the workbook
    sheet_list = wb.get_sheet_names()

    #iterates through every sheet
    for (index, sheet_name) in enumerate(sheet_list):
        #indexing starts at 0.
        if index == 0:
            #get the announcement date from the first sheet
            sheet = wb.get_sheet_by_name(sheet_name)
            announcement_date = sheet['B5'].value
        #if the sheet_name matches the stock sheet pattern:
        if re.match(stock_sheet_pattern, sheet_name):
            #load the stock workbook and save it to the stock_sheet variable
            stock_sheet = wb.get_sheet_by_name(sheet_name)
            update_sheet_index(sheet_name= stock_sheet, date=announcement_date, start_row= data_start_row)

        #elif the sheet_name matches an options contract sheet 
        elif(re.match(option_sheet_pattern_int, sheet_name) or re.match(option_sheet_pattern_float, sheet_name)):
            option_sheet = wb.get_sheet_by_name(sheet_name)
            copy_data(reference_sheet=stock_sheet, main_sheet=option_sheet,index_start_row=data_start_row, data_column=index_column)
    wb.save(workbook_path)
    print('Indexed each sheet. Saving workbook...')


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
    '''
    Opens an Excel workbook in read_only mode, removing links to function calls and keeps just the data stored in each cell.
    '''
    wb= openpyxl.load_workbook(file_path, data_only= True)
    wb.save(file_path)


def delet_workbook_sheets(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)
    start_sheet_num = len(wb.get_sheet_names())
    #if there is more than one sheet in the workook
    if start_sheet_num > 1:
        #loop through every sheet name in the workbook except the first sheet
        for (index,sheet) in enumerate(wb.get_sheet_names()[1:]):
            #if the length of the sheet list ==1 stop deleting cells
            if len(wb.get_sheet_names()) ==1:
                break
            else:
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
    average_index = floor((end_index + start_index)/2)
    #variable for the while loop
    found = False
    while not found:
        #print(start_index, found)        
        curr_date = worksheet.cell(row=average_index, column=2).value
        if (date_0 == curr_date):
            found = True

        elif (date_0 > curr_date):
            start_index = average_index +1
            average_index = floor((end_index + start_index)/2)

        elif (date_0 < curr_date):
            end_index = average_index -1
            average_index = floor((end_index + start_index)/2)

    return average_index


def copy_data(reference_sheet, main_sheet,index_start_row,data_column):
    '''
    copies data from the reference_sheet to the main_sheet
    '''
    #gets the total rows of the reference sheet
    total_rows = reference_sheet.max_row

    #converts data_column to a list of numbers in case it was passed in as column letters
    data_column = convert_to_numbers(data_column)
    #iterate over each item in data_column:
    for (index,column_num) in enumerate(data_column):
        #iterate over each row
        for i in range(index_start_row, total_rows+1):
            main_sheet.cell(row=i, column=column_num).value = reference_sheet.cell(row=i, column=column_num).value 


def update_stock_price_sheet(workbook_path, sheet_name, stock_sheet_index, sheet_start_date_cell,sheet_announce_date_cell, sheet_end_date_cell,  data_header_row, data_table_index, data_table_header, BDH_optional_arg=None, BDH_optional_val=None ):
    '''
    Adds a sheet with stock price information to the workbook
    '''
    #load the given workbook
    wb = openpyxl.load_workbook(workbook_path)

    #gets the reference sheet
    reference_sheet = wb.get_sheet_by_name(sheet_name)
    ticker = '{} {}'.format(reference_sheet['B2'].value, reference_sheet['B3'].value)

    #create a new sheet, and makes it the second sheet in the workbook. sheet indexing starts at 0.
    new_sheet = wb.create_sheet(index=stock_sheet_index)
    #sets the title of the new worksheet
    new_sheet.title = ticker
    #basic data to be added to the sheet
    data = [('Company Name', reference_sheet['B1'].value),
            ('Company Ticker',ticker),
            ('Start Date', reference_sheet[sheet_start_date_cell].value),
            ('Announcement Date',reference_sheet[sheet_announce_date_cell].value),
            ('End Date',reference_sheet[sheet_end_date_cell].value)]

    #appends the data to the top of the spreadsheet
    for (index,data_lst) in enumerate(data):
        new_sheet.append(data_lst)

    #combines both passed lists:
    total_headers = data_table_index + data_table_header
    #set the index and column headers for the worksheet
    for (index, value) in enumerate(total_headers, start= 1):
        new_sheet.cell(row=data_header_row,column=index).value = value
        if value.upper() == ('DATE'):
            #sets the BDH function into place
            new_sheet.cell(row= data_header_row+1, column= index).value = abxl.add_option_BDH(security_name = data[1][1],
                                                                            fields = data_table_header, 
                                                                            start_date = reference_sheet[sheet_start_date_cell].value,
                                                                            end_date = reference_sheet[sheet_end_date_cell].value,
                                                                            optional_arg = BDH_optional_arg,
                                                                            optional_val = BDH_optional_val)
    #saves the newly added sheet to the workbook.
    wb.save(workbook_path)
    print('Adding stock sheet to the Workbook...')


def update_workbook_average_column(reference_wb_path, column_header, header_row, data_start_row, ignore_sheet_list=[]):
    '''
    Given the path to an excel workbook, Averages are calculated for each sheet of data
    '''
    #loads an excel workbook from the given file_path
    reference_wb = openpyxl.load_workbook(reference_wb_path, data_only=True)
    #returns a dictionary of 'sheet_names':[column data indexes] for each sheet of the given workbook
    sheet_data_columns =find_column_index_by_header(reference_wb= reference_wb, column_header= column_header, header_row= header_row)

    #removes any sheets that are ment to be ignored if provided
    if ignore_sheet_list != []:
        #iterates over every sheet name passed into ignore_sheet_list
        for index, ignore_sheet in enumerate(ignore_sheet_list):
            #removes the sheet name from the dictionary sheet_data_columns, so that it wont be iterated over next
            sheet_data_columns.pop(ignore_sheet)

    #iterate over each key(sheet_name) in sheet_data_columns:
    for (index,key) in enumerate(sheet_data_columns):
        #update the given sheet with the average column
        update_sheet_average_column(reference_wb= reference_wb, 
                                    sheet_name= key,
                                    data_columns= sheet_data_columns[key],
                                    data_start_row= data_start_row,
                                    column_header= column_header)

    #saves the excel workbook
    reference_wb.save(reference_wb_path)
    print('Saving Workbook...')


def update_sheet_average_column(reference_wb,sheet_name,data_columns, data_start_row, column_header):
    '''
    Given an excel worksheet, and a specified list of columns, averages are calcualted for each row of the data
    '''
    #loads the sheet of the reference_wb
    sheet = reference_wb.get_sheet_by_name(sheet_name)

    #gets the max row of the sheet 
    max_row = sheet.max_row

    #gets the max column of the sheet
    max_col = sheet.max_column
    
    #sets the header for the average column to the average_col_header and places it one row above the data
    sheet.cell(row=data_start_row-1, column=max_col+1).value = '{} {} {}'.format(sheet_name, 'Average', column_header)

    #iterate over each row of the workbook:
    for i in range(data_start_row,max_row+1):
        #an empty lest to store the values for the cells of the given row
        cell_values = []
        #iterate over each cell in the data column
        for (index, column_ref) in enumerate(data_columns):
            #if the value of the cell isn't 0, append it to the cell_values list
            if sheet.cell(row=i, column=column_ref).value != 0:
                cell_values.append(sheet.cell(row=i, column=column_ref).value)

        #assing the value of the average column
        #if the cell_values list is an empyt list
        if cell_values == []:
            #set the value of the cell to 0
            sheet.cell(row=i, column=max_col+1).value = 0
        #else cell_values is populated
        else:
            #set the value of the average column to the mean of the cell_values
            sheet.cell(row=i, column=max_col+1).value = statistics.mean(cell_values)


def find_column_index_by_header(reference_wb, column_header, header_row):
    '''
    Given a reference Wb, an average is calculated for the non zero cells of a specified column
    '''
    #an empty dictionary that will store the sheet_name as the key, and a list of data_columns as the key's value 
    data_columns_by_sheet= {}

    #iterate over all the sheetnames in the workbook
    for (index,sheet_name) in enumerate(reference_wb.get_sheet_names()):
        #load the given worksheet.
        sheet = reference_wb.get_sheet_by_name(sheet_name)

        #get the max_column for the worksheet:
        max_col =sheet.max_column

        #add a key in the dictionary for the given sheet
        data_columns_by_sheet.setdefault(sheet_name, [])

        #loop through all the cells in the header_row
        for i in range(max_col+1):
            #If the value in the column header matches the header_value we're searching for, then append the column index to the key's list:
            if  column_header == sheet.cell(row=header_row, column=i+1).value:
                data_columns_by_sheet[sheet_name].append(i+1)

        #if no columns were found, remove that key from the dictionary
        if data_columns_by_sheet[sheet_name] == []:
            data_columns_by_sheet.pop(sheet_name)

    #return the dictionary with the data for each sheet
    return data_columns_by_sheet


def stock_data_to_list(reference_wb,price_column_header, header_start_row, start_index, end_index):
    '''
    Given the file path to a workbook, data in a particular cell is added to a list and then the list is returned
    '''
    #returns a dictionary with {'sheet_name':[data_column]}
    data_column = find_column_index_by_header(reference_wb = reference_wb, column_header= price_column_header, header_row= header_start_row)

    #data list to store all the values:
    data_list = []

    #re to insure the key is refering to the stock sheet
    stock_sheet_pattern =re.compile(r'^\w+\s\w+\s\w+$')
    
    #iterate over all the keys in the data_column:
    for (index,key) in enumerate(data_column):
        if re.match(stock_sheet_pattern, key):
            #load the worksheet
            sheet=reference_wb.get_sheet_by_name(key)
            for i in range(start_index, end_index+1):
                if sheet.cell(row=i,column=data_column[key][0]).value !=0:
                    data_list.append(sheet.cell(row=i,column=data_column[key][0]).value)
    #return the data_list
    return data_list

def data_average(data_list):
    '''
    returns the average of a given list
    '''
    return floor(mean(data_list))


def data_standard_dev(data_list):
    '''
    returns the standard deviation of a given list
    '''
    return ceil(stdev(data=data_list))


def historic_stock_mean_and_std(reference_wb_path,price_column_header, header_start_row, date_0):
    '''
    calculates the mean and standard deviation for prices up to the announcemnt date
    '''
    #loads the workbook and the specified sheet
    wb = openpyxl.load_workbook(reference_wb_path)
    #get the second sheet in the workbook
    sheet = wb.get_sheet_by_name(wb.get_sheet_names()[1])

    total_rows=sheet.max_row

    index0=find_index_0(worksheet=sheet,start=header_start_row+1, end=total_rows, date_0=date_0)
    data_list=stock_data_to_list(reference_wb=wb, price_column_header=price_column_header, 
                                 header_start_row=header_start_row, start_index=header_start_row+1, end_index=index0)

    average = data_average(data_list)
    st_dev = data_standard_dev(data_list)

    return(average, st_dev)


def merger_stock_mean_and_std(reference_wb_path, price_column_header, header_start_row, date_0):
    '''
    calculates the mean and standard deviation for prices from the merger date to the end of the M&A
    '''
    wb = openpyxl.load_workbook(reference_wb_path)
    #get the second sheet in the workbook
    sheet = wb.get_sheet_by_name(wb.get_sheet_names()[1])

    total_rows=sheet.max_row

    index0=find_index_0(worksheet=sheet,start=header_start_row+1, end=total_rows, date_0=date_0)
    data_list=stock_data_to_list(reference_wb=wb, price_column_header=price_column_header, 
                                 header_start_row=header_start_row, start_index=index0, end_index=total_rows)

    average = data_average(data_list)
    st_dev = data_standard_dev(data_list)

    return(average, st_dev)

def is_in_range(num, high, low):
    '''
    Given a number, and a high and low range, True is returned if the number is within the range 
    '''        
    if low <=num <= high:
        return True
    else:
        return False


def fill_option_wb_empty_cells(reference_wb_path, column_start, row_start, fill_value):
    '''
    Goes through each sheet and fills in the blanks with the designated fill_vale
    '''
    #load the workbook
    wb = openpyxl.load_workbook(reference_wb_path)

    #re for options sheets with int strikes
    option_sheet_pattern_int = re.compile(r'^\w+\s\w+\s\d{2}-\d{2}-\d{2}\s\w+$')

    #re for options sheets with float strikes
    option_sheet_pattern_float = re.compile(r'^\w+\s\w+\s\d{2}-\d{2}-\d{2}\s\w+\.\w+$')

    #iterate over each sheet
    for (index,sheet_name) in enumerate(wb.get_sheet_names()):
        #if the sheet is an option sheet
        if re.match(option_sheet_pattern_int, sheet_name) or re.match(option_sheet_pattern_float, sheet_name):
            sheet = wb.get_sheet_by_name(sheet_name)
            fill_option_sheet_empty_cells(reference_sheet=sheet,column_start= column_start, row_start= row_start, fill_value= fill_value)
    
    #save the workbook:
    wb.save(reference_wb_path)
    print('Done filling empty cells with {}.'.format(fill_value))


def fill_option_sheet_empty_cells(reference_sheet, column_start, row_start, fill_value):
    '''
    Goes through a sheet and fills in the empty cells with the designated fill_value
    '''
    #get the max_rows
    total_rows=reference_sheet.max_row
    #get the max_columns
    total_columns = reference_sheet.max_column
    #iterate_over_each column:
    for i in range(column_start, total_columns+1):
        #iterate over each row:
        for j in range(row_start, total_rows+1):
            if reference_sheet.cell(row=j, column=i).value == None:
                reference_sheet.cell(row=j, column=i).value = fill_value



    





