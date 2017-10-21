import openpyxl
import os
import re


def group_contracts_by_strike(workbook_path):
    '''
    Given a workbook with many option contracs, those that have the same strike price are grouped together
    Returns a dictionary with keys representing the strike and type of the option
    '''
    #regular expression to determin if the contract is a put or a call
    call = re.compile(r'[C]\d+')
    put = re.compile(r'[P]\d+')
    
    #a dictionary to store the sorted data
    options_contracts = {}

    #given the workbook_path a new excel workbook is loaded
    wb = openpyxl.load_workbook(workbook_path)
    #exclude the first sheet because that isn't an options contract
    contract_list = wb.get_sheet_names()[1:]

    #loop through each contract sheet:
    for (index, contract) in enumerate(contract_list):
        #split the sheet name by whitespace and take only the last item in the list
        #the last item will either look similar to 'C(some numbers)' or 'P(some numbers)'
        contract_type = contract.split(' ')[-1]
        
        #if the contract is a call set the default value, create a new key for the contract if it
        #doesn't already exist, increase the count by 1, and append the contract to the appropriate list 
        if re.match(call, contract_type):
            options_contracts.setdefault('call',{'count':0})
            options_contracts['call'].setdefault(contract_type, [])
            options_contracts['call']['count'] += 1
            options_contracts['call'][contract_type].append(contract)

        #if the contract is a put set the default value, create a new key for the contract if it
        #doesn't already exist, increase the count by 1, and append the contract to the appropriate list
        elif re.match(put, contract_type):
            options_contracts.setdefault('put',{'count':0})
            options_contracts['put'].setdefault(contract_type, [])
            options_contracts['put']['count'] += 1
            options_contracts['put'][contract_type].append(contract)

    #finally return the options_contracts dictionary
    return options_contracts


def create_sorted_sheet(new_workbook, reference_wb, new_sheet_title, reference_sheet_list, data_start_row, data_column, index_column):
    '''
    new_workbook          should be a new openpyxl Workbook where the data will be stored
    
    reference_wb          should be an openpyxl.workbook.Workbook containig the sheets contined in reference_sheet_list
    
    
    new_sheet_title       should be a string with the name that you would like to give the newly created sheet
            
    
    reference_sheet_list  should be a list of sheets that you would like to pull data from
    
    data_start_row        should be an integer indicating which row or the reference_sheet the data stars on
    
    data_column           should be a list of numbers, where 1=column A, 2=column B, 3=column C, and so on. 
                          specifies which columns of data are of interest from the reference_sheet contained in
                          in reference_sheet_list
    
    
    index_column          should be a list of numbers, where 1=column A, 2=column B, 3=column C, and so on.
                          specifies which columns from the reference_sheet_list[0], you want to use as your index's
                          NOTE: For accurate indexing, each sheet of data should have the same index
    
    '''
    #creates a new sheet where the combined data will be stored and names it based on the passed in new_sheet_title argument
    #If there are already worksheets in the workbook create a new sheet
    if len(new_workbook.get_sheet_names()) > 1:
        new_sheet = new_workbook.create_sheet(title = new_sheet_title )
    #If the default sheet is the only sheet in the workbook, then get it and change its name to the passed in new_sheet_title argument
    else:
        new_sheet = new_workbook.get_active_sheet()
        new_sheet.title = new_sheet_title
    
    #iterate over all the sheets in the reference_wb that were passed in reference_sheet_list
    for (index,contract) in enumerate(reference_sheet_list):  
        #if its the first iteration, update the sheet with the index_colum and the data_column
        if index ==0:
            #load in the new worksheet.
            data_sheet = reference_wb.get_sheet_by_name(contract)
            
            #get the max rows of the loaded worksheet
            max_row = data_sheet.max_row
            
            #iterate over the columns and the rows that we're interested in
            for (index,column_num) in enumerate(index_column+data_column):
                #find the max_column of the new_worksheet
                max_col = new_sheet.max_column
                
                #if we reach the fist item in the data column set a header for the contract name
                if column_num == data_column[0]:
                    #the header is set at the current max_col+1, which data will go under
                    new_sheet.cell(row= data_start_row-1, column= max_col+1).value = contract
                    
                    
                #iterate over row indexes from data_start_row to max_row+1
                for i in range(data_start_row, max_row+1):
                    #if its the first iteration of adding data to the sheet
                    if index == 0:
                        new_sheet.cell(row=i, column=max_col).value = data_sheet.cell(row= i, column= column_num).value
                    
                    else:
                        new_sheet.cell(row=i, column=max_col+1).value = data_sheet.cell(row= i, column= column_num).value
                                

        #else: just grab the data_column    
        else:
            #load in the new worksheet.
            data_sheet = wb.get_sheet_by_name(contract)
            
            #get the max rows of the loaded worksheet
            max_row = data_sheet.max_row
            
            #iterate over the columns and rows that we're interested in
            for (index,column_num) in enumerate(data_column):
                #find the max_column of the new_worksheet
                max_col = new_sheet.max_column
                
                #if we reach the fist item in the data column set a header for the contract name
                if column_num == data_column[0]:
                    #the header is set at the current max_col+1, which data will go under
                    new_sheet.cell(row= data_start_row-1, column=max_col+1).value = contract
                
                #iterate over row indexes from data_start_row to max_row+1
                for i in range(data_start_row, max_row+1):
                    new_sheet.cell(row=i, column=max_col+1).value = data_sheet.cell(row= i, column= column_num).value



#work on this
def save_new_workbook(new_workbook,workbook_path, file_name, file_extension):
        #checks to see if the given workbook_path exists



        if os.path.exists(workbook_path):
            #joins the path with the file Name 'file_name.file_extension', replacing / with _ to create valid excel file names
            final_path = '/'.join([workbook_path,'{}.{}'.format(file_name.replace('/','_'), file_extension)])
            #save the worksheet
            new_workbook.save(final_path)    
        else:
            #if the path doesn't exist, create it
            os.makedirs(workbook_path, exist_ok=False)
            print('Generating file path: {}'.format(workbook_path))
            final_path = '/'.join([workbook_path,'{}.{}'.format(file_name.replace('/','_'), file_extension)])
            #save the worksheet
            new_workbook.save(final_path)












