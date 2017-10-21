import openpyxl
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


def new_workbook_by_strike(workbook_path)

    contracts = group_contracts_by_strike(workbook_path= workbook_path)

    #create a new sheet
    #skips the first key of the list, becuase that is the count key
    for (index,key) in enumerate(contracts['put'].keys()):
        if index > 0:
            
            #loop through the sheet names organized in each key's list
            for index,contract in enumerate(contracts['put'][key]):
                print(contract)
            break












