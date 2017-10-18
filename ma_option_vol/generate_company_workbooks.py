#imports for the module
import openpyxl
import datetime as dt
import os
#imports the add_bloomberg_excel_functions from the current module
import add_bloomberg_excel_functions as abxl

class Create_Company_Workbooks():
    

    def __init__(self, source_file, target_path, acquirer_path):
        self.source_file = source_file
        self.target_path = target_path
        self.acquirer_path = acquirer_path


    def create_company_workbooks(self):
        #saves the n
        wb = openpyxl.load_workbook(self.source_file)
        sheet = wb.get_sheet_by_name('Filtered Sample Set')
        
        #iterates over all the rows of the worksheet
        for (i, row) in enumerate(sheet.rows):
            #skips the first row of the worksheet because it contains column titles
            if i < 1:
                continue
            else:
                #creates the new workbooks
                self.new_target_workbook(row_data=row, target_path= self.target_path)
                self.new_acquirer_workbook(row_data=row, acquirer_path= self.acquirer_path)
                break #<----remember to remove this after done testing
        print('\nDone creating company files.')


    def new_target_workbook(self,row_data, target_path):
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
            #tuple unpacking to set the cell values 
            (cell[0].value, cell[1].value) = data[index]
        target_sheet['A10'] = abxl.add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B7')
        
        self.save_new_workbook( new_workbook= wb_target, workbook_path= target_path, 
                                file_name= row_data[3].value, file_extension= 'xlsx')        
       

    def new_acquirer_workbook(self,row_data, acquirer_path):
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
        acquirer_sheet = wb_acquirer.get_active_sheet()
        acquirer_sheet.title = 'Options Chain'     
        
        #appends the data to the workbook        
        for (index, cell) in enumerate(acquirer_sheet['A1:B8']):
            #tuple unpacking to set the cell values 
            (cell[0].value, cell[1].value) = data[index]
        acquirer_sheet['A10'] = abxl.add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B7')
        
        #saves the workbook
        self.save_new_workbook( new_workbook= wb_acquirer, workbook_path= acquirer_path,
                                file_name= row_data[6].value, file_extension= 'xlsx')


    def save_new_workbook(self,new_workbook,workbook_path, file_name, file_extension):
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



