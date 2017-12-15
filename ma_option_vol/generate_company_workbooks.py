import openpyxl
import datetime as dt
import os
#imports the add_bloomberg_excel_functions from the current module
import add_bloomberg_excel_functions as abxl
from update_excel_workbooks import store_data_to_txt_file
from CONSTANTS import ACQUIRER_DIR, TARGET_DIR, MERGER_SAMPLE

class Create_Company_Workbooks():
    

    def __init__(self, source_sheet_name, source_file=MERGER_SAMPLE, target_path=TARGET_DIR, acquirer_path=ACQUIRER_DIR):
        self.source_sheet_name = source_sheet_name
        self.source_file = source_file
        self.target_path = target_path
        self.acquirer_path = acquirer_path


    def create_company_workbooks(self):
        #saves the n
        wb = openpyxl.load_workbook(self.source_file)
        sheet = wb.get_sheet_by_name(self.source_sheet_name)
        
        #iterates over all the rows of the worksheet
        for (i, row) in enumerate(sheet.rows):
            #skips the first row of the worksheet because it contains column titles
            if i < 1:
                continue
            else:
                #creates the new workbooks
                self.new_target_workbook(row_data=row, target_path= self.target_path)
                self.new_acquirer_workbook(row_data=row, acquirer_path= self.acquirer_path)


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
        
        #defines a date that acts as our cut off date
        cut_off_date = dt.datetime(2012,2,15)
        #if the start_date is greater than the cut off date, then create a new file
        if start_date > cut_off_date:

            #checks that each announcement date is a weekday, if not, the date will be adjusted to the following Monday
            announcement_date = self.adjust_to_weekday(row_data[1].value.date())

            #a list of data that will be added to each newly created worksheet
            data = [['Target Name', row_data[3].value], 
                    ['Target Ticker', row_data[4].value],
                    ['Type', 'Equity'],
                    ['Start Date', start_date.date()],
                    ['Announcement Date', announcement_date],
                    ['End Date', row_data[2].value.date()],
                    ['Formated Start Date',int(str(start_date.date()).replace('-',''))],
                    ['Formated Announcement Date',int(str(announcement_date).replace('-',''))],
                    ['Formated End Date',int(str(row_data[2].value.date()).replace('-',''))]]

            #creates a new Workbook
            wb_target = openpyxl.Workbook()
            target_sheet = wb_target.get_active_sheet()
            target_sheet.title = 'Options Chain'
            
            #appends the data to the workbook        
            for (index, cell) in enumerate(target_sheet['A1:B9']):
                #tuple unpacking to set the cell values 
                (cell[0].value, cell[1].value) = data[index]
            self.get_company_options_tickers(reference_sheet=target_sheet, start_date=start_date.date(), announcement_date=announcement_date, 
                                    row=10, start_column=1, interval=90, ticker_cell='B2', type_cell='B3')
            
            self.save_new_workbook( new_workbook= wb_target, workbook_path= target_path, 
                                    file_name= row_data[3].value, start_date_str= str(row_data[1].value.date()),
                                    file_extension= 'xlsx')        
           

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
        
        #defines a date that acts as our cut off date
        cut_off_date = dt.datetime(2012,2,15)

        #if the start_date is greater than the cut off date, then create a new file
        if start_date > cut_off_date:

            #checks that each announcement date is a weekday, if not, the date will be adjusted to the following Monday
            announcement_date = self.adjust_to_weekday(row_data[1].value.date())

            #a list of data that will be added to each newly created worksheet
            data = [['Acquirer Name', row_data[6].value], 
                    ['Acquirer Ticker', row_data[7].value],
                    ['Type', 'Equity'],
                    ['Start Date', start_date.date()],
                    ['Announcement Date', announcement_date],
                    ['End Date', row_data[2].value.date()],
                    ['Formated Start Date',int(str(start_date.date()).replace('-',''))],
                    ['Formated Announcement Date',int(str(announcement_date).replace('-',''))],
                    ['Formated End Date',int(str(row_data[2].value.date()).replace('-',''))]]
            #creates a new Workbook
            wb_acquirer = openpyxl.Workbook()
            acquirer_sheet = wb_acquirer.get_active_sheet()
            acquirer_sheet.title = 'Options Chain'     
            
            #appends the data to the workbook        
            for (index, cell) in enumerate(acquirer_sheet['A1:B9']):
                #tuple unpacking to set the cell values 
                (cell[0].value, cell[1].value) = data[index]
            self.get_company_options_tickers(reference_sheet=acquirer_sheet, start_date=start_date.date(), announcement_date=announcement_date, 
                                    row=10, start_column=1, interval=90, ticker_cell='B2', type_cell='B3')
            
            #saves the workbook
            self.save_new_workbook( new_workbook= wb_acquirer, workbook_path= acquirer_path,
                                    file_name= row_data[6].value, start_date_str=str(row_data[1].value.date()),
                                    file_extension= 'xlsx')


    def get_company_options_tickers(self,reference_sheet, start_date, announcement_date, row, start_column, interval, ticker_cell, type_cell):
        #loop through and call the BDS function while to start_date+the interval is less than 1 months past the announcement date
        while start_date < (announcement_date + dt.timedelta(days=30)):
            reference_sheet.cell(row=row,column=start_column).value = abxl.add_BDS_OPT_CHAIN(ticker_cell=ticker_cell,
                                                                type_cell=type_cell, 
                                                                date_override_cell=str(start_date).replace('-',''))
            start_date += dt.timedelta(days=interval)
            start_column +=2


    def save_new_workbook(self,new_workbook,workbook_path, file_name, start_date_str, file_extension):
        #checks to see if the given workbook_path exists
        if os.path.exists(workbook_path):
            #joins the path with the file Name 'file_name_start_date.file_extension', replacing / with _ to create valid excel file names
            file_name_and_extension='{}_{}.{}'.format(file_name.replace('/','_'),start_date_str , file_extension)
            final_path = '/'.join([workbook_path,file_name_and_extension])
            #save the worksheet
            new_workbook.save(final_path)
            if workbook_path == TARGET_DIR:
                store_data_to_txt_file(file_name='target_workbooks', data='Created {}\n'.format(file_name_and_extension))
            elif workbook_path == ACQUIRER_DIR: 
                store_data_to_txt_file(file_name='acquirer_workbooks', data='Created {}\n'.format(file_name_and_extension))
        else:
            #if the path doesn't exist, create it
            os.makedirs(workbook_path, exist_ok=False)
            print('Generating file path: {}'.format(workbook_path))
            #joins the path with the file Name 'file_name_start_date.file_extension', replacing / with _ to create valid excel file names
            file_name_and_extension='{}_{}.{}'.format(file_name.replace('/','_'),start_date_str , file_extension)
            final_path = '/'.join([workbook_path,file_name_and_extension])
            #save the worksheet
            new_workbook.save(final_path)
            if workbook_path == TARGET_DIR:
                store_data_to_txt_file(file_name='target_workbooks', data='Created {}\n'.format(file_name_and_extension))
            elif workbook_path == ACQUIRER_DIR: 
                store_data_to_txt_file(file_name='acquirer_workbooks', data='Created {}\n'.format(file_name_and_extension))
    

    def adjust_to_weekday(self, date):
        '''
        A given datetime object is checked to see whether it is a weekend.  If it is, the date is adjusted to the next monday.
        '''
        #the dt.weekday() function returns a number from 0-6 corresponding to Monday-Sunday
        #if its Saturday
        if date.weekday() == 5: 
            #adjusted to Monday
            date += dt.timedelta(days=2) 
        #if its Sunday 
        elif date.weekday() == 6:
            #adjusted the date to Monday
            date += dt.timedelta(days=1) 
        #return the adjusted date
        return date 



