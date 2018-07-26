import os
import sys
import unittest
import datetime as dt

BASE_DIR = os.path.abspath(os.pardir)
path = os.path.join(BASE_DIR,'ma_option_vol')
sys.path.append(path)

import add_bloomberg_excel_functions as abxl

class Test_Bloomberg_Functions(unittest.TestCase):
	def test_add_BDS_OPT_CHAIN(self):
		self.assertEqual(abxl.add_BDS_OPT_CHAIN(ticker_cell='B2',type_cell='B3', date_override_cell='B6'), 
			'=BDS(CONCATENATE(B2, " ", B3),"OPT_CHAIN","OPTION_CHAIN_OVERRIDE","M",CONCATENATE("SINGLE_DATE_OVERRIDE=",B6))')
		self.assertEqual(abxl.add_BDS_OPT_CHAIN(ticker_cell='AAPL US',type_cell='EQUITY', date_override_cell='20151231'),
			'=BDS(CONCATENATE("AAPL US", " ", "EQUITY"),"OPT_CHAIN","OPTION_CHAIN_OVERRIDE","M",CONCATENATE("SINGLE_DATE_OVERRIDE=","20151231"))')


	def test_add_BDP_fuction(self):
		self.assertEqual(abxl.add_BDP_fuction(security_cell='A3', field_cell='A4'),
			'=BDP(A3,A4)')
		self.assertEqual(abxl.add_BDP_fuction(security_cell='BBG00673J6L5 Equity', field_cell='SECURITY_DES'),
			'=BDP("BBG00673J6L5 Equity","SECURITY_DES")')
		self.assertEqual(abxl.add_BDP_fuction(security_cell='A3', field_cell='PX_LAST'),
			'=BDP(A3,"PX_LAST")')


	def test_add_option_BDH(self):
		self.assertEqual(abxl.add_option_BDH(security_name='B4', fields='C3', start_date='A1', end_date='A2', optional_arg=None, optional_val=None),
			'=BDH(B4,C3,A1,A2)')
		self.assertEqual(abxl.add_option_BDH(security_name='B4', fields='C3', start_date=dt.datetime(day=5, month=6, year=2013), end_date='A2', optional_arg = None, optional_val=None),
			'=BDH(B4,C3,"20130605",A2)')
		self.assertEqual(abxl.add_option_BDH(security_name='B4', fields='C3', start_date=dt.date(day=5, month=6, year=2013), end_date='A2', optional_arg = None, optional_val=None),
			'=BDH(B4,C3,"20130605",A2)')
		self.assertEqual(abxl.add_option_BDH(security_name='B4', fields='C3:C5', start_date='A1', end_date='A2', optional_arg = 'E1:E2', optional_val='F1:F2'),
			'=BDH(B4,C3:C5,A1,A2,E1:E2,F1:F2)')
		self.assertEqual(abxl.add_option_BDH(security_name='IBM US EQUITY', fields='PX_BID, PX_ASK, PX_VOLUME', start_date='20150520', end_date='20170915', optional_arg = 'Days', optional_val='W'),
			'=BDH("IBM US EQUITY","PX_BID, PX_ASK, PX_VOLUME","20150520","20170915","Days","W")')
		self.assertEqual(abxl.add_option_BDH(security_name='B4', fields=['PX_LAST','PX_BID'], start_date='20131230', end_date='A2', optional_arg = 'Days, Fill', optional_val=['W','0']),
			'=BDH(B4,"PX_LAST, PX_BID","20131230",A2,"Days, Fill","W, 0")')
		self.assertRaises(ValueError, abxl.add_option_BDH, security_name='B4', fields='C3', start_date='A1', end_date='A2', optional_arg = 'Fill', optional_val=None)
		self.assertRaises(ValueError, abxl.add_option_BDH, security_name='B4', fields='C3', start_date='A1', end_date='A2', optional_arg = None, optional_val=0)


if __name__ == "__main__":
	unittest.main()
