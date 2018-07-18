import os
import sys
import unittest
import datetime as dt

BASE_DIR = os.path.abspath(os.pardir)
path = os.path.join(BASE_DIR,'ma_option_vol')
sys.path.append(path)

from options import Option, parse_option_description

@unittest.skip('Tested')
class Test_Option(unittest.TestCase):
	@classmethod
	def setUpClass(cls):
		cls.exp_date = dt.datetime(year=2014, month=1, day=3)
		cls.call = Option('Call', 50, cls.exp_date)
		cls.put = Option('Put', 45, cls.exp_date)

	def test_print_option(self):
		call_str = self.call.__str__()
		self.assertEqual(call_str, '01/03/2014 50 Call')
		put_str = self.put.__str__()
		self.assertEqual(put_str, '01/03/2014 45 Put')

	def test_option_attributes_call(self):
		self.assertEqual(self.call.type, 'Call')
		self.assertEqual(self.call.strike, 50)
		self.assertEqual(self.call.expiration, self.exp_date)

	def test_option_attributes_put(self):
		self.assertEqual(self.put.type, 'Put')
		self.assertEqual(self.put.strike, 45)
		self.assertEqual(self.put.expiration, self.exp_date)

	def test_days_till_exp(self):
		date = dt.datetime(year=2013, month=12, day=18)
		while date < self.exp_date:
			days_call = self.call.days_till_expiration(date)
			days_put = self.put.days_till_expiration(date)
			expected_days = (self.exp_date - date).days
			self.assertEqual(days_call, expected_days)
			self.assertEqual(days_put, expected_days)
			date+=dt.timedelta(days=1)

	def test_from_description_put(self):
		decription = 'PFE US 12/20/14 P18'
		option = Option.from_description(decription)
		self.assertIsInstance(option, Option)

	def test_from_description_put(self):
		decription = 'PFE US 12/20/14 C21.5'
		option = Option.from_description(decription)
		self.assertIsInstance(option, Option)
		

	def test_implied_volatility(self):
		date = dt.datetime(year=2013, month=12, day=18)
		stock_price = 50
		option_price = 3.45
		rf_rate = 0.002
		iv = self.call.implied_volatility(date, stock_price, option_price, rf_rate)
		self.assertEqual(iv, 0.8266282080383016)

	def test_vega(self):
		date = dt.datetime(year=2013, month=12, day=18)
		stock_price = 50
		option_price = 3.45
		rf_rate = 0.002
		vega = self.call.vega(date, stock_price, option_price, rf_rate)
		self.assertEqual(vega, 0.7607798174593228)


if __name__ == '__main__':
	unittest.main()




	