import datetime as dt
from wallstreet.blackandscholes import BlackandScholes as BS

class Option:
	def __init__(self, type_, strike, expiration):
		self.type = type_
		self.strike = strike
		self.expiration = expiration

	def __str__(self):
		return f'{self.expiration :%m/%d/%Y} {self.strike} {self.type}'

	@staticmethod
	def parse_option_description(description):
		'''
		description should be a string that looks similar to 'PFE US 12/20/14 P18'
		return formatted option data
		'''
		option_data = description.split(' ')
		if option_data[-1][0] == 'P':
			option_type = 'Put'
		elif option_data[-1][0] == 'C':
			option_type = 'Call'
		expiration_date = dt.datetime.strptime(option_data[2], '%m/%d/%y')
		strike_price = float(option_data[-1][1:])
		return [option_type, expiration_date, strike_price]

	@classmethod
	def from_description(cls, description):
		type_, exp, strike = cls.parse_option_description(description)
		return cls(type_, strike, exp)

	def days_till_expiration(self, date):
		return (self.expiration-date).days

	def implied_volatility(self, date, stock_price, option_price, rf_rate, div_yeild=0):
		days = self.days_till_expiration(date)/365
		option = BS(S=stock_price, K=self.strike, T=days, price=option_price, 
					r=rf_rate, option=self.type, q=div_yeild)
		return option.impvol

	def vega(self, date, stock_price, option_price, rf_rate, div_yeild=0):
		days = self.days_till_expiration(date)
		option = BS(S=stock_price, K=self.strike, T=days, price=option_price, 
					r=rf_rate, option=self.type, q=div_yeild)
		return option.vega()