import numpy as np
import datetime as dt
from CONSTANTS import MINUTES_PER_MONTH as N30
from CONSTANTS import MINUTES_PER_YEAR as N365
from iv_calculation import days_till_expiration


class Option_Contract_Data():
	'''
	Defines data for an option contract on a given date
	'''
	def __init__(self, option_description, exp_date, strike_price, px_last, px_bid, px_ask):
		self.option_description = option_description
		self.exp_date = exp_date
		self.strike_price = strike_price
		self.px_last = px_last
		self.px_bid = px_bid
		self.px_ask = px_ask
		if self.px_bid == 0:
			self.px_mid = self.px_ask
		elif self.px_ask == 0:
			self.px_mid = self.px_bid
		else:
			self.px_mid = round((self.px_ask + self.px_bid)/2,2)

	def set_px_mid(self, value):
		'''
		sets the mid price for the option
		'''
		self.px_mid = value




class Near_Term():
	'''
	A class containing Options Contracts. orders strikes from smallest to largest
	'''
	def __init__(self, option_dict, risk_free_rate, current_date, current_time='4:00 PM', settlement_time='8:30 AM'):
		#assumes that the values for the keys of option_dic are list of Option_Contract_Data objects
		self.current_date = current_date
		self.option_dict = self.sort_dict_by_strike(option_dict)
		self.current_time= self.convert_time(current_time)
		self.settlement_time = self.convert_time(settlement_time)
		#each option has the same expiration date, so that information is just being passed from the first contract in the call list
		self.exp_date = option_dict['call'][0].exp_date
		self.R = risk_free_rate
		self.T = self.time_to_expiration()
		self.e_RT = np.exp(self.R * self.T)
		self.F_strike = self.smallest_mid_price_diff()
		self.F = self.forward_index_level()
		#both the call and the put list will contain the same strikes, in this case the call list is used to determin k0, the stirk immideatily below F
		self.k0 = self.find_k0(option_list= self.option_dict['call'])
		#list of (strike_price, midpoint) tuples
		self.non_zero_bid_list = self.create_calculation_list()
		self.variance = self.calculate_variance()
	

	def sort_dict_by_strike(self, some_dict):
		'''
		Sorts the 'call' and 'put' list by strike
		'''
		#loop through all the keys in the dictionary, in this case, just 'call' and 'put':
		for key in some_dict.keys():
			self.sort_by_strike(some_dict[key])
		#return the sorted dictionary
		return some_dict


	def sort_by_strike(self, alist):
		'''
		Given a list of Option_Contract_Data objects, they are sorted by strike.
		This method is based on an insertion sort method
		'''
		for index in range(1,len(alist)):
			current_object = alist[index]
			current_strike = alist[index].strike_price
			position = index

			while position>0 and alist[position-1].strike_price>current_strike:
				alist[position]=alist[position-1]
				position = position-1

			alist[position]=current_object


	def convert_time(self,str_time):
		'''
		Given a string in the form 'HH:MM (AM or PM)' the appropriate 24 hour datetime object is returned.
		'''
		return dt.datetime.strptime(str_time, '%I:%M %p')


	def convert_days_to_minutes(self, num):
		'''
		Given a number of days, the number of minutes is returned
		Note: there are 1440 minutes in a single day
		'''
		return num* 1440


	def minutes_till_exp(self, settlement_time):
		'''
		Given the settlemt time, the minutes from mindnight till settlement on the settlement day are calculated
		'''
		return (settlement_time - self.convert_time('12:00 AM')).seconds / 60


	def minutes_till_midnight(self, current_time):
		'''
		Given the current_time, the minutes till midnight are returned
		'''
		return (self.convert_time('12:00 AM')- current_time).seconds / 60


	def time_to_expiration(self):
		'''
		Given the current date and the expiration date, the the minutes/year till expiration are returned

		m_current_day		Minutes remaining until midnight of the current day. Markets close at 4:00 PM ET

		m_settlement_day	Minutes from midnight until the expirtion on the settlement dat
							expiration is 8:30 am for standard monthly expirations
							expiration is 3:00 pm for standard weekly expirations

		m_other_day			Total minutes in the days between current date and expiration date				
		'''
		m_current_day = self.minutes_till_midnight(self.current_time)

		m_settlement_day = self.minutes_till_exp(self.settlement_time)

		m_other_day = self.convert_days_to_minutes(days_till_expiration(start_date=self.current_date, expiration_date=self.exp_date)-1) 

		return (m_current_day + m_settlement_day + m_other_day)/ N365


	def smallest_mid_price_diff(self):
		'''
		Returns the strike with the smallest absolute difference between the price of its respective call and put
		'''
		#creates a list of (strike price, mid price differences) tuples for each strike and midprice in both call and put lists
		diff_list = [(round(np.abs(x.px_mid - y.px_mid),2), x.strike_price) for (x,y) in zip(self.option_dict['call'], self.option_dict['put'])]

		return min(diff_list)[1] #returns just the strike price from the tuple


	def forward_index_level(self):
		'''
		strike_price		strike price where the absolute difference between the call_price and put_price is smallest

		call_price			call price associated with the given strike_price

		put_price			put price associated with the given strike price

		risk_free_rate		the bond equivalent yeild of the U.S T-Bill maturing cloest to the expiration date of the given option

		time_to_expiration	time to expiration in minutes
		'''
		call_price = self.get_mid_price_by_strike(strike_price= self.F_strike, call=True, put=False)
		put_price =  self.get_mid_price_by_strike(strike_price= self.F_strike, call=False, put=True)

		return self.F_strike +self.e_RT*(call_price - put_price)


	def get_mid_price_by_strike(self, strike_price, call=True, put=False):
		'''
		will return the mid price of a given call or put based on the strike_price
		'''
		#if searching for a call price
		if call:
			#iterate through each option contract
			for option in self.option_dict['call']:
				#if the option's strike matches the one we're searching for, then return the options mid price
				if option.strike_price == strike_price:
					return option.px_mid
		#if searching for a put price
		if put:
			#iterate through each option contract
			for option in self.option_dict['put']:
				#if the option's strike matches the one we're searching for, then return the options mid price
				if option.strike_price == strike_price:
					return option.px_mid


	def find_k0(self, option_list):
		'''
		Given F, the forward_index_level, K0 is returned.
		K0 is defined as the strike immideately below F.
		'''
		#creates a list of strike prices if the strike price is less than the forward level, F
		#uses the call list, but both call and puts have the same strikes
		below_F = [x.strike_price for x in option_list if x.strike_price < self.F ]

		#return the largest strike in the list, which will be the closest strike below F
		return max(below_F)


	def create_calculation_list(self):
		'''
		Creates a list of options to be included in the variance calculation. options are centered around 1 at the money option K0.
		Calls with strike prices > K0 are included and, puts with strike prices <= K0 are included.
		The mid price for the K0 strike is determined to be the average of the call and put with strike price of K0. When searching for options 
		to include, if two consecutive options are found to have 0 bid values, then no further options are considered beyond what has already been included
		'''
		#list of call options if their strikes are greater than self.K0
		initial_call_list = [x for x in self.option_dict['call'] if x.strike_price > self.k0]
		#list of put options if their strikes are less than or equal to self.K0
		initial_put_list = [x for x in self.option_dict['put'] if x.strike_price <= self.k0]

		#combining the call and put list, while removing zero bid options
		combined_option_list = self.remove_zero_bid(option_list= initial_put_list[::-1]) + self.remove_zero_bid(option_list= initial_call_list)

		#sort the combined_option_list
		self.sort_by_strike(combined_option_list)

		#go through the combined_option_list, and set the mid price of the k0 option to the average of the call and put mid price.
		for option in combined_option_list:
			if option.strike_price == self.k0:
				#get the mid price for the call with stirke of k0
				call_price = self.get_mid_price_by_strike(strike_price=self.k0, call=True, put=False)
				#get the mid price for the put with stirke of k0
				put_price = self.get_mid_price_by_strike(strike_price=self.k0, call=False, put=True)
				#calculate the mean
				mean_px_mid = (call_price + put_price)/2
				#set the px_mid of the given option.
				option.set_px_mid(value= mean_px_mid)

		return combined_option_list

	def remove_zero_bid(self, option_list):
		'''
		Goes through an option list and addes non zero bid options to a new list.  
		If two consecutive zero bid options are found, no further options are considered
		'''
		final_list = []
		#iterate through ever item in the give list
		for (index, item) in enumerate(option_list):
			#import pdb; pdb.set_trace()
			#if the bid price does not equal zero, then add the option to the final list
			if item.px_bid != 0:
				final_list.append(item)
			else:
				if item.px_bid == option_list[index+1].px_bid ==0:
					break
		
		return final_list


	def delta_k(self, strike_high, strike_low):
		'''
		Calculates the interval between the two given strike prices
		'''
		return (strike_high - strike_low)/2


	def option_contribution(self, strike_price, delta_strike, mid_price):
		'''
		strike_price	:given strike price, either an integer or a float

		delta_strike	:should be the average difference between the

		mid_price		:mid price for the option with the given strike price
		'''
		
		return (delta_strike/strike_price**2)*(self.e_RT)*mid_price


	def sum_option_contribution(self):
		'''
		Loops through each option and calculates that options contribution to the formula
		'''
		sum_contribution = 0
		#loop through every option that was added to the non_zero_bid_list
		for (index,option) in enumerate(self.non_zero_bid_list):
			#calculate delta_k
			#first contract, which has the lowest strike
			if option == self.non_zero_bid_list[0]:
				delta_k = self.delta_k(strike_high=self.non_zero_bid_list[index+1].strike_price, strike_low= self.non_zero_bid_list[index].strike_price)

			#last contract, which has the highest strike
			elif option == self.non_zero_bid_list[-1]:
				delta_k = self.delta_k(strike_high=self.non_zero_bid_list[index].strike_price, strike_low= self.non_zero_bid_list[index-1].strike_price)				

			else:
				delta_k = self.delta_k(strike_high=self.non_zero_bid_list[index+1].strike_price, strike_low= self.non_zero_bid_list[index-1].strike_price)
			
			sum_contribution += self.option_contribution(strike_price = option.strike_price, delta_strike= delta_k, mid_price= option.px_mid)

		return sum_contribution


	def forward_contribution(self):
		'''
		Returns the forward contribution for the given option chain
		'''
		return (1/self.T)*((self.F/self.k0)-1)**2


	def calculate_variance(self):
		'''
		returns the Variance for the entire options chain
		'''
		return (2/self.T)* self.sum_option_contribution() - self.forward_contribution()




class Next_Term(Near_Term):
	'''
	Inherits everything from the Near_Term class
	'''
	pass




class VIX_Calculation(object):
	'''
	Given a Near_Term and Next_Term object, the VIX volatility calculation is performed
	'''
	def __init__(self, Near_Term, Next_Term):
		self.T1 = Near_Term.T
		self.T2 = Next_Term.T
		self.N_T1 = self.T1 * N365
		self.N_T2 = self.T2 * N365
		self.w1 = self.calculate_weight1()
		self.w2 = self.calculate_weight2()
		self.var1 = Near_Term.variance
		self.var2 = Next_Term.variance
		self.VIX = self.calculate_VIX()


	def calculate_weight1(self):
		return self.T1*((self.N_T2 - N30)/(self.N_T2 - self.N_T1))


	def calculate_weight2(self):
		return self.T2*((N30 - self.N_T1)/(self.N_T2 - self.N_T1))


	def Near_Next_30_day_weighted_average(self):
		'''
		calculates the 30 day weighted average between the Near and Next term variance's
		'''
		x = ((self.w1 * self.var1)+(self.w2 * self.var2))*(N365/N30)
		return np.sqrt(x)


	def calculate_VIX(self):
		'''
		Calculates the VIX for a given day based on the Near_Term and Next_Term options chains
		'''
		return 100 * self.Near_Next_30_day_weighted_average()
	



