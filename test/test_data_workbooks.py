import os
import sys
import unittest
import random
import datetime as dt
from openpyxl import Workbook, load_workbook

BASE_DIR = os.path.abspath(os.pardir)
path = os.path.join(BASE_DIR,'ma_option_vol')
sys.path.append(path)

from data_workbooks import Data_WorkSheet, Merger_Sample_Data
from base_test import Test_Base


class Test_Data_WorkSheet(Test_Base):
	def setUp(self):
		self.path = 'samples/test_sample.xlsx'
		self.sheet_name = 'Filtered Sample Set'
		self.wb = load_workbook(self.path)
		self.data_ws = Data_WorkSheet(self.wb, self.sheet_name)

	def tearDown(slef):
		super().tearDown()

	def test_ws_length_property(self):
		ws = load_workbook(self.path)[self.sheet_name]
		self.assertEqual(ws.max_row, self.data_ws.ws_length)

	def test_ws_width_property(self):
		ws = load_workbook(self.path)[self.sheet_name]
		self.assertEqual(ws.max_column, self.data_ws.ws_width)		

	def test_column_headers(self):
		headers = []
		ws = load_workbook(self.path)[self.sheet_name]
		for i in range(ws.max_column):
			data = ws.cell(row=1, column=i+1).value
			headers.append(data)
		self.assertEqual(headers, self.data_ws.headers)

	def test_row_values(self):
		data = {}
		row = 2
		for i, header in enumerate(self.data_ws.headers, start=1):
			value = self.data_ws.ws.cell(row=row, column=i).value
			data[i] = {'column': header, 'value':value}
		self.assertEqual(data, self.data_ws.row_values(2))


	def test_row_values_exception(self):
		with self.assertRaises(IndexError):
			index = self.data_ws.ws_length+1
			self.data_ws.row_values(index)

	def test_get_value(self):
		for i, item in enumerate(self.data_ws.headers, start=1):
			value = self.data_ws.get_value(1, i)
			#test that providing the row and column gets you the correct cell value
			self.assertEqual(item, value)

	def test_get_value_col_exception(self):
		with self.assertRaises(IndexError):
			index = self.data_ws.ws_width+1
			self.data_ws.get_value(row_index=2, col_index=index)

	def test_get_value_row_exception(self):
		with self.assertRaises(IndexError):
			row = self.data_ws.ws_length+1
			self.data_ws.get_value(row_index=row, col_index=2)

	def test_set_value(self):
		self.data_ws.set_value(10, 10, "Test Value")
		self.assertEqual(self.data_ws.ws.cell(row=10, column=10).value, "Test Value")


class Test_Merger_Sample_Data(Test_Base):
	def setUp(self):
		self.path = 'samples/test_sample.xlsx'
		self.sheet_name = 'Filtered Sample Set'
		self.wb = load_workbook(self.path)
		self.data_ws = Merger_Sample_Data(self.wb, self.sheet_name)

	def tearDown(self):
		super().tearDown()

	def test_row_values_all(self):
		row=3
		output = {}
		for i, head in enumerate(self.data_ws.headers, start=1):
			output[i] = {
				'column':head, 
				'value':self.data_ws.ws.cell(row=row, column=i).value,}
		#call to the function being tested		
		row_vals = self.data_ws.row_values(row)
		#quick check to see if the output is the same length
		self.assertEqual(len(output.keys()), len(row_vals.keys()))
		#check that the function output return the correct data
		for key in output.keys():
			self.assertEqual(output[key], row_vals[key])
		
	def test_row_values_include(self):
		output = {}
		row = 2
		for i, head in enumerate(self.data_ws.headers, start=1):
			output[i] = {
				'column':head, 
				'value':self.data_ws.ws.cell(row=row, column=i).value,}
		include = self.data_ws.headers[:5]
		output = {key:value for key, value in output.items() if value['column'] in include}
		values = self.data_ws.row_values(row, include)
		self.assertEqual(output, values)

	def test_row_values_include_random(self):
		output = {}
		row = 2
		for i, head in enumerate(self.data_ws.headers, start=1):
			output[i] = {
				'column':head, 
				'value':self.data_ws.ws.cell(row=row, column=i).value,}
		#randomly choose 5 heders
		include = random.choices(self.data_ws.headers, k=5)
		output = {key:value for key, value in output.items() if value['column'] in include}
		values = self.data_ws.row_values(row, include)
		#test that the output was filtered by the include list
		self.assertEqual(output, values)





if __name__ == "__main__":
	unittest.main()