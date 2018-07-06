import openpyxl
import datetime as dt
from CONSTANTS import MERGER_SAMPLE, TREASURY_WORKBOOK_PATH, VIX_INDEX_PATH


class Data_WorkSheet:
	def __init__(self, workbook, sheet_name, column_header_index=1):
		self.wb = workbook
		self.ws = self.wb[sheet_name]
		self.headers = self.column_headers(column_header_index)

	@property
	def ws_length(self):
		return self.ws.max_row

	@property
	def ws_width(self):
		return self.ws.max_column

	def column_headers(self, row=1):
		headers = []
		for i in range(self.ws_width):
			headers.append(self.ws.cell(row=row, column=i+1).value)
		return headers

	def row_values(self, row_index):
		if row_index <= self.ws_length:
			values = {}
			for i, col in enumerate(self.headers, start=1):
				value = self.ws.cell(row=row_index, column=i).value
				values[i] = {'column':col, 'value':value}
		else:
			raise IndexError(f'The sheet has {self.ws_length} rows. Unable to get data from row {row_index}')
		return values

	def get_value(self, row_index, col_index):
		try:
			data = self.row_values(row_index)
		except IndexError:
			print()
			raise IndexError
		if col_index <= self.ws_width:
			return data[col_index]['value']
		else:
			raise IndexError(f'The sheet has {self.ws_width} columns. Unable to get data from column {col_index}')
		return data

	def set_value(self, row_index, col_index, value):
		self.ws.cell(row=row_index, column=col_index).value = value


class Merger_Sample_Data(Data_WorkSheet):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

	def row_values(self, row_index, include=[]):
		'''
		Returns a dictionary of key value pairs for the cell values at the given index
		include: a list of column numbers to return in the response
		'''
		data = super().row_values(row_index)
		values = [item['value'] for item in data.values()]
		if len(include) > 0:
			# dict_variable = {key:value for (key,value) in dictonary.items()}
			output = {key:value for key, value in data.items() if value['column'] in include}
			return output
		return data


class Treasury_Sample_Data(Data_WorkSheet):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)


class VIX_Sample_Data(Data_WorkSheet):
	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)

		







