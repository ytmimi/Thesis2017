import os
import zipfile

def get_dir_xlsx_files(path):
	'''Returns a list of .xlsx files in the given directory'''
	return [os.path.join(path, file) for file in os.listdir(path) 
			if file[-5:]=='.xlsx']

def zip_xlsx_files(dir_path, zip_path):
	''' 
	dir_path: path to the directory containing .xlsx files
	zip_path: full path including file name to store the zip file
	writes .xlsx files from dir_path to a zip file in zip_path
	return zipFile object
	'''
	if zip_path[-4:]!='.zip':
		raise ValueError('zip_path must include the full path and file name to store the zip file')
	xlsx_files = get_dir_xlsx_files(dir_path)
	with zipfile.ZipFile(zip_path, 'x') as myzip:
		for file in xlsx_files:
			file_name = file.split('/')[-1]
			myzip.write(file, file_name)
	return myzip


