import os
import unittest
from shutil import rmtree

BASE_DIR = os.path.abspath(os.pardir)

class Test_Base(unittest.TestCase):
	@classmethod
	def setUpClass(cls):
		'''
		Creates a temporary folder to store the files created by the tests
		'''
		cls.test_dir_path = os.path.join(BASE_DIR, 'test_files')
		os.makedirs(cls.test_dir_path, exist_ok=True)
		print(f'\nCreating Temporary Directory \n{cls.test_dir_path}\n')
		cls.target_path = os.path.join(cls.test_dir_path, 'target')
		os.makedirs(cls.target_path, exist_ok=True)
		cls.acquirer_path = os.path.join(cls.test_dir_path, 'acquirer')
		os.makedirs(cls.acquirer_path, exist_ok=True)
			
	@classmethod
	def tearDownClass(cls):
		''' 
		Removes the temporary folder and the files that were created in 
		SetUpClass
		'''
		rmtree(cls.test_dir_path)
		print(f'\nDeleting Temporary Directory \n{cls.test_dir_path}\n')

	def tearDown(self):
		''' clears test folders after each test '''
		if len(os.listdir(self.target_path)) > 0:
			for file in os.listdir(self.target_path):
				self.remove_file(self.target_path, file)

		if len(os.listdir(self.acquirer_path)) > 0:
			for file in os.listdir(self.acquirer_path):
				self.remove_file(self.acquirer_path, file)

	def remove_file(self, path, file_name):
		path = os.path.join(path,file_name)
		os.remove(path)