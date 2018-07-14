import os
import unittest
from shutil import rmtree

BASE_DIR = os.path.abspath(os.pardir)

TEST_DIR_PATH = os.path.join(BASE_DIR, 'test_files')
TEST_TARGET_PATH = os.path.join(TEST_DIR_PATH, 'target')
TEST_ACQUIRER_PATH = os.path.join(TEST_DIR_PATH, 'acquirer')

def setUpModule():
	'''Creates a temporary directory for tests'''
	os.makedirs(TEST_DIR_PATH, exist_ok=True)
	os.makedirs(TEST_TARGET_PATH, exist_ok=True)
	os.makedirs(TEST_ACQUIRER_PATH, exist_ok=True)
	print('\nCreating temporary directory:\n'+f'{TEST_DIR_PATH}\n')

def tearDownModule():
	''' Removes temporary directories'''
	rmtree(TEST_DIR_PATH)
	print('\n\nRemoving temporary directory:\n'+f'{TEST_DIR_PATH}')

def clean_up():
	'''clears test folders'''
	if len(os.listdir(TEST_TARGET_PATH)) > 0:
		for file in os.listdir(TEST_TARGET_PATH):
			remove_file(TEST_TARGET_PATH, file)

	if len(os.listdir(TEST_ACQUIRER_PATH)) > 0:
		for file in os.listdir(TEST_ACQUIRER_PATH):
			remove_file(TEST_ACQUIRER_PATH, file)

def remove_file(path, file_name):
		path = os.path.join(path,file_name)
		os.remove(path)