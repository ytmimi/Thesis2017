from setuptools import setup

setup(name='ma_option_vol',
      version='1.0.2',
      
      #project desctiption
      description='Thesis Code',
      long_description='Code to help with the collection and analysis of data for my Fall 2017 Grossman School of Business Honors Thesis',

      #author info
      author='Yacin Tmimi',
      author_email='ytmimi@uvm.edu',

      #homepage for the project
      url='https://github.com/ytmimi/Thesis2017',

      #open source license
      license='MIT',

      #local packages to be installed
      packages=['ma_option_vol', 'company_data'],

      #python modules that the code needs to run properly
      install_requires=['openpyxl'],
      
      #non python files to be included with the source distribution
      package_data={
        'company_data': ['sample/*.xlsx'],
        '': ['LICENSE.txt','README.md'],
      },
     )