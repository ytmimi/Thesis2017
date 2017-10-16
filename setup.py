from distutils.core import setup

setup(name='ma_option_vol',
      version='1.0',
      description='Code to help with the collection and analysis of data for my Fall 2017 Grossman School of Business Honors Thesis',
      author='Yacin Tmimi',
      author_email='ytmimi@uvm.edu',
      url='https://github.com/ytmimi/Thesis2017',
      download_url='https://github.com/ytmimi/Thesis2017',
      license='MIT',

      packages=['ma_option_vol'],
      data_files=[('company_data/sample',['M&A List A-S&P500 T-US Sample Set.xlsx'])]
     )