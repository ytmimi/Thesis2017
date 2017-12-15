#imports to get the files
import os
import sys
import openpyxl
import pandas as pd
from test_path import AbbVie_path, AbbVie_ATM_path
parent_path = os.path.abspath(os.pardir)
path = os.path.join(parent_path,'ma_option_vol')
#adds the file path for the ma_options_vol module to the path that python will search in order to look for modules
sys.path.append(path)

import create_atm_vol_series as cavs


#cavs.create_atm_vol_workbook(AbbVie_path)

cavs.create_average_vol_sheet(AbbVie_ATM_path)
#cavs.add_mean_and_market_model(AbbVie_ATM_path)