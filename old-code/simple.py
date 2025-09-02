import numpy as np

import pandas as pd

import time

import os

import glob

 

#import selenium

import openpyxl

from datetime import datetime

 

 

dp_new = pd.read_excel('C:\\Users\\TandP\projects\\llm_engineering\\DAILY_PRICING _new_r.xlsx',engine='openpyxl' )

 

templatee = pd.read_excel('C:\\Users\\TandP\projects\\llm_engineering\\DAILY_PRICING_ HUDSON_TEMPLATE_2021.xlsx', engine='openpyxl')

 

tp1 = templatee.copy()

 

tp2 = tp1.drop(columns=['Unnamed: 0', 'Start Date', 'Zone', 'Load Factor', '6', '12', '18',

       '24', '30', '36', '48', '60', 'Unnamed: 12', 'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16', 'Unnamed: 17'])

 

tp2.columns = dp_new.columns

 

tp2.loc[0,'ID'] = dp_new['ID'][dp_new.shape[0]-1]+1   #  x1 #   dp_new['ID'][dp_new.shape[0]-1]+1  #  [25080]+1

 

tp2['ID'] = range(int(tp2['ID'][0]), int(tp2['ID'][0]) + len(tp2))

 

 

# tp2['Date'] =  pd.to_datetime(dp_new['Date'] )       #  df1   dp_new['Date'].dt.strftime(date_format)

 

tp2 =tp2.astype({'Term': 'float', 'Min_MWh': 'float','Max_MWh': 'float', 'Daily_No_Ruc': 'float', 'RUC_Nodal': 'float', 'Daily': 'float', 'Com_Disc': 'float',

'HOA_Disc': 'float', 'Broker_Fee': 'float', 'Meter_Fee': 'float', 'Max_Meters': 'float'})

 

tp3 = tp2.copy()

 

tp3['REP1'] = tp2['Load'].copy()

tp3['Load'] = tp2['REP1'].copy()

 

result = pd.concat([dp_new, tp3], axis=0, ignore_index=True)

 

print(dp_new.shape)

print(tp2.shape)

print(tp3.shape)

print(result.shape)