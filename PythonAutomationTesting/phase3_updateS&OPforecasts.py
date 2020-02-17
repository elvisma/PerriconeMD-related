# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 15:46:11 2019

@author: ema
"""

import pandas as pd
import numpy as np
import os

### laod the book2 file ###
os.chdir("O:/Supply Chain/Demand Planning/Forecast Master/October S&OP")
book2_df=pd.read_excel("October SOP Roll Up Master With Notes - Book 2.xlsx", sheet_name='Sheet1', header=0)

### load the forecast file ###

os.chdir("C:/Users/ema/Desktop/Elvis Local/forecasting files")
forecast_combine=pd.DataFrame()
book=pd.ExcelFile("forecast_to_load.xlsx")
for sheet in book.sheet_names:
    df=pd.read_excel(book, sheet_name=sheet)
    forecast_combine=forecast_combine.append(df)

new_book2=forecast_combine.iloc[:,:6].merge(book2_df,right_on="Concatenate", left_on="concat",how='outer')

## need to keep working on it
#def f(row):
#    if row['concat'] is None:
 #       val=row['Concatenate']
  #  else:
  #      val=row['concat']
  #  return val

#new_book2['new_concat']=new_book2.apply(f, axis=1)
  
###
def concat(row):
    if pd.isnull(row[0]):
        val=row[6]
    else:
        val=row[0]
    return val

new_book2['new_concat']=new_book2.apply(concat, axis=1)



def channel(row):
    if pd.isnull(row[1]):
        val=row[7]
    else:
        val=row[1]
    return val

new_book2['new_channel']=new_book2.apply(channel, axis=1)


def account(row):
    if pd.isnull(row[2]):
        val=row[8]
    else:
        val=row[2]
    return val

new_book2['new_account']=new_book2.apply(account, axis=1)


def sku(row):
    if pd.isnull(row[3]):
        val=row[9]
    else:
        val=row[3]
    return val

new_book2['new_sku']=new_book2.apply(sku, axis=1)


def description(row):
    if pd.isnull(row[4]):
        val=row[10]
    else:
        val=row[4]
    return val

new_book2['new_description']=new_book2.apply(description, axis=1)


def franchise(row):
    if pd.isnull(row[5]):
        val=row[11]
    else:
        val=row[5]
    return val

new_book2['new_franchise']=new_book2.apply(franchise, axis=1)   

##
column_list=list(new_book2.columns.values[-6:])+list(new_book2.columns.values[12:-6])

new_book2_update=new_book2[column_list]
new_book2_update=new_book2_update.sort_values(['new_channel','new_account','new_sku'])

os.chdir("C:/Users/ema/Desktop/projects")
new_book2_update.to_excel('new_book2_to_update.xlsx', index=False)  #### finish the mapping
forecast_combine.to_excel('forecast_combine.xlsx', index=False)


