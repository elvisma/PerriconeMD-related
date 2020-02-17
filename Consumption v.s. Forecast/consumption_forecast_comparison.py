# -*- coding: utf-8 -*-
"""
Created on Mon Dec  2 09:56:20 2019

@author: ema
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook
import os

#### need to modify every pass####
######################################################################################################################
################### haas been modified to include ORDER QTY, getting ready for FILL RATE calculation #################
######################################################################################################################
mon_name=['February 20']
rolling_index='Feb Rolling 2 weeks'


os.chdir("C:/Users/ema/Desktop/projects/consumption_forecast_inventory")
mtd_shipment=pd.read_excel("IDC shipment Accounts Mapping Tool.xlsx", sheet_name='Pivot Shipment Overview',header=0)

lm_forecast=pd.read_excel('S&OP LM Publication.xlsx',sheet_name='Channel Forecast', header=0)
lm_forecast.rename(columns={'\xa0Account\xa0':'Account'}, inplace=True)
col_names=['Account','Sku','Description']


total_cols=col_names+mon_name
monthly_forecast=lm_forecast[total_cols]
monthly_forecast=monthly_forecast.groupby(['Account','Sku','Description'], as_index=False).sum()

df_merge=pd.merge(monthly_forecast, mtd_shipment, how='outer', left_on=['Sku','Account'], right_on=['PartNum','Forecast Accounts (Final)'])


def account(row):
    if pd.isnull(row[0]):
        val=row[5]
    else:
        val=row[0]
    return val

df_merge['forecasted_account']=df_merge.apply(account, axis=1)


def sku(row):
    if pd.isnull(row[1]):
        val=row[4]
    else:
        val=row[1]
    return val

df_merge['Item #']=df_merge.apply(sku, axis=1)

def description(row):
    if pd.isnull(row[2]):
        val=row[6]
    else:
        val=row[2]
    return val

df_merge['item_description']=df_merge.apply(description, axis=1)
df_merge['rolling_index']=rolling_index

os.chdir("C:/Users/ema/Desktop/Elvis Local/data mapping & price information")
franchise_mapping=pd.read_excel("SKU to franchise map.xlsx", header=0)
franchise_map=franchise_mapping[['Sku','Franchise']].drop_duplicates(subset=['Sku'])  ### to avoid left outer join result in table longer than left table
df_merged=df_merge.merge(franchise_map, how='left', left_on='Item #', right_on='Sku')
                   

#filtered_cols=['rolling_index','forecasted_account','Item #','item_description','Franchise','Sum of ShipQty',"".join(str(x) for x in mon_name)]
filtered_cols=['rolling_index','forecasted_account','Item #','item_description','Franchise','Sum of ShipQty','Sum of OrderQty',"".join(str(x) for x in mon_name)]
data_to_load=df_merged[filtered_cols]
#data_to_load.iloc[:,-2:]=data_to_load.iloc[:,-2:].fillna(value=0)            
data_to_load.iloc[:,-3:]=data_to_load.iloc[:,-3:].fillna(value=0)                    


#book=load_workbook('Consumption_Shipment_Comparison.xlsx')
#writer = pd.ExcelWriter('Consumption_Shipment_Comparison.xlsx', engine='openpyxl') 
#writer.book = book
#writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
#df_merge.to_excel(writer, rolling_index,index=False)
#writer.save()
os.chdir("C:/Users/ema/Desktop/projects/consumption_forecast_inventory")
data_to_load.to_excel("data_to_load.xlsx", index=False)
