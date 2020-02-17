# -*- coding: utf-8 -*-
"""
Created on Thu Nov 21 15:26:59 2019

@author: ema
"""

import pandas as pd
import numpy as np
import os

################################################################################################################################
################################################################################################################################


#####sephora actual & forecast combine ####
os.chdir("C:/Users/ema/Desktop/Elvis Local/forecasting files/Sephora October")
sephora_actual=pd.read_excel("October Pass Sephora ACTUAL.xlsx",sheet_name="ACTUAL",header=0)

### IMPORTANT!!!!! FORECAST FILE --> forecast # from last month S&OP, SKU list from current month S&OP
sephora_fcst=pd.read_excel("October Pass Sephora FCST.xlsx", sheet_name="FCST",header=0)

sephora_combine=sephora_actual.merge(sephora_fcst, right_on='Sku', left_on='Item Number', how='outer')

def sku_f(row):
    if pd.isnull(row[0]):
        val=row[2]
    else:
        val=row[0]
    return val

sephora_combine['new_sku']=sephora_combine.apply(sku_f,axis=1)

columns=np.append(sephora_combine.columns.values[1],sephora_combine.columns.values[3:])
cols=np.append(columns[-1], columns[0:-1])
sephora_combine_update=sephora_combine[cols]

sephora_combine_update=sephora_combine_update.fillna(0)
sephora_combine_update.to_excel("October Pass Sephora Combine.xlsx", index=False)
#####sephora actual & forecast combine ####

################################################################################################################################
################################################################################################################################




#####Ulta actual & forecast combine ####
os.chdir("C:/Users/ema/Desktop/Elvis Local/forecasting files/Ulta October")
ulta_actual=pd.read_excel("October Pass Ulta ACTUAL.xlsx",sheet_name="ACTUAL",header=0)

### IMPORTANT!!!!! FORECAST FILE --> forecast # from last month S&OP, SKU list from current month S&OP
ulta_fcst=pd.read_excel("October Pass Ulta FCST.xlsx", sheet_name="FCST",header=0)

ulta_combine=ulta_actual.merge(ulta_fcst, right_on='Sku', left_on='Item Number', how='outer')

def sku_f(row):
    if pd.isnull(row[0]):
        val=row[2]
    else:
        val=row[0]
    return val

ulta_combine['new_sku']=ulta_combine.apply(sku_f,axis=1)

columns=np.append(ulta_combine.columns.values[1],ulta_combine.columns.values[3:])
cols=np.append(columns[-1], columns[0:-1])
ulta_combine_update=ulta_combine[cols]

ulta_combine_update=ulta_combine_update.fillna(0)
ulta_combine_update.to_excel("October Pass ulta Combine.xlsx", index=False)
##### ulta actual & forecast combine ####


################################################################################################################################
################################################################################################################################





#####EC Scott actual & forecast combine ####
os.chdir("C:/Users/ema/Desktop/Elvis Local/forecasting files/ec_scott October")
ec_scott_actual=pd.read_excel("October Pass EC Scott ACTUAL.xlsx",sheet_name="ACTUAL",header=0)

### IMPORTANT!!!!! FORECAST FILE --> forecast # from last month S&OP, SKU list from current month S&OP
ec_scott_fcst=pd.read_excel("October Pass EC Scott FCST.xlsx", sheet_name="FCST",header=0)

ec_scott_combine=ec_scott_actual.merge(ec_scott_fcst, right_on='Sku', left_on='Item Number', how='outer')

def sku_f(row):
    if pd.isnull(row[0]):
        val=row[2]
    else:
        val=row[0]
    return val

ec_scott_combine['new_sku']=ec_scott_combine.apply(sku_f,axis=1)

columns=np.append(ec_scott_combine.columns.values[1],ec_scott_combine.columns.values[3:])
cols=np.append(columns[-1], columns[0:-1])
ec_scott_combine_update=ec_scott_combine[cols]

ec_scott_combine_update=ec_scott_combine_update.fillna(0)
ec_scott_combine_update.to_excel("October Pass EC Scott Combine.xlsx", index=False)
##### EC Scott actual & forecast combine ####


################################################################################################################################
################################################################################################################################



##### Other.com actual & forecast combine ####
os.chdir("C:/Users/ema/Desktop/Elvis Local/forecasting files/Other.com October")
othercom_actual=pd.read_excel("October Pass othercom ACTUAL.xlsx",sheet_name="ACTUAL",header=0)

### IMPORTANT!!!!! FORECAST FILE --> forecast # from last month S&OP, SKU list from current month S&OP
othercom_fcst=pd.read_excel("October Pass Other.com FCST.xlsx", sheet_name="FCST",header=0)

othercom_combine=othercom_actual.merge(othercom_fcst, right_on='Sku', left_on='Item Number', how='outer')

def sku_f(row):
    if pd.isnull(row[0]):
        val=row[2]
    else:
        val=row[0]
    return val

othercom_combine['new_sku']=othercom_combine.apply(sku_f,axis=1)

columns=np.append(othercom_combine.columns.values[1],othercom_combine.columns.values[3:])
cols=np.append(columns[-1], columns[0:-1])
othercom_combine_update=othercom_combine[cols]

othercom_combine_update=othercom_combine_update.fillna(0)
othercom_combine_update.to_excel("October Pass Other.com Combine.xlsx", index=False)
#####Other.com actual & forecast combine ####

