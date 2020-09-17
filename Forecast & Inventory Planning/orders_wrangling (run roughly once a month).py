# -*- coding: utf-8 -*-
"""
Created on Thu Jul 23 12:34:09 2020

@author: ema
"""

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
sop_pass='July'
sop_mon=['Aug 20']
runout_mons=['Sep 20','Oct 20']
year_total=['2020 TOTAL']


###################################################################################################################
###################################################################################################################
###################################################################################################################
col_names=['Account','Sku']
#filter_indicator='MTD'
os.chdir("C:/Users/ema/Desktop/projects/Forecast_Inventory_Report")


orders_shipped=pd.read_excel("shift_MTD_consumption_breakdown.xlsx", sheet_name='Shipped',header=1,usecols="A:E")
orders_processed=pd.read_excel("shift_MTD_consumption_breakdown.xlsx", sheet_name='Processed',header=1,usecols="A:E")

total_orders=pd.concat([orders_shipped, orders_processed],axis=0)
#TM_orders=total_orders[total_orders['Indicator'].str.contains(filter_indicator)]
TM_orders=total_orders
TM_orders=TM_orders.groupby(['Account','Sku'],as_index=False).sum()
TM_orders['Sku']=TM_orders['Sku'].apply(str)

###### last month's forecast (form M-1 S&OP forecast )##################
os.chdir("O:/Supply Chain/Demand Planning")
old_sop=pd.read_excel("S&OP Roll Up - "+sop_pass+" 2020 Publication.xlsx", sheet_name='Channel Forecast', header=0)
sop_cols=col_names+sop_mon
sop_forecast=old_sop[sop_cols]
sop_forecast['Sku']=sop_forecast['Sku'].apply(str)

###### this month & next month forecast (from Runout Tool) #############
os.chdir("O:/Supply Chain/Demand Planning/Forecast Master")
current_runout=pd.read_excel("Master Forecast.xlsm", sheet_name='Channel Forecast', header=0)
runout_cols=col_names+runout_mons+year_total
runout_forecast=current_runout[runout_cols]
runout_forecast=runout_forecast.iloc[1:,:]
runout_forecast[runout_mons[0]]=pd.to_numeric(runout_forecast[runout_mons[0]])
runout_forecast[runout_mons[1]]=pd.to_numeric(runout_forecast[runout_mons[1]])
runout_forecast[year_total[0]]=pd.to_numeric(runout_forecast[year_total[0]])


################# combine  2 forecast sources ######################################
forecast_total=pd.merge(sop_forecast, runout_forecast, how='outer', on=['Sku','Account'])
forecast_total.iloc[:,-4:]=forecast_total.iloc[:,-4:].fillna(0)
forecast_total=forecast_total.groupby(['Account','Sku'], as_index=False).sum()

############### BOM list & UK Kitting demand breakdown ##################
os.chdir("C:/Users/ema/Desktop/projects/Perricone Dashboard")
bom_list=pd.read_excel("SmartList BOM Extract.xlsx", sheet_name='Current BOM Extract', haeder=0)
bom_cols=['Item_Number_FGI','CMPTITNM_C','CMPITQTY_C']
bom_list=bom_list[bom_cols]

forecast_total['Sku']=forecast_total['Sku'].apply(str)
bom_list['Item_Number_FGI']=bom_list['Item_Number_FGI'].apply(str)
bom_list['CMPTITNM_C']=bom_list['CMPTITNM_C'].apply(str)

UK_SKUs=forecast_total[forecast_total['Account']=='UK']['Sku'].append(TM_orders[TM_orders['Account']=='UK']['Sku']).reset_index(drop=True).astype(str)

UK_SKUs.drop_duplicates(inplace=True)

df_uk=pd.DataFrame()
for sku in UK_SKUs:
    df1=bom_list[bom_list['CMPTITNM_C']==sku]
    df2=forecast_total[(forecast_total['Account']=='UK')&(forecast_total['Sku'].isin(df1.Item_Number_FGI))]
    df3=pd.merge(df1,df2,how='right',left_on='Item_Number_FGI', right_on='Sku')
    df3=df3.fillna(0)
    df3.iloc[:,-3:]=df3.iloc[:,-3:].multiply(df3['CMPITQTY_C'],axis='index')
    df4=pd.concat([df3[['Account','CMPTITNM_C']],df3.iloc[:,-3:]],axis=1)
    df4.rename(columns={'CMPTITNM_C':'Sku'},inplace=True)
    
    if len(df4)==0:
        continue
    else:
        df_uk=df_uk.append(df4)
  

forecast_total=forecast_total.append(df_uk)    
forecast_total=forecast_total.groupby(['Account','Sku'], as_index=False).sum()

df_merge=pd.merge(forecast_total, TM_orders, how='outer', left_on=['Sku','Account'], right_on=['Sku','Account'])

df_merge.iloc[:,-5:]=df_merge.iloc[:,-5:].fillna(0)
df_merge = df_merge.sort_values(by='QTY',ascending=False)


df_merge.rename(columns={'QTY':'MTD & shifts('+runout_mons[0]+')'},inplace=True)



os.chdir("C:/Users/ema/Desktop/Data Master")
desc_mapping=pd.read_excel("product_table.xlsx", header=0)
desc_map=desc_mapping.drop_duplicates(subset=['Sku'])  ### to avoid left outer join result in table longer than left table
desc_map['Sku']=desc_map['Sku'].apply(str)
df_merged=df_merge.merge(desc_map, how='left', left_on='Sku', right_on='Sku')
                   

                         
                         
#filtered_cols=['rolling_index','forecasted_account','Item #','item_description','Franchise','Sum of ShipQty',"".join(str(x) for x in mon_name)]
#filtered_cols=['rolling_index','forecasted_account','Item #','item_description','Franchise','Sum of ShipQty','Sum of OrderQty',"".join(str(x) for x in mon_name)]
#data_to_load=df_merged.drop(['Description_x'],axis=1)
#data_to_load.rename(columns={'Description_y':'Description'},inplace=True)

cols=df_merged.columns.tolist()
order=[0,1,-3,-2,-1,2,5,4,3,6]


new_cols=[cols[i] for i in order]
data_to_load=df_merged[new_cols]         


##############################################################################################################################
################ !!!!!!!!!!!! need to adjust the 3 months sequence every time !!!!!!!!!!!!!!##################################
##############################################################################################################################

os.chdir("C:/Users/ema/Desktop/projects/Forecast_Inventory_Report")
data_to_load.to_excel("orders_wrangling.xlsx", index=False)


##############################################################################################################################
################ !!!!!!!!!!!! need to adjust the 3 months sequence every time !!!!!!!!!!!!!!##################################
##############################################################################################################################