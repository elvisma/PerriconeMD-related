# -*- coding: utf-8 -*-
"""
Created on Wed Oct 30 16:17:32 2019

@author: ema
"""

import pandas as pd
import numpy as np
import os


################################################################################################################################################
############### Make sure before you actualize data, save last pass forecast numbers (for the column/month that will be replaced)###############
################################################################################################################################################



##### update actual sales #####
### load monthly _financial_data_tool file
os.chdir("C:/Users/ema/Desktop/Elvis Local/actual files/monthly data convert")
actual_df=pd.read_excel("monthly_financial_data_tool.xlsx", sheet_name='pivot',header=0)


######  For all accounts except for PMD.com ###########

os.chdir("C:/Users/ema/N.V. PERRICONE LLC/Michael Li - Runout Tool")
runout_df=pd.read_excel("Total Demand v6.4.xlsx", sheet_name='S&OP Fcst (Channel)', header=0)
### change "Account" Column name
runout_df.rename(columns={runout_df.columns[1]:runout_df.columns[1].strip()}, inplace=True)
runout_df['Sku_str']=runout_df['Sku'].astype(str)
runout_df['Concatenate']=runout_df['Account']+runout_df['Sku_str']

cols=runout_df.columns.tolist()
cols=cols[-1:]+cols[:-2]
runout_df=runout_df[cols]

runout_list=["Ulta","Sephora","Other.com","EC Scott","Neiman Marcus","Dillard's","Amazon","Guthy Renker","QVC","Sample","Zulily","TSC","UK","APAC+LA","Nordstrom","Costco","Lord & Taylor","Macy's","Bloomingdales"]


actual_df2=actual_df[actual_df['Account'].isin(runout_list)]
actual_df2.drop_duplicates(inplace=True)

####

runout_newdf=actual_df2[['concat','Sum of QTY']].merge(runout_df, right_on="Concatenate", left_on="concat",how='outer')





def sku_f(row):
    if pd.isnull(row[0]):
        val=row[2]
    else:
        val=row[0]
    return val

runout_newdf['new_concat']=runout_newdf.apply(sku_f,axis=1)
columns=[runout_newdf.columns.values[-1],runout_newdf.columns.values[1]]+list(runout_newdf.columns.values[3:-1])

runout_newdf_update=runout_newdf[columns]

# sort and write actual sales files
runout_newdf_update=runout_newdf_update.sort_values(['Channel','Account','Sku'])
runout_newdf_update.iloc[:,1].fillna(0, inplace=True)



os.chdir("C:/Users/ema/Desktop")
runout_newdf_update.to_excel("Actuals all accounts(except PMD).xlsx",index=False)




