# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np
import os

os.chdir("O:/Supply Chain/Demand Planning/Forecast Master")

book2_forecast=pd.read_excel("November SOP Roll Up Master With Notes - Book 2.xlsx", sheet_name='Sheet1',header=0)

def read_forecast(df,col,name):
    df_piece=df[df[col]==name]
    return df_piece

ulta_forecast=read_forecast(book2_forecast,"Account","Ulta")
sephora_forecast=read_forecast(book2_forecast,"Account","Sephora")
othercom_forecast=read_forecast(book2_forecast,"Account","Other.com")
ec_scott_forecast=read_forecast(book2_forecast,"Account","EC Scott")
dillards_forecast=read_forecast(book2_forecast,"Account","Dillard's")
nm_forecast=read_forecast(book2_forecast,"Account","Neiman Marcus")
