# -*- coding: utf-8 -*-
"""
Created on Tue Nov 26 16:59:31 2019

@author: ema
"""

import pandas as pd
import numpy as np
import os

os.chdir("C:/Users/ema/Desktop")

sop_channel=pd.read_excel("December SOP Roll Up Master Final.xlsx", header=0)
cols=np.append(sop_channel.columns.values[3],sop_channel.columns.values[11:])
sop_company=sop_channel[cols].groupby("Sku", as_index=False).sum()
sop_company.to_excel("Company Forecast.xlsx", index=False)