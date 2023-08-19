# -*- coding: utf-8 -*-
"""
Created on Mon Aug 23 17:06:57 2021

@author: a.yousefiyan
"""

import glob
import pandas as pd
  

path = "D:/QUERY/Air Heater"

file_list = glob.glob(path + "/*.xlsx")

excl_list = []
excl_list2 = []
for file in file_list:
    n=pd.read_excel(file, sheet_name = None)
    for key in dict.items(n):
       for m in key:
        if len(m) >= 10 and len(m) <= 32: #for 2 section Air heater
           excl_list.append(m)
        if len(m) >= 10 and len(m) > 32: #for 2 section Air heater
           excl_list2.append(m)

b_aa = pd.DataFrame()

i=0

while i<=(len(excl_list)-1) :
    
    a=excl_list[i]
    b0=a.iloc[2:, 1:6]

    b0 = b0.set_axis(['N',"B","P_S_"+ str(a.iloc[3, 2]),"D",'U_S_'+str(a.iloc[3, 2])], axis=1,              inplace=False)
    b0=b0.drop(['B', 'D'], axis=1)
    b0=b0.dropna(subset=['N'])

    b0=b0.reset_index(drop=True)
    b0=b0.set_index('N')

    """
    b1=a.iloc[2:, 4:6]
    b1 = b1.set_axis(['N',  str(a.iloc[3, 2])], axis=1, inplace=False)
    b1=b1.dropna(subset=['N'])
    b1=b1.reset_index(drop=True)
  
    c=pd.concat([b0,b1])
    c=c.reset_index(drop=True)
    c=c.set_index('N')
    """
    
    excl_list[i]=b0
    i=i+1



dfs = [df for df in excl_list]
total=pd.concat(dfs, axis=1)
total2=total.reset_index()
total2.to_excel('Air Heater.xlsx', index=False)