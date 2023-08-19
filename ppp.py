# -*- coding: utf-8 -*-
"""
Created on Mon Aug 23 17:06:57 2021

@author: a.yousefiyan
"""


import glob
import pandas as pd
  

path = "D:/QUERY/PUMP"


 
file_list = glob.glob(path + "/*.xlsx")




excl_list = []

for file in file_list:
    n=pd.read_excel(file, sheet_name = None)
    for key in dict.items(n):
       for m in key:
        if len(m) >= 10:
           excl_list.append(m)
        
                

global b_aa
b_aa = pd.DataFrame()

i=0

while i<=(len(excl_list)-1) :
    
    a=excl_list[i]
    
    b0=a.iloc[2:, 1:4]
    b0 = b0.set_axis(['N', str(a.iloc[3, 2]),'CIP_'+str(a.iloc[3, 2])], axis=1,              inplace=False)
    b0=b0.dropna(subset=['N'])
    b0=b0.reset_index(drop=True)
  
    b1=a.iloc[15:, 4:6]
    b1 = b1.set_axis(['N',  str(a.iloc[3, 2])], axis=1, inplace=False)
    b1=b1.dropna(subset=['N'])
    b1=b1.reset_index(drop=True)
  
    c=pd.concat([b0,b1])
    c=c.reset_index(drop=True)
    c=c.set_index('N')

    
    excl_list[i]=c

    i=i+1


dfs = [df for df in excl_list]
total=pd.concat(dfs, axis=1)
total2=total.reset_index()
total2.to_excel('Pump.xlsx', index=False)