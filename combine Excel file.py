# -*- coding: utf-8 -*-
"""
Created on Mon Aug 23 17:06:57 2021

@author: a.yousefiyan
"""

# importing the required modules
import glob
import pandas as pd
  
# specifying the path to csv files
#path = "C:/downloads"
path = "D:/My file/SPRAY DRYER/new file"


 
# csv files in the path
file_list = glob.glob(path + "/*.xlsx")



#xls = xlrd.open_workbook(r'<path_to_your_excel_file>', on_demand=True)
# list of excel files we want to merge.
# pd.read_excel(file_path) reads the excel
# data into pandas dataframe.
excl_list = []

for file in file_list:
    n=pd.read_excel(file, sheet_name = None)
    for key in dict.items(n):
       for m in key:
        if len(m) >= 10:
           excl_list.append(m)
        
                
  
   # b=(pd.read_excel(file)).iloc[3, 2]
    #df['Name'] = 'abc'
# create a new dataframe to store the 
# merged excel file.






#excl_merged = pd.DataFrame()
global b_aa
b_aa = pd.DataFrame()

i=0

while i<=(len(excl_list)-1) :
    
    a=excl_list[i]
    b0=a.iloc[2:27, 1:4]
    #b0 = b0.rename(columns={'oldName1': 'newName1', 'oldName2': 'newName2'})
    #b0.columns =['N', 'Value','CIP VALUE']
    b0 = b0.set_axis(['N', str(a.iloc[3, 2]),'CIP_'+str(a.iloc[3, 2])], axis=1,              inplace=False)
    b0=b0.reset_index(drop=True)
  


    b1=a.iloc[15:27, 4:6]
    b1 = b1.set_axis(['N',  str(a.iloc[3, 2])], axis=1, inplace=False)
    b1=b1.reset_index(drop=True)
  
    c=pd.concat([b0,b1])
    c=c.reset_index(drop=True)
    c=c.set_index('N')
    #result = pd.concat([df1, df4], axis=1)
    #c["TAG"]=a.iloc[3, 2]
    
    excl_list[i]=c
    #b_aa = b_aa.append(c,ignore_index=False)
    i=i+1


dfs = [df for df in excl_list]
total=pd.concat(dfs, axis=1)
#for excl_file in excl_list:
    
    # appends the data into the excl_merged 
    # dataframe.
#    excl_merged = excl_merged.append( excl_file, ignore_index=True)
  
# exports the dataframe into excel file with
# specified name.
total.to_excel('Combine.xlsx', index=False)