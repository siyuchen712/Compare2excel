#!/usr/bin/python3
# -*- coding: utf-8 -*-

import sys
import pandas as pd

df1 = pd.read_excel(r"C:\Users\s.chen6\Desktop\compare_data\manual.xlsx", sheetname = [0,1,2,3,4,5,6,7,8])
df2 = pd.read_excel(r"C:\Users\s.chen6\Desktop\compare_data\Program.xlsx", sheetname = [0,1,2,3,4,5,6,7,8])

#difference = df1-df2[df1!=df2]

def round_df(df_OrderedDict):
    for i in range(len(df_OrderedDict)):
        df_OrderedDict[i] = df_OrderedDict[i].round(2)
    return df_OrderedDict

def difference(df1, df2):
    difference = []
    for i in range(len(df1)):
        difference_sheet = df1[i]-df2[i][df1[i]!=df2[i]]
        difference_sheet.to_excel(writer,'Sheet'+str(i+1))
        difference.append(difference_sheet)
    return difference

df1 = round_df(df1)
df2 = round_df(df2)

writer = pd.ExcelWriter('Fast_TShock_ra_compare_output.xlsx')
df_difference = difference(df1, df2)
writer.save()
