import sys
import pandas as pd
import numpy as np
import xlrd
import math

df2 = pd.read_excel(r"\\Chfile1\ecs_landrive\Automotive_Lighting\LED\Test Engineering\E2E Testing\Profile Program\Fast TShock\Fast TShock (MCA)\Fast_TShock(Ver6.0)-output.xlsx", sheetname = [0,1,2,3,4,5,6,7,8])
df1 = pd.read_excel(r"\\Chfile1\ecs_landrive\Automotive_Lighting\LED\Test Engineering\E2E Testing\Profile Program\Fast TShock\Fast TShock (MCA)\Fast TShock-expected output.xlsx", sheetname = [0,1,2,3,4,5,6,7,8])

GetSheetName = xlrd.open_workbook(r"\\Chfile1\ecs_landrive\Automotive_Lighting\LED\Test Engineering\E2E Testing\Profile Program\Fast TShock\Fast TShock (MCA)\Fast_TShock(Ver6.0)-output.xlsx")
SheetLabel = GetSheetName.sheet_names()

def round_df(df_OrderedDict):
    for i in range(len(df_OrderedDict)):
        df_OrderedDict[i] = df_OrderedDict[i].round(2)
    return df_OrderedDict

def difference(df1, df2):
    difference = []
    for i in range(len(df1)):
        difference_sheet = df1[i]-df2[i][df1[i]!=df2[i]]
        difference.append(difference_sheet)
    return difference

def reset_excel(df):
    df = df.tail(21).set_index('Unnamed: 0')
    df = df[pd.notnull(df.index)]   
    df.columns = ['cold_soak_duration_minute', 'cold_soak_mean_temp_c', 'cold_soak_max_temp_c', 'cold_soak_min_temp_c', 'hot_soak_duration_minute', 'hot_soak_mean_temp_c', 'hot_soak_max_temp_c', 'hot_soak_min_temp_c', 'down_recovery_time_minute', 'down_RAMP_temp_c', 'down_RAMP_rate_c/minute', 'up_recovery_time_minute', 'up_RAMP_temp_c', 'up_RAMP_rate_c/minute']
    df = df.drop(df.index[6])
    return df


def clean_df_ls(df_OrderedDict):
    summary_df = []
    for i in range(len(df_OrderedDict)):
        print(i)
        #if i == 7: import pdb; pdb.set_trace()
        df_OrderedDict[i] = reset_excel(df_OrderedDict[i])
        
        summary_df.append(df_OrderedDict[i])
    return summary_df

def create_wb(test_name):
    writer = pd.ExcelWriter(str(test_name)+'-output.xlsx', engine = 'xlsxwriter')
    return writer    

def write_multiple_dfs(writer, df_list, worksheet_name, spaces, content_instruction):
    row = 2
    for x in list(range(len(df_list))):
        df_list[x].to_excel(writer, sheet_name=worksheet_name, startrow=row , startcol=0)   

        #import pdb; pdb.set_trace()

        worksheet = writer.sheets[worksheet_name]
        row = row - 2
        df_instruction(worksheet, row, content_instruction[x])
        row = row + len(df_list[x].index) + spaces + 9

def df_instruction(worksheet, row, text):
    col = 0
    # Example
    options = {
        'font': {'bold': True, 'color': '#67818a'},
        'border': {'color': 'red', 'width': 3,
                   'dash_type': 'round_dot'},
        'width': 512,
        'height': 30
    }
    worksheet.insert_textbox(row, col, text, options)

def format_excel_file(writer):
    workbook = writer.book
    table_format = workbook.add_format({'align':'center'})
    for sheet_name in writer.sheets:
        if sheet_name != 'Version Info':
            writer.sheets[sheet_name].set_column('A:Z', 27, table_format)



df1_ls = clean_df_ls(df1)
df2_ls = clean_df_ls(df2)
difference = difference(df1, df2)
content_instruction = ["Manual", "Program", "Compare"]

writer = create_wb('tt')
for i in range(len(difference)):
    write_multiple_dfs(writer, [df1_ls[i], df2_ls[i], difference[i]], SheetLabel[i], 4, content_instruction)
info_df = pd.DataFrame({'Version 6.0': []})
info_df.to_excel(writer, sheet_name='Version Info')

### format output excel file
format_excel_file(writer)
writer.save()
