import pandas as pd
# import numpy as np
# from time import perf_counter
# import openpyxl
import datetime

# print(delta.days)
# print(today)
# # path = r"R:/MAINT/PUBLIC/Public SDE/SDE Regen Status/SDE- Regen & Carrier Change Schedule.xlsm"
path = r"SDE- Regen & Carrier Change Schedule.xlsm"


def ColIdxToXlName(idx):
    # if idx < 1:
    #     raise ValueError("Index is too small")
    result = ""
    while True:
        if idx > 26:
            idx, r = divmod(idx - 1, 26)
            result = chr(r + ord('A')) + result
        else:
            return chr(idx + ord('A') - 1) + result

# regen_path = r"R:/MAINT/PUBLIC/Public SDE/SDE Regen Status/SDE- Regen & Carrier Change Schedule.xlsm"
# rgx = pd.ExcelFile(path)


def get_regen():
    today = datetime.date.today()
    start_date = datetime.date(2021, 4, 25)
    delta = today - start_date
    start_no = 863 + delta.days
    end_no = start_no + 25
    # cell column "AGE" on 25/4/2021
    start = ColIdxToXlName(start_no)
    end = ColIdxToXlName(end_no)
    df = pd.read_excel(path, sheet_name=0,
                       index_col=None, na_values=['NA'],
                       usecols=f"A, {start}:{end}")[2:47]
    df = df.dropna(how='all')
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    df.columns = df.columns.strftime('%Y-%m-%d')
    df = df.rename(columns={'NaT': 'line'})
    df = df.dropna(subset=['line'])
    df = df.set_index('line')
    df1 = df.loc[:, (df == 'REGEN(AM)').any()]
    df1 = df1.reset_index()
    df1 = pd.melt(df1, id_vars=['line'], var_name='date')
    df1 = df1[df1['value'] == 'REGEN(AM)']
    print(df1)


get_regen()
# df.to_excel("regen_schedule.xlsx", index=False)
