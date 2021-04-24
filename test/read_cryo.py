import pandas as pd
import numpy as np

cryopath = 'R:/MAINT/SDE/Sputter/Sputter Group/Elgin/Cryopump PM2.xlsx'
crx = pd.ExcelFile(cryopath)


def clean(x):
    if x == 10:
        y = 1
    elif x == '10E':
        y = 3
    else:
        y = 2
    return y


def cryopump(crx):
    df = pd.read_excel(crx, sheet_name='Eq summary',
                       usecols=["Unnamed: 0", "Unnamed: 1",
                                "Unnamed: 9", "Unnamed: 3", "Unnamed: 4",
                                "Unnamed: 5"])[34:400]
    # print(df)
    df = df.rename(columns={"Unnamed: 0": 'Line', "Unnamed: 1": 'Location',
                            "Unnamed: 9": 'Type',
                            "Unnamed: 3": 'Serial Number',
                            "Unnamed: 4": 'Date', "Unnamed: 5": 'Due Date'})
    # df = df[(df['Type'] == 'Cryopump')]
    df = df.dropna()
    df = df[pd.to_datetime(df['Date'], errors='coerce').notnull()]
    df = df[pd.to_numeric(df['Line'], errors='coerce').notnull()]
    df = df[['Line', 'Location', 'Type', 'Date', 'Due Date', 'Serial Number']]
    df1 = df[['Line', 'Location', 'Serial Number',  'Date', 'Due Date', 'Type']]
    df1['Date'] = pd.to_datetime(df1['Date'])
    df1['Due Date'] = pd.to_datetime(df1['Due Date'])
    df1['Date'] = df1['Date'].dt.strftime('%Y-%m-%d')
    df1['Due Date'] = df1['Due Date'].dt.strftime('%Y-%m-%d')
    df1['Line'] = df1['Line'].astype(np.int64)
    df1['Location'] = df1['Location'].astype(str)
    df1.rename(columns={'Line': 'line', 'Location': 'chamber',
                        'Type': 'type'},
               inplace=True)
    # print(df1.iloc[8])
    df1.type = df1.type.apply(clean)
    # print(df1.head(10))
    return df1


df = cryopump(crx)
# print(type(df.chamber.iloc[1]))
# print(df.dtypes)
configdf = pd.read_excel("chamber_view.xlsx")
configdf.chamber = configdf.chamber.str.strip()
# configdf.id_ = configdf.id_.astype(np.int64)
# print(type(configdf.chamber.iloc[1]))
# print(configdf.dtypes)
result = pd.merge(df, configdf, how="left", on=['line', 'chamber'])
result.to_excel("cryopump.xlsx")
