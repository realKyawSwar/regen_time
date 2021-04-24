import pandas as pd
import numpy as np

gvpath = 'R:/MAINT/SDE/Sputter/Sputter Group/Sputter Maint/Li Dongneng/'\
         'Sputter/Dong/Gate Cylinder/GV, VAC & '\
         'ATM Cylinder Position (Latest).xlsx'

gvx = pd.ExcelFile(gvpath)


def get_gv():
    line_list = [201, 202, 203, 204, 211, 208, 209, 210, 205]
    combined = pd.DataFrame(columns=['line', 'chamber', 'Type', 'Date',
                                     'Due Date', 'Serial Number'])
    for i in [gv(j) for j in line_list]:
        combined = pd.concat([combined, i])
    return combined


def gv(line):
    df = pd.read_excel(gvx, sheet_name=f'{line}',
                       usecols=["Unnamed: 2", "Unnamed: 3",
                                "Unnamed: 4", "Unnamed: 6"])[52:101]
    df = df.rename(columns={"Unnamed: 2": 'Due Date', "Unnamed: 3": 'Date',
                            "Unnamed: 4": 'chamber',
                            "Unnamed: 6": 'Serial Number'})
    df = df.dropna()
    df['line'] = line
    df = df[pd.to_datetime(df['Date'], errors='coerce').notnull()]
    # if line != 205:
    df['Type'] = np.where(df['Serial Number'].str.contains(r'K(?!$)'),
                          5, '')
    # else:
    #     df['Type'] = 'PGV'
    df.at[df.index[-1], 'Type'] = 7
    df.at[df.index[-2], 'Type'] = 7
    df.at[df.index[-3], 'Type'] = 6
    df.at[df.index[-4], 'Type'] = 6
    df1 = df[['Serial Number', 'line', 'chamber', 'Date', 'Due Date', 'Type']]
    # df.chamber = df.chamber.apply(lambda z: z.replace("ULC", "Unload")
    #                                 .replace("LC", "Load").replace(" Vac", "")
    #                                 .replace(" Atm", ""))
    df1.chamber = df1.chamber.apply(lambda z: z.replace("ULC", "UNLOAD")
                                    .replace("LC", "LOAD")
                                    .replace("LC", "Load").replace(" Vac", "")
                                    .replace(" Atm", ""))
    df1.Type = df1.Type.astype(np.int64)
    df1['Date'] = pd.to_datetime(df1['Date'])
    df1['Due Date'] = pd.to_datetime(df1['Due Date'])
    df1['Date'] = df1['Date'].dt.strftime('%Y-%m-%d')
    df1['Due Date'] = df1['Due Date'].dt.strftime('%Y-%m-%d')
    # print(df[-4:])
    return df1


df = get_gv()
configdf = pd.read_excel("chamber_view.xlsx")
configdf.chamber = configdf.chamber.str.strip()
# configdf.id_ = configdf.id_.astype(np.int64)
# print(type(configdf.chamber.iloc[1]))
# print(configdf.dtypes)
result = pd.merge(df, configdf, how="left", on=['line', 'chamber'])
lol = result[['Serial Number', 'id_', 'Date', 'Due Date', 'Type']]
# print(lol.head(5))
lol.to_excel("gv1.xlsx", index=False)
