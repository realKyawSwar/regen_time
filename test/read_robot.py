import pandas as pd
import numpy as np

# from multiprocessing import Pool
# import numba


path = 'R:/MAINT/SDE/Sputter/Sputter Group/Elgin/Equipment PM.xlsx'
# gvpath = 'R:/MAINT/SDE/Sputter/Sputter Group/Sputter Maint/Li Dongneng/'\
#          'Sputter/Dong/Gate Cylinder/GV, VAC & '\
#          'ATM Cylinder Position (Latest).xlsx'
# cryopath = 'R:/MAINT/SDE/Sputter/Sputter Group/Elgin/Cryopump PM2.xlsx'
xls = pd.ExcelFile(path)


def correction(original_function):
    def wrapper_function(*args, **kwargs):
        df = original_function(*args, **kwargs)
        df = df[['Serial Number', 'line', 'chamber', 'install_date',
                 'due_date', 'Type']]
        df.Type = df.Type.astype(np.int64)
        df['install_date'] = pd.to_datetime(df['install_date'])
        df['due_date'] = pd.to_datetime(df['due_date'])
        df['install_date'] = df['install_date'].dt.strftime('%Y-%m-%d')
        df['due_date'] = df['due_date'].dt.strftime('%Y-%m-%d')
        return df
    return wrapper_function


def drypump(xls):
    df = pd.read_excel(xls, sheet_name='Dry pump',
                       usecols=["S/N", "Line", "Last O/H", "Next O/H",
                                "Type"])
    df = df.rename(columns={"Last O/H": 'Date', "Next O/H": "Due Date",
                            "S/N": "Serial Number"})
    df['Location'] = 'All'
    df = df[(df['Type'] == 'MU600') | (df['Type'] == 'NeoDry E30')]
    df = df[['Line', 'Location', 'Type', 'Date', 'Due Date', 'Serial Number']]
    df = df[pd.to_numeric(df['Line'], errors='coerce').notnull()]
    return df


def robot(xls):

    def tidyUp(things):
        a = things.replace("ULD ", "").replace("LD ", "")
        # b = a.replace("Xfer Arm", "Robot Arm").replace(
        #                   "PP Arm", "Robot Arm").replace(
        #                   "Epson ", "Epson Robot")
        # c = b.replace("PP", "Brooks Robot").replace(
        #                   "Xfer", "Brooks Robot")
        return a

    def get_robot_id(x):
        if x == 'Xfer':
            y = 9
        elif x == 'PP':
            y = 8
        elif x == 'Xfer Arm':
            y = 11
        elif x == 'PP Arm':
            y = 10
        else:
            y = x
        return y

    df = pd.read_excel(xls, sheet_name='Robots',
                       usecols=["Unnamed: 0", "Unnamed: 1",
                                "Unnamed: 2", "Unnamed: 3",
                                "Unnamed: 6"])[7:]
    df = df.rename(columns={"Unnamed: 0": 'line', "Unnamed: 1": 'install_date',
                            "Unnamed: 2": 'due_date',
                            "Unnamed: 3": 'Type',
                            "Unnamed: 6": 'Serial Number'})
    df['chamber'] = np.where(df['Type'].str.contains(r'U(?!$)'),
                             'UNLOAD', 'LOAD')
    df = df.dropna()
    df.Type = df.Type.astype(str).apply(lambda x: tidyUp(x))
    df.Type = df.Type.apply(get_robot_id)
    # df.Type = df.Type.apply(lambda x: tidyUp(x))
    df = df[['Serial Number', 'line', 'chamber', 'install_date',
             'due_date', 'Type']]
    df.loc[df['Serial Number'].str.startswith('G', na=False), 'Type'] = 16
    df.loc[df['Type'].str.contains('Epson', na=False), 'Type'] = 15
    df.Type = df.Type.astype(np.int64)
    df['install_date'] = pd.to_datetime(df['install_date'])
    df['due_date'] = pd.to_datetime(df['due_date'])
    df['install_date'] = df['install_date'].dt.strftime('%Y-%m-%d')
    df['due_date'] = df['due_date'].dt.strftime('%Y-%m-%d')
    return df


@correction
def tdu(xls):

    def tidyUp(things):
        a = things.replace("ULD", "UNLOAD").replace("LD", "LOAD")
        b = a.replace("ULC", "UNLOAD").replace("LC", "LOAD")
        return b

    df = pd.read_excel(xls, sheet_name='TDU',
                       usecols=['Line', 'Chamber', 'Change Date', 'Due Date'])
    df = df.rename(columns={'Change Date': 'install_date',
                            'Chamber': 'chamber',
                            'Line': 'line',
                            'Due Date': 'due_date'})
    df['Type'] = 28
    df.chamber = df.chamber.astype(str).apply(lambda x: tidyUp(x))
    # df['Type'] = np.where(df['Location'].str.contains(r'P(?!$)'),
    #                       'Process ', 'Corner ') + df.Type
    df["Serial Number"] = df.apply(lambda x: 'TDU'+str(x.line)+str(x.chamber),
                                   axis=1)
    return df


@correction
def carrier_force(xls):
    df = pd.read_excel(xls, sheet_name='Rail Alignment',
                       usecols=['Line', 'Carrier Force'])[:11]
    df = df.rename(columns={'Carrier Force': 'install_date',
                            'Line': 'line'})
    df['chamber'] = 'ALL'
    df['Type'] = 27
    df["Serial Number"] = df.apply(lambda x: 'CFORCE'+str(x.line)+str(x.chamber),
                                   axis=1)
    df['install_date'] = pd.to_datetime(df['install_date'])
    df['due_date'] = df['install_date'] + pd.DateOffset(days=90)
    return df


@correction
def orifice(xls):
    df = pd.read_excel(xls, sheet_name='Mbox,orifice',
                       usecols=['Line', 'Equipment', 'Last service',
                                'Next service'])[:25]
    df = df.rename(columns={'Equipment': 'Type',
                            'Last service': 'install_date',
                            'Line': 'line',
                            'Next service': 'due_date'})
    df['chamber'] = 'ALL'
    df["Serial Number"] = df.apply(lambda x: str(x.Type).upper()[:6]+str(x.line)+str(x.chamber),
                                   axis=1)
    df.loc[df['Type'].str.contains('Orifice', na=False), 'Type'] = 4
    df.loc[df['Type'].str.contains('Linear', na=False), 'Type'] = 26
    # df = df[['Line', 'Location', 'Type', 'Date', 'Due Date', 'Serial Number']]
    df = df.dropna()
    return df[~df.Type.str.contains("VHF", na=False)]


@correction
def carrier_lock(xls):
    df = pd.read_excel(xls, sheet_name='carrier lock asembly',
                       usecols=['Line', 'LOAD', 'UNLOAD', 'ROT'])[:11]
    df = df.rename(columns={'ROT': 'P9'})
    df = df.melt(id_vars=["Line"], var_name="Location", value_name="Date")
    df['Type'] = 24
    df = df.drop(df[df.Date == '-'].index)
    df['Date'] = pd.to_datetime(df['Date'])
    df['Due Date'] = df['Date'] + pd.DateOffset(days=365)
    df = df.rename(columns={'Line': 'line', 'Location': 'chamber',
                            'Date': 'install_date', 'Due Date': 'due_date'})
    df["Serial Number"] = df.apply(lambda x: 'CLOCK'+str(x.line)+str(x.chamber),
                                   axis=1)
    # df = df[['Line', 'Location', 'Type', 'Date', 'Due Date', 'Serial Number']]
    return df


@correction
def vacuum_motor(xls):
    df = pd.read_excel(xls, sheet_name='Vacuum motor',
                       usecols=['Line', 'LOAD', 'UNLOAD', 'ROT'])[:11]
    df = df.rename(columns={'ROT': 'P9'})
    df = df.melt(id_vars=["Line"], var_name="Location", value_name="Date")
    df = df.drop(df[df.Date == '-'].index)
    df['Date'] = pd.to_datetime(df['Date'])
    df['Type'] = 12
    df['Due Date'] = df['Date'] + pd.DateOffset(days=545)
    df = df.rename(columns={'Line': 'line', 'Location': 'chamber',
                            'Date': 'install_date', 'Due Date': 'due_date'})
    df["Serial Number"] = df.apply(lambda x: 'VMOTOR'+str(x.line)+str(x.chamber),
                                   axis=1)
    # df["Serial Number"] = False
    # df = df[['Line', 'Location', 'Type', 'Date', 'Due Date', 'Serial Number']]
    return df


@correction
def tdu_greasing(xls):
    df = pd.read_excel(xls, sheet_name="TDU greasing",
                       usecols=['Line', 'Last service', 'Due Date'])
    df.columns = ['Line', 'Date', 'Due Date']
    df['Location'] = 'ALL'
    df['Type'] = 23
    df = df.rename(columns={'Line': 'line', 'Location': 'chamber',
                            'Date': 'install_date', 'Due Date': 'due_date'})
    df["Serial Number"] = df.apply(lambda x: 'TDUGREASE'+str(x.line)+str(x.chamber),
                                   axis=1)
    return df


def bias(xls):
    df = pd.read_excel(xls, sheet_name='Bias',
                       usecols=['Line', 'Chamber', 'Chg Date', 'Due Date',
                                'Type'])
    df = df.rename(columns={'Chg Date': 'Date', 'Chamber': 'Location'})
    df['Date'] = pd.to_datetime(df['Date'])
    df["Serial Number"] = False
    df = df[['Line', 'Location', 'Type', 'Date', 'Due Date', 'Serial Number']]
    df.Type = df.Type.apply(lambda x: f"{x} Bias")
    return df


# df = pd.concat([vacuum_motor(xls), tdu_greasing(xls), orifice(xls),
#                 carrier_lock(xls), carrier_force(xls), tdu(xls)])
# df = compressor(xls)
df = df.reset_index(drop=True)
df.to_excel("premisc.xlsx", index=False)
configdf = pd.read_excel("chamber_view.xlsx")
configdf.chamber = configdf.chamber.str.strip()
result = pd.merge(df, configdf, how="left", on=['line', 'chamber'])
lol = result[['Serial Number', 'id_', 'install_date', 'due_date', 'Type']]
lol.to_excel("misc.xlsx", index=False)
