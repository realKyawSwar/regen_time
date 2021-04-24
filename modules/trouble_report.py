import pandas as pd
import numpy as np
import re
from modules import config, mus_regen
# import mus_regen
# import config


view_conn = config.get_connection(config.db_config)
date_1, date_2 = config.get_last_month_dates()


def extract_leak(words):
    """Remove punctuation from list of tokenized words"""
    new_words = []
    number_list = []
    stringy = words.splitlines()
    list_ = ['leak']
    list1_ = ['ok', 'no', 'any']
    stringx = []
    for i, word in enumerate(stringy):
        # remove puntuation and
        new_word = re.sub(r'[^\w\s\'\"]', ' ', word).strip().lower()
        # change
        new_word = re.sub(r'^chg', 'Activities', new_word)
        # remove double space
        new_word = re.sub(re.compile(r' +'), ' ', new_word)
        stringx.append(new_word)
        if any(word in new_word for word in list_) and all(word not in new_word for word in list1_):
            new_words.append(new_word)
            number_list.append(i)
    listy = [stringx[i+1] for i in number_list]
    good_list = np.unique(new_words+listy)
    final_list = []
    for i in good_list:
        if all(word not in i for word in ['internal', 'external']):
            final_list.append(i.capitalize())
    return final_list


def extract(words):
    """Remove punctuation from list of tokenized words"""
    new_words = []
    stringy = words.splitlines()
    list_ = ['pump', 'cylinder', 'tdu', 'robot', 'arm', 'compressor', 'motor']
    list1_ = ['driver', 'alarm', 'rmc motor', 'tdu motor', 'speed', 'ring',
              'mandrel', 'pub', 'pick']
    for word in stringy:
        # remove puntuation and
        new_word = re.sub(r'[^\w\s\'\"]', ' ', word).strip().lower()
        # change
        new_word = re.sub(r'^chg', 'Activities', new_word)
        # remove double space
        new_word = re.sub(re.compile(r' +'), ' ', new_word)
        if new_word.startswith("change") or new_word.startswith("replace"):
            if any(word in new_word for word in list_) and all(word not in new_word for word in list1_):
                # new_words.append(new_word)
                for i in ['change ', ' joshua', ' ngai', ' toh kc', ' lim th',
                          ' koo', ' vincent toh']:
                    new_word = new_word.replace(i, '')
                new_words.append(new_word.capitalize())
    return new_words


def get_report(*args):
    SqlStr = f"SELECT * FROM V_TROUBLE WHERE TROUBLE_DATE BETWEEN"\
             f"'{(args[1])}' AND '{(args[2])}'"\
             "AND PROCESS_DESC= 'SPUTTER' AND TROUBLE_DESC in "\
             f"{(args[3])}"\
             "ORDER BY TROUBLE_DATE"
    # HEADER CALL
    results = args[0].execute(SqlStr)
    df = args[0].result_to_dataframe(results)
    return df


def youtec():
    pd.set_option('display.max_columns', None)
    listy = [f"('{i}')" for i in mus_regen.get_regen_dates()]
    date_string = ','.join(listy)
    # stringy = ''
    stringy = f' AND TROUBLE_DATE in ({date_string})'
    eqt = "('PCVD YOUTEC CHANGE(SCHEDULED)', 'TARGETS CHANGE', 'PCVD YOUTEC CHANGE (TROUBLE)')"
    eqt += stringy
    df = get_report(view_conn, date_1, date_2, eqt)
    # df.to_csv('cvd.csv')
    df = df.drop(columns=['MACHINE_ID', 'PROCESS_ID', 'PROCESS_DESC',
                          'EQUIP_ID', 'COMMENT_TIME', 'EQUIP_DESC', 'ACTION',
                          'UPDATED_DATE', 'UPDATED_TIME', 'ENGINEER_COMMENT',
                          'COMMENT_BY', 'COMMENT_DATE', 'DURATION',
                          'ATTEND_BY', 'UPDATED_BY', 'CAUSES', 'TROUBLE_ID'])
    df['TROUBLE_STIME'] = pd.to_datetime(df['TROUBLE_STIME'])
    df1 = df[(df['TROUBLE_STIME'].dt.time > pd.to_datetime('08:30:00').time()) & (df['TROUBLE_STIME'].dt.time < pd.to_datetime('17:30:00').time())]
    # df1.to_csv('cvd1.csv', index=False)
    df1.rename(columns={'TROUBLE_DATE': 'Date', 'LINE_ID': 'CVD Change'},
               inplace=True)
    df1 = df1.drop(columns=['TROUBLE_DESC', 'TROUBLE_STIME', 'TROUBLE_ETIME'])
    df1['Date'] = pd.to_datetime(df1['Date']).dt.strftime('%d-%b')
    # convert same date line as list to string
    df2 = df1.groupby('Date')['CVD Change'].apply(list).reset_index()
    df2['CVD Change'] = df2['CVD Change'].apply(lambda x: ','.join(str(i) for i in x))
    return df2


def regen():
    eqt = "('REGEN')"
    df = get_report(view_conn, date_1, date_2, eqt)
    df = df.drop(columns=['MACHINE_ID', 'PROCESS_ID', 'PROCESS_DESC',
                          'EQUIP_ID', 'COMMENT_TIME', 'EQUIP_DESC',
                          'UPDATED_DATE', 'UPDATED_TIME', 'ENGINEER_COMMENT',
                          'COMMENT_BY', 'COMMENT_DATE',
                          'ATTEND_BY', 'UPDATED_BY', 'CAUSES', 'TROUBLE_ID'])
    # df.to_csv('regen.csv', index=False)
    df['Activities'] = df.ACTION.apply(lambda x: extract(x))
    df['Leakages'] = df.ACTION.apply(lambda x: extract_leak(x))
    # df = df.drop(df[df.CHANGE.map(len) < 1].index)
    df = df.drop(['DURATION', 'TROUBLE_DESC'], axis=1)
    df['Activities'] = df['Activities'].apply(lambda x: ','.join(map(str, x)))
    df['Leakages'] = df['Leakages'].apply(lambda x: ','.join(map(str, x)))
    df = df[['LINE_ID', 'TROUBLE_DATE', 'TROUBLE_STIME',
             'TROUBLE_ETIME', 'ACTION', 'Activities', 'Leakages']]
    df.rename(columns={'TROUBLE_DATE': 'Date', 'LINE_ID': 'Line'},
              inplace=True)
    df["Line"] = pd.to_numeric(df["Line"], downcast='integer')
    df['Date'] = pd.to_datetime(df['Date'])
    df['TROUBLE_ETIME'] = pd.to_datetime(df['TROUBLE_ETIME'])
    df['TROUBLE_STIME'] = pd.to_datetime(df['TROUBLE_STIME'])
    df['Date'] = df['Date'].dt.strftime('%d-%b')
    df['CHECK'] = pd.to_timedelta(df['TROUBLE_ETIME'].dt.time.astype(str))-pd.to_timedelta('08:30:00')
    df['CHECK'] = df['CHECK'] / pd.to_timedelta(1, unit='m')
    df = df[['Line', 'Date', 'TROUBLE_STIME', 'TROUBLE_ETIME',
             'CHECK', 'ACTION', 'Activities', 'Leakages']]
    # df.ACTION.apply(lambda x: x.splitlines())
    # df = df.assign(ACTION=df['ACTION'].str.split('\n')).explode('ACTION')
    # df = df.replace('\r', '', regex=True)
    # df = df[df['ACTION'].str.strip().astype(bool)]
    # df.set_index(['TROUBLE_DATE', 'LINE_ID', 'TROUBLE_STIME',
    #               'TROUBLE_ETIME'], inplace=True)
    # # df.set_index(df.groupby(level=[0, 1, 2, 3]).cumcount(), append=True)
    # # .sort_index()
    # # pd.set_option('display.max_columns', None)
    # # df.to_excel("multi.xlsx", sheet_name='Regen Report')
    # # print(df.head(5))
    # # df.to_csv('trouble_report.csv', index=True)
    return df


# print(regen())
# youtec()
