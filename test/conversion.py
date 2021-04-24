import numpy as np
import config
# import database
import pandas as pd
view_conn = config.get_connection(config.db_config)
date_1, date_2 = config.get_last_month_dates()


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


def get_musconvo(*args):
    # SQL STR
    SqlStr = f"SELECT * FROM V_MU_RESULT WHERE PROD_DATE BETWEEN "\
             f"'{(args[1])}' AND '{(args[2])}'"\
             "AND ERROR_DESC = 'REGEN PM'"\
             "ORDER BY PROD_DATE"

    # HEADER CALL
    results = args[0].execute(SqlStr)
    df = args[0].result_to_dataframe(results)
    return df


def add_dayfornight(x):
    start = pd.to_datetime('20:00:00').time()
    end = pd.to_datetime('8:30:00').time()
    if (x['TROUBLE_STIME'] > start) and (x['TROUBLE_ETIME'] < end):
        last_row = x.iloc[-1]
        print(last_row)


def report_convo():
    eqt = "('CONVERSION/PLAN DOWN')"
    df = get_report(view_conn, date_1, date_2, eqt)
    df.to_excel('mar.xlsx')
    df1 = df.drop(columns=['MACHINE_ID', 'PROCESS_ID', 'PROCESS_DESC',
                           'EQUIP_ID', 'COMMENT_TIME', 'EQUIP_DESC',
                           'UPDATED_DATE', 'UPDATED_TIME', 'ENGINEER_COMMENT',
                           'COMMENT_BY', 'COMMENT_DATE', 'DURATION',
                           'TROUBLE_DESC', 'ATTEND_BY', 'UPDATED_BY',
                           'CAUSES', 'TROUBLE_ID', 'ACTION'])
    df1['TROUBLE_ETIME'] = pd.to_datetime(df1['TROUBLE_ETIME'],
                                          format='%H:%M').dt.time
    df1['TROUBLE_STIME'] = pd.to_datetime(df1['TROUBLE_STIME'],
                                          format='%H:%M').dt.time
    df1.loc[df1['TROUBLE_ETIME'] > pd.to_datetime('19:59:59').time(),
            'TROUBLE_ETIME'] = pd.to_datetime('23:59:59').time()
    s = df1.groupby('LINE_ID').head(1).index.to_list()
    df1.loc[s, 'TROUBLE_STIME'] = pd.to_datetime('08:30:00').time()
    end = df1.groupby('LINE_ID').tail(1).index.to_list()
    mid = [x for x in df1.index.to_list() if x not in s+end]
    df1.loc[mid, 'TROUBLE_STIME'] = pd.to_datetime('00:00:00').time()
    df1 = df1.drop_duplicates(subset=['LINE_ID', 'TROUBLE_DATE'],
                              keep='last').reset_index(drop=True)
    df1['TROUBLE_DATE'] = pd.to_datetime(df1['TROUBLE_DATE'])
    # print(df1)
    maxd = df1.loc[df1.groupby('LINE_ID')['TROUBLE_DATE'].idxmax()]
    start = pd.to_datetime('20:00:00').time()
    end = pd.to_datetime('08:30:00').time()
    idx = np.where((maxd['TROUBLE_STIME'] > start) & (maxd['TROUBLE_ETIME'] < end))
    maxd = maxd.iloc[idx]
    maxd['TROUBLE_DATE'] = maxd['TROUBLE_DATE'] + pd.DateOffset(days=1)
    maxd['TROUBLE_STIME'] = pd.to_datetime('00:00:00').time()
    df1 = pd.concat([df1,
                     maxd]).sort_values(['LINE_ID',
                                        'TROUBLE_DATE']).reset_index(drop=True)
    # change value of night time
    df1.loc[(df1['TROUBLE_STIME'] > start) & (df1['TROUBLE_ETIME'] < end),
            ['TROUBLE_STIME', 'TROUBLE_ETIME']] = [pd.to_datetime('00:00:00').time(),
                                                   pd.to_datetime('23:59:59').time()]
    df1['Duration'] = (pd.to_timedelta(df1['TROUBLE_ETIME'].astype(str)) -
                       pd.to_timedelta(df1['TROUBLE_STIME'].astype(str))
                       ).dt.total_seconds()/3600
    df2 = df1.groupby(['LINE_ID'],
                      as_index=False).agg({'Duration': 'sum'})
    df3 = df1.groupby(['LINE_ID'], as_index=False).agg(['first', 'last']).stack().reset_index()
    df3 = df3.drop(['TROUBLE_ETIME', 'TROUBLE_STIME', 'Duration'], axis=1)
    df4 = df3.pivot(index='LINE_ID', columns='level_1', values='TROUBLE_DATE').reset_index()
    finaldf = pd.merge(df2, df4, how='left', on='LINE_ID')
    finaldf.rename(columns={'LINE_ID': 'Line', 'first': 'Start', 'last': 'End'}, inplace=True)
    finaldf.Duration = finaldf.Duration.round(decimals=2)
    finaldf = finaldf[['Line', 'Start', 'End', 'Duration']].sort_values(['Line','Start']).reset_index(drop=True)
    print(finaldf)


report_convo()
# df = get_musconvo(view_conn, date_1, date_2)
# df.to_excel("musconvo1.xlsx")
# print('done')


# df.ACTION.apply(lambda x: x.splitlines())
# df = df.assign(ACTION=df['ACTION'].str.split('\n')).explode('ACTION')
# df = df.replace('\r', '', regex=True)
# df = df[df['ACTION'].str.strip().astype(bool)]
# multi = df.set_index(['TROUBLE_DATE', 'LINE_ID', 'TROUBLE_STIME',
#                      'TROUBLE_ETIME']).sort_index()
# print(df.head(5))
# df.to_excel('conversionfeb.xlsx')
# lol = df[df['LINE_ID'].unique()]

# df = df.groupby('TROUBLE_DATE').last().reset_index()
# df1['TROUBLE_STIME'] = '08:30'
# lol = df1.groupby('LINE_ID')['TROUBLE_DATE'].count()
# df1['TROUBLE_DATE'] = df1.groupby('LINE_ID')['TROUBLE_DATE'].transform(min)

# df2 = df1[df1['TROUBLE_ETIME'] > pd.to_datetime('20:00:00').time()]
# df2['TROUBLE_ETIME'] = pd.to_datetime('23:59:59').time()
# df1.update(df2)

# posy = (pd.to_timedelta(df1['TROUBLE_ETIME'].astype(str)) - pd.to_timedelta(df1['TROUBLE_STIME'].astype(str))).dt.total_seconds()/3600
# negy = pd.to_timedelta(df1['TROUBLE_ETIME'].astype(str))
# negy1 = negy + (pd.to_timedelta('23:59:59') - pd.to_timedelta(df['TROUBLE_STIME'].astype(str)))
# con1 = df1['TROUBLE_ETIME'] > df1['TROUBLE_STIME']
# con2 = df1['TROUBLE_STIME'] > df1['TROUBLE_ETIME']
# df1['Duration'] = np.select([con1, con2],
#                             [posy, negy1.dt.total_seconds()/3600])
