import pandas as pd
from modules import config
pd.options.mode.chained_assignment = None


def get_regen(*args):
    # SQL STR
    SqlStr = f"SELECT * FROM V_MU_RESULT WHERE PROD_DATE BETWEEN "\
             f"'{(args[1])}' AND '{(args[2])}'"\
             "AND ERROR_DESC = 'REGEN PM'"\
             "ORDER BY PROD_DATE"
    # HEADER CALL
    results = args[0].execute(SqlStr)
    df = args[0].result_to_dataframe(results)
    return df


# if __name__ == '__main__':
def regen_time():
    try:
        view_conn = None
        view_conn = config.get_connection(config.db_config)
        date_1, date_2 = config.get_last_month_dates()
        # pd.set_option('display.max_columns', None)
        df = get_regen(view_conn, date_1, date_2)
        df = df.drop(['DEPT_ID', 'STATION', 'SPECODE', 'MEDIA_SIZE',
                      'THICKNESS_TYPE', 'ID', 'OPR_TIME',
                      'CODE', 'ERROR_DESC', 'PROD_COMMENT',
                      'ERR_TYPE_ID', 'ERR_TYPE_DESC', 'ERR_ID', 'ERR_DESC',
                      'ISSUE', 'CHAMBER_NAME', 'UPDATED_DATE', 'UPDATED_BY',
                      'YEAR', 'MONTH', 'MONTHLY_OPR_TIME', 'MONTHLY_DOWNTIME',
                      'MONTHLY_DOWNTIME_PCT', 'MONTHLY_DOWNTIME_MU_PCT'],
                     axis=1)
        # header = df.columns.values.tolist()
        df['PROD_DATE'] = pd.to_datetime(df['PROD_DATE'])
        df['STOP_TIME'] = pd.to_datetime(df.groupby(['PROD_DATE',
                                                     'LINE_ID'])['STOP_TIME'].transform(min))
        df['START_TIME'] = pd.to_datetime(df.groupby(['PROD_DATE',
                                                      'LINE_ID'])['START_TIME'].transform(max))
        df1 = df.groupby(['PROD_DATE', 'LINE_ID', 'PROGRAM',
                          'DISK_TYPE', 'STOP_TIME', 'START_TIME'],
                         as_index=False).agg({'DURATION': 'sum'})
        df2 = df1[df1['STOP_TIME'].dt.time < pd.to_datetime('08:30:00').time()]
        df2['CHECK'] = pd.to_timedelta('08:30:00') - pd.to_timedelta(df2['STOP_TIME'].dt.time.astype(str))
        df2['CHECK'] = df2['CHECK'] / pd.to_timedelta(1, unit='m')
        df2['DURATION'] = df2['DURATION'] - df2['CHECK']
        df2 = df2.drop(['CHECK'], axis=1)
        df1.update(df2)
        df1.DURATION = df1.DURATION.round(decimals=2)
        df1 = df1[(df1['DURATION'] > 0)]
        df1['HOURS'] = df1.DURATION.apply(lambda x: (x/60)).round(decimals=2)
        df1.rename(columns={'PROD_DATE': 'Date', 'LINE_ID': 'Line',
                            'PROGRAM': 'Program', 'DISK_TYPE': 'Disk Type'},
                   inplace=True)
        df1["Line"] = pd.to_numeric(df1["Line"], downcast='integer')
        df1['Date'] = df1['Date'].dt.strftime('%d-%b')
        # raw = df1
        # df1 = df1.drop(['START_TIME', 'STOP_TIME'], axis=1)
        df1["Target"] = 9
        # clean = df1[['Date', 'Line', 'Hours', 'Target', 'Program', 'Disk Type',
        #              'DURATION']]
        # df1.to_csv('MUS.csv', index=False)
        # print(df1)
    #     df1.to_excel("clean1.xlsx",  index=False, sheet_name='Regen Report')
    except Exception as e:
        raise e
    finally:
        if view_conn is not None:
            view_conn.close()
        return df1


def get_regen_dates():
    view_conn = None
    view_conn = config.get_connection(config.db_config)
    date_1, date_2 = config.get_last_month_dates()
    # pd.set_option('display.max_columns', None)
    df = get_regen(view_conn, date_1, date_2)
    listy = df.PROD_DATE.unique()
    return listy


# regen_raw, regen_clean = regen_time()

# def get_convo():
#     view_conn = None
#     view_conn = config.get_connection(config.db_config)
#     date_1, date_2 = config.get_last_month_dates()
#     # SQL STR
#     SqlStr = f"SELECT * FROM V_MU_RESULT WHERE PROD_DATE BETWEEN "\
#              f"'{(args[1])}' AND '{(args[2])}'"\
#              # "AND ERROR_DESC = 'REGEN PM'"\
#              # "ORDER BY PROD_DATE"
#     # HEADER CALL
#     results = args[0].execute(SqlStr)
#     df = args[0].result_to_dataframe(results)
#     df.to_excel("testy.xlsx")
#     return df

# # print(regen_time().dtypes)
# # regen_time()
# get_convo()
