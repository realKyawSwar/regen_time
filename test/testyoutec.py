import pandas as pd
import mus_regen


df = pd.read_csv('cvd.csv')
listy = mus_regen.get_regen_dates()
lol = [','.join(ele.split()) for ele in listy]


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

dfs = pd.DataFrame({'Date': lol})
df5 = df1[df1['Date'].isin(dfs['Date'])]

df5 = df5.drop(columns=['Unnamed: 0', 'TROUBLE_DESC', 'TROUBLE_STIME', 'TROUBLE_ETIME'])
df5['Date'] = pd.to_datetime(df1['Date']).dt.strftime('%d-%b')
# # print(df1)
df6 = df5.groupby('Date')['CVD Change'].apply(list).reset_index()
df6['CVD Change'] = df6['CVD Change'].apply(lambda x: '.'.join(str(i) for i in x))
print(df6)
