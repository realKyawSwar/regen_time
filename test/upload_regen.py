from database import Postgres
import pandas as pd
pd.options.mode.chained_assignment = None
import numpy as np


def upload(dfnew):
    myPg = None
    # dfnew = pd.read_excel('to upload.xlsx')
    # dfnew = df.head(5)
    # dfnew['timing'] = dfnew['timing'].round(decimals=2)
    try:
        myPg = Postgres("128.53.1.198/5432", "spt_db", "spt_admin",
                        "sptadmin")
        myPg.connect()
        print("Connection succeeded..")
        # Convert df to list
        listy = dfnew.to_csv(None, header=False, index=False).split('\n')
        # ['2016-08-01,16.5,6', '2016-09-01,15.71,7']
        vals = [','.join(ele.split()) for ele in listy]
        # print(vals)
        for i in vals:
            y = i.split(",")
            x = tuple(y)
            if x == ('',):
                pass
            else:
                strSQL = f"INSERT INTO pm.misc_eqt values {x}"
                myPg.execute(strSQL)
        myPg.commit()
        print('done')
    except Exception as err:
        print(f"Upload Error Occured: {err}")
    finally:
        if myPg is not None:
            myPg.close()


df = pd.read_excel('misc.xlsx')
# df['Serial Number'] = df['Serial Number'].str.strip()
df.id_ = df.id_.astype(np.int64)
# print(df)
# df['install'] = pd.to_datetime(df.install)
# df['due'] = pd.to_datetime(df.due)
# df['install'] = df['install'].dt.strftime('%Y-%m-%d')
# df['due'] = df['due'].dt.strftime('%Y-%m-%d')

# # print(df.dtypes)
upload(df)
