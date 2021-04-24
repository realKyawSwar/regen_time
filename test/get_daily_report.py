from database import Oracle


db_config = {
        "VIEW": {
            "server": "128.53.1.45:1521/ORCL",
            "user": "EP_VIEW",
            "password": "EP_VIEW"
        }
}


def get_connection(db_config):
    try:
        conn = None
        # INIT ORACLE CONNECTIONS
        server = "VIEW"
        conn = Oracle(db_config[server]['server'], "",
                      db_config[server]['user'],
                      db_config[server]['password'])
        conn.connect()
        return conn
    except Exception as e:
        raise(e)


def get_report(*args):
    SqlStr = f"SELECT * FROM V_TROUBLE WHERE TROUBLE_DATE BETWEEN"\
             f"'{(args[1])}' AND '{(args[2])}'"\
             "AND PROCESS_DESC= 'SPUTTER'"\
             "ORDER BY TROUBLE_DATE"
    # HEADER CALL
    results = args[0].execute(SqlStr)
    df = args[0].result_to_dataframe(results)
    return df


date_1 = '2021/01/01'
date_2 = '2021/04/23'
view_conn = get_connection(db_config)
df = get_report(view_conn, date_1, date_2)
df = df.drop(columns=['MACHINE_ID', 'PROCESS_ID', 'PROCESS_DESC',
                      'EQUIP_ID', 'COMMENT_TIME', 'EQUIP_DESC',
                      'UPDATED_DATE', 'UPDATED_TIME', 'ENGINEER_COMMENT',
                      'COMMENT_BY', 'COMMENT_DATE',
                      'ATTEND_BY', 'UPDATED_BY', 'CAUSES', 'TROUBLE_ID'])
df.to_excel('trouble_report.xlsx', index=False)
print(df.head(3))
