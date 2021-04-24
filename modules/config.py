from modules.database import Oracle
import datetime

db_config = {
        "VIEW": {
            "server": "128.53.1.45:1521/ORCL",
            "user": "EP_VIEW",
            "password": "EP_VIEW"
        }
}


now = datetime.datetime.now()
last_month = now.month-3 if now.month > 1 else 12
# last_month = 12
month_name = datetime.date(1900, last_month, 1).strftime('%B')
year = now.year


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


def get_last_month_dates():
    # return 1st and last day of last month in string
    if last_month == 12:
        first_day = datetime.date(now.year-1, last_month, 1)
    else:
        first_day = datetime.date(now.year, last_month, 1)

    def last_day(any_day):
        # get close to the end of the month for any day, and add 4 days 'over'
        next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
        # subtract the number of remaining 'overage' days to get last day of current month, or said programattically said, the previous day of the first of next month
        return next_month - datetime.timedelta(days=next_month.day)

    lastly = last_day(first_day).strftime('%Y/%m/%d')
    return (first_day.strftime('%Y/%m/%d')), (lastly)


# print(get_last_month_dates())
# print(now.year - 1)


