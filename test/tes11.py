from database import Postgres
from time import perf_counter
# sql = """
# SELECT
#     id,
#     line,
#     chamber
# FROM
#     pm.chamber_config
# INNER JOIN pm.line USING (line_id)
# INNER JOIN pm.chamber USING (chamber_id)
# WHERE line <> 212 and chamber in ('CRC1', 'CRC2', 'CRC3', 'CRC4', 'CRC5', 'CRC6', 'CRC7', 'CRC8', 'CRC9');
# """

table = 'action_view'
sql = f"SELECT * FROM pm.{table}"
begin = perf_counter()
myPg = None
myPg = Postgres("128.53.1.198/5432", "spt_db", "spt_admin",
                "sptadmin")
df1 = myPg.query(f"{sql}")[0]
# df1 = df1.rename(columns={'id': 'id_'})
end = perf_counter()
df1.to_excel(f'{table}.xlsx', index=False)
# print(df1)

print('done fetching')
print(f"pd read {(end - begin):.4f}s")
if myPg is not None:
    myPg.close()
