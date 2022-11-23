def insert_db(dataframe, engine, table):
    conn = engine.connect()
    try:
        dataframe.to_sql(table, conn, if_exists='append', index=False, schema='logging')
    finally:
        conn.close()