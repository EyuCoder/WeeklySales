import os
import pyodbc
import threading
from datetime import datetime, timedelta
from DBHandler import Read, Write

conn = pyodbc.connect(
    "Driver=SQL Server Native Client 11.0;"
    "Server=64.225.3.202;"
    "Integrated_Security=false;"
    "Encrypt=no;"
    "Database=master;"
    "UID=sa;"
    "PWD=cbe_sql@offline_2020;"
)

conn_local = pyodbc.connect(
    "Driver=SQL Server Native Client 11.0;"
    "Server=KINGSCXR\\SQLSERVER;"
    "Database=WeeklySales;"
    "Trusted_Connection=Yes;"
)

def sync(conn, gateway_db, gateway_name):

    date = datetime.date(datetime.now()) - timedelta(days=7)
    print("@@@" + str(date))

    curr_sales = reader.curr_sales(conn, gateway_db)
    last_date = reader.last_date(conn, gateway_db)
    prev_sales = reader.prev_sales(conn, gateway_db, date) 

    writer.write(conn_local, gateway_name, prev_sales, curr_sales, last_date)

def weekly_sync():
    with open("gateways") as textFile:
        lines = [line.split() for line in textFile]
        for ip in lines:
            print(ip[0])
            sync(conn, ip[0], ip[1])

reader = Read()
writer = Write()
weekly_sync()

conn_local.close()
conn.close()
input(" press 'Enter' to exit: ...")