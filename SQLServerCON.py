import os
import pyodbc
import xlsxwriter
from datetime import datetime

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

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(
    '{} WeeklySales.xlsx'.format(datetime.date(datetime.now())))
worksheet = workbook.add_worksheet()


def write(conn, gateway_name, curr_sales, last_date):
    print("Writing to Weekly Sales table: ")
    cursor = conn.cursor()
    cursor.execute(f'''
                INSERT INTO WeeklySales.dbo.Sales (gateWays, prevSales
                ,currSales, incBy, lastDate)
                VALUES ('{gateway_name}', 1
                ,'{curr_sales}', 1, '{last_date}')''')
    conn_local.commit()
    print('successfully updated')
    print()


def format_excel():
    ###### FORMAT EXCEL ######
    format_excel.cell_format = workbook.add_format({
        'bold':     True,
        'text_wrap': True,
        'align':    'center',
        'border':   3,
        'valign':   'vcenter',
        'font_color': 'white',
        'bg_color': 'black',
        'font_name': 'Calibri',
        'font_size': 12
    })
    border = workbook.add_format({'border':   5})
    format_excel.date_format = workbook.add_format({'num_format': 'd-m-yyyy'})
    worksheet.conditional_format(
        'A1:F8', {'type': 'no_blanks', 'format': border})
    worksheet.set_column(1, 5, 20)


def write_excel(cursor, row):
    format_excel()
    worksheet.write('F2', datetime.now(), format_excel.date_format)

    # Some data we want to write to the worksheet.
    expenses = (
        ['Rent', 1000],
        ['Gas',   100],
        ['Food',  300],
        ['Gym',    50],)

    titles = (
        'ID', 'Payment GateWays', 'previous Week No of Payments',
        'No of Payments', 'Increased By', 'Last Date')

    # Start from the first cell. Rows and columns are zero indexed.
    # row = 1
    col = 1

    # Iterate over the data and write it out row by row.
    worksheet.write(row, col, cursor)
    # worksheet.write(row, col + 1, cost)
    row += 1

    # Write a total using a formula.
    col = 0
    for item in (titles):
        worksheet.write(0, col, item, format_excel.cell_format)
        col += 1
    worksheet.write(0, 0, 'ID', format_excel.cell_format)
    worksheet.write(row, 0, 'Total')
    worksheet.write(row, 1, '=SUM(B1:B4)')


def read(conn, gateway_db, gateway_name):
    clients = 0
    prev_sales = 0
    curr_sales = 0
    inc_by = 0
    last_date = 0
    
    print(f"Sales: {gateway_name}")
    cursor = conn.cursor()
    cursor.execute(f"SELECT Count(id) FROM {gateway_db}.dbo.OfflineMessage")

    

    for row in cursor:
        print(f'row = {row[0]}')
        curr_sales = str(row[0])
    cursor.execute(
        f"SELECT max(deliveryDate) FROM {gateway_db}.[dbo].[OfflineMessage] where delivered=1")
    for row in cursor:
        print(f'row = {row[0]}')
        last_date = str(row[0])
    write(conn_local, gateway_name, curr_sales, last_date)
    print()


def weekly_sync():
    with open("gateways") as textFile:
        lines = [line.split() for line in textFile]
        for ip in lines:
            print(ip[0])
            read(conn, ip[0], ip[1])


weekly_sync()
# read(conn)
# write_excel()
conn_local.close()
conn.close()
workbook.close()
