import os
import pyodbc
import xlsxwriter
from datetime import datetime
from DBHandler import Read
import subprocess

conn_local = pyodbc.connect(
    "Driver=SQL Server Native Client 11.0;"
    "Server=KINGSCXR\\SQLSERVER;"
    "Database=WeeklySales;"
    "Trusted_Connection=Yes;"
)

workbook = xlsxwriter.Workbook(
    '{} WeeklySales.xlsx'.format(datetime.date(datetime.now())))
worksheet = workbook.add_worksheet()
def openReport():
    cwd = os.getcwd()
    excel = '\\{} WeeklySales.xlsx'.format(datetime.date(datetime.now()))
    report = cwd + excel
    subprocess.Popen(f'explorer /select, {report}"')

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
    titles = (
        'ID', 'Payment GateWays', 'previous Week No of Payments',
        'No of Payments', 'Increased By', 'Last Date')
    col= 0
    for item in (titles):
        worksheet.write(0, col, item, format_excel.cell_format)
        col+=1

def generate():
    col = 2
    row = 1
    names = 1
    format_excel()
    

    with open("WeeklySales/gateways") as textFile:
        lines = [line.split() for line in textFile]
        for ip in lines:
            worksheet.write(names, 0, names)
            worksheet.write(names, 1, ip[1])
            
            sales_list = reader.weekly_sales(conn_local, ip[1])
            print("@@@@@@@@@@@@@@@@@@@@@@@@@")
            print(sales_list)
            for item in sales_list:
                worksheet.write(row, col, item)
                col += 1
            row += 1
            col = 2
            names+=1

if __name__ == "__main__":
    reader = Read()
    generate() 

conn_local.close()
workbook.close()
openReport()
input(" press 'Enter' to exit: ...")