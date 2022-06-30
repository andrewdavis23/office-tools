import pyodbc
from openpyxl import Workbook
import pandas as pd 
from tkinter import filedialog

cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=;"
                      "Database=;"
                      "Trusted_Connection=yes;")
# cursor = cnxn.cursor()

wb = Workbook()
ws = wb.active
ws.title = 'data'

sql = ''

try:
    df = pd.read_sql(sql, cnxn)

except Exception as e:
    print('Query failed. Run it manually.')
    print(sql)

out_path = filedialog.askdirectory()

report_name = 'report.xlsx'
out_path += '/'+report_name

wb.save(out_path)
xlr = pd.ExcelWriter(report_name)
df.to_excel(xlr, 'data')
xlr.save()


