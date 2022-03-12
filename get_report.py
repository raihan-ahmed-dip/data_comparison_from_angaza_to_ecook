import openpyxl
from openpyxl.styles import Font, Alignment
import os
import mysql.connector
import datetime

start_date = datetime.date(2022, 1, 1)
end_date = datetime.date.today() - datetime.timedelta(days=1)
delta = datetime.timedelta(days=1)

db_host = "localhost"
db_username = "root"
db_password = ""
db_name = "db_ecook_data"
db_con = mysql.connector.connect(host=db_host, user=db_username, password=db_password, database=db_name)

'''
filename = "ecook_data_comparison.xlsx"

if os.path.exists(filename):
    os.remove(filename)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Comparison"
'''
row_counter = 2

while start_date <= end_date:
    '''cell_1 = ws.cell(row=row_counter, column=1)
    cell_2 = ws.cell(row=row_counter, column=2)
    cell_3 = ws.cell(row=row_counter, column=3)
    cell_4 = ws.cell(row=row_counter, column=4)
    cell_5 = ws.cell(row=row_counter, column=5)
    cell_6 = ws.cell(row=row_counter, column=6)
    cell_7 = ws.cell(row=row_counter, column=7)
    cell_8 = ws.cell(row=row_counter, column=8)'''
    
    date = start_date
    
    bd_registered_query = f"SELECT COUNT(tbl_ecook.unit_number) AS 'BD Units' FROM tbl_ecook WHERE tbl_ecook.country = 'Bangladesh' AND tbl_ecook.registration_date <= '{date}'"
    bd_eook_registered_cursor = db_con.cursor()
    bd_eook_registered_cursor.execute(bd_registered_query)
    bd_registered_ecook = bd_eook_registered_cursor.fetchone()[0]

    bd_active_query = f"SELECT COUNT(reports_dailyusagedata.serial_number) FROM reports_dailyusagedata INNER JOIN tbl_ecook ON reports_dailyusagedata.serial_number = tbl_ecook.unit_number WHERE tbl_ecook.country= 'Bangladesh' AND reports_dailyusagedata.when_date = '{date}'"
    bd_active_ecook_cursor = db_con.cursor()
    bd_eook_registered_cursor.execute(bd_active_query)
    bd_active_ecook = bd_eook_registered_cursor.fetchone()[0]

    cam_registered_query = f"SELECT COUNT(tbl_ecook.unit_number) AS 'CAM Units' FROM tbl_ecook WHERE tbl_ecook.country = 'Cambodia' AND tbl_ecook.registration_date <= '{date}'"
    cam_eook_registered_cursor = db_con.cursor()
    cam_eook_registered_cursor.execute(cam_registered_query)
    cam_registered_ecook = cam_eook_registered_cursor.fetchone()[0]
    
    cam_active_query = f"SELECT COUNT(reports_dailyusagedata.serial_number) FROM reports_dailyusagedata INNER JOIN tbl_ecook ON reports_dailyusagedata.serial_number = tbl_ecook.unit_number WHERE tbl_ecook.country= 'Cambodia' AND reports_dailyusagedata.when_date = '{date}'"
    cam_active_ecook_cursor = db_con.cursor()
    cam_eook_registered_cursor.execute(cam_active_query)
    cam_active_ecook = cam_eook_registered_cursor.fetchone()[0]

    start_date += delta
    row_counter += 1
    print(f"{date} :: {bd_registered_ecook} :: {bd_active_ecook} :: {cam_registered_ecook} :: {cam_active_ecook}")

#wb.save(filename)