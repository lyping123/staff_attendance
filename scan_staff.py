import sqlite3
import openpyxl
from datetime import datetime,timedelta,date
import time
import os

mydb = sqlite3.connect('attendance.db')

cursor=mydb.cursor()

cursor.execute('''CREATE TABLE IF NOT EXISTS staff_list(
                    id INTEGER PRIMARY KEY,
                    staff_name TEXT NOT NULL,
                    staff_department TEXT NOT NULL,
                    staff_id TEXT NOT NULL
                )''')

def read_excel():
    file_path="Lecturer ID Number.xlsx"
    try:
        cursor.execute("delete from staff_list")
        mydb.commit()
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=3,values_only=True):
            try:
                qry="INSERT INTO `staff_list`(`staff_name`,`staff_department`,`staff_id`) VALUES(?,?,?)"
                cursor.execute(qry,(row[1],row[2],row[3]))
                mydb.commit()

            except Exception as e:
                 print(f"An error occured send {e}")
    except Exception as e:
         print(f"An error occured send {e}")
    print("staff import success")
    
    
read_excel()
