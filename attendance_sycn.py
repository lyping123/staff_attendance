import openpyxl
import sqlite3
import calendar
import math
from datetime import datetime, date

mydb = sqlite3.connect('attendance.db')

cursor=mydb.cursor()

def load_staff_name(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        rows=[]
        headers = [cell.value for cell in sheet[1]] 
        
        for row in sheet.iter_rows(values_only=True, min_row=2, min_col=1):
            staff_name=row
            rows.append(staff_name)
        return rows, headers
            
    except Exception as e:
        print(f"Error loading workbook: {e}")
        return None, None

def sycn_attendance():
    staff_list,headers=load_staff_name("staff_attendance2025(September).xlsx")
    
    for staff in staff_list:
        for col_num,header in enumerate(headers[1:-1],start=2):
            staff_name=staff[0]
            time_checkin=staff[col_num-1]
            
            array_leave=["U","AL","MC","H"]
            if time_checkin not in array_leave:
                time_format=f"{time_checkin}:00" if len(str(time_checkin))<=5 else f"{time_checkin}"
            else:
                time_format=time_checkin
            
            year_month_date=header+"-"+datetime.now().strftime("%Y")
            parsed_date = datetime.strptime(year_month_date, "%d-%b-%Y")  # parse text
            sql_date = parsed_date.strftime("%Y-%m-%d")
            
            cursor.execute("""
                SELECT id FROM staff_attendance
                WHERE staff_id = (
                    SELECT staff_id FROM staff_list WHERE staff_name = ?
                )
                AND date_checkin = ?
                ORDER BY time_checkin ASC
                LIMIT 1
            """, (staff_name, sql_date))

            row = cursor.fetchone()
            if row:
                attendance_id = row[0]
                cursor.execute("""
                    UPDATE staff_attendance
                    SET time_checkin = ?
                    WHERE id = ?
                """, (time_format, attendance_id))
                mydb.commit()
            
            
            
        
file=load_staff_name("staff_attendance2025(September).xlsx")
sycn_attendance()
