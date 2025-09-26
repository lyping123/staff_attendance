import mysql.connector
import openpyxl

mydb = mysql.connector.connect(
    host="sv94.ifastnet.com",
    user="synergyc",
    password="synergy@central",
    database="synergyc_staffleave",
)
cursor=mydb.cursor()
wookbook = openpyxl.load_workbook('Lecturer ID Number.xlsx')
sheet=wookbook.active
count=0
for row in sheet.iter_rows(min_row=3,values_only=True):
    qry="UPDATE user set staff_id=%s where u_name like %s"
    cursor.execute(qry,(row[3],row[1]))
    mydb.commit()

mydb.close()
print("staff import success")
