import requests ,sqlite3
from datetime import datetime,timedelta,date

mydb = sqlite3.connect('attendance.db')
cursor=mydb.cursor()

def export_daily_attendance():
    query="SELECT * FROM staff_list order by staff_name"
    cursor.execute(query)
    staffs=cursor.fetchall()
    current_date=datetime.now().strftime("%Y-%m-%d")
    message="Today's Staff Attendance:\n"
    message+=f"Date: {current_date}\n"
    send_status=0
    for staff in staffs:
        staff_id=staff[0]
        staff_name=staff[1]
        
        qry=f"SELECT  GROUP_CONCAT(time_checkin) as timecheckin FROM staff_attendance where DATE(date_checkin)=DATE('{current_date}') AND staff_id='{staff[3]}' AND (time_section='morning' or time_section='afternoon' or time_section='other') group by date_checkin"
        cursor.execute(qry)
        row=cursor.fetchone() 
            
        if row is None:
            message+=f"{staff[1]} : no attend today \n"
            send_status+=1
        
    return message,send_status
    
    

def remind_whatapp():
    ACCESS_TOKEN = 'b4c1ff649fbc1173a2d03776a97860e22e77d87f4cf4235aba92b6b18ee54aa5'
    TO_PHONE_NUMBER = '60168813607'
    
    
    url = f"https://onsend.io/api/v1/send"
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {ACCESS_TOKEN}',
        'Content-Type': 'application/json',
    }
    
    Phone_numbers=['60129253398','60124859595','60164456145','60164456145']
    message,status=export_daily_attendance()
    for pnumber in Phone_numbers:
        data={
            'phone_number':pnumber,
            'message':message,
        }
        if status>0:
            response = requests.post(url, headers=headers, json=data)
            if response.status_code == 200:
                print('Message sent successfully')
            else:
                print('Failed to send message:', response.json())

remind_whatapp()