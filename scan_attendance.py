import sqlite3
import tkinter as tk
from tkinter import ttk,messagebox,filedialog
from datetime import datetime,timedelta,date
import time
import winsound

from export_attendance import export_excel,export_daily

mydb = sqlite3.connect('attendance.db')

cursor=mydb.cursor()

cursor.execute('''CREATE TABLE IF NOT EXISTS staff_attendance(
                    id INTEGER PRIMARY KEY,
                    staff_id TEXT NOT NULL,
                    time_checkin TIMESTAMP NOT NULL,
                    time_section TEXT NOT NULL,
                    date_checkin TEXT NOT NULL
                )''')

# cursor.execute("delete from staff_attendance")
# mydb.commit()
def adapt_date(date_obj):
    return date_obj.isoformat()


sqlite3.register_adapter(date, adapt_date)

class clockApp:
    def __init__(self,master):
        self.master=master
        master.title("Staff Attendance System")
        master.state('zoomed')
        self.label=tk.Label(master,font=("Helvetica",72))
        self.label.pack()
        self.update_clock()
        self.label1=tk.Label(master,text="scan attendance")
        self.label1.pack()
        self.entry=tk.Entry(master)
        self.entry.focus()
        self.entry.bind("<Return>",self.submit)
        self.entry.pack()
        self.submit_button=tk.Button(master,text="submit",command=self.submit_attend)
        self.submit_button.pack()
        self.message=tk.Label(master,font=("Helvetica",36))
        self.message.pack()
        self.count=tk.Label(master,font=("Helvetica",24))
        self.count.pack()
        self.countstaff()
        self.tree=ttk.Treeview(master,height=50)
        self.scrollbar=tk.Scrollbar(master,orient="vertical",command=self.tree.yview)
        self.tree["columns"]=("column1","column2","column3","column4","column5","column6")
        # self.tree.heading("#0",text="",anchor=tk.W)
        # self.tree.detach("#0")
        self.tree.heading("#1",text="count",anchor=tk.W)
        self.tree.heading("#2",text="staff_id",anchor=tk.W)
        self.tree.heading("#3",text="staff_name",anchor=tk.W)
        self.tree.heading("#4",text="time_scan",anchor=tk.W)
        self.tree.heading("#5",text="section",anchor=tk.W)
        self.tree.heading("#6",text="date",anchor=tk.W)
        self.tree.config(yscrollcommand=self.scrollbar.set)
        self.tree.pack(side="left", fill="both",expand=True)
        self.scrollbar.pack(side="left",fill="y")
        self.load_attendance()
        
    def messagelabel(self,status):
        if status=="success":
            self.message.config(text="Success Attend",fg="green")
            self.master.after(3000,self.closemessage)
        elif status=="fail":
            self.message.config(text="You already scan before",fg="red")
            self.master.after(3000,self.closemessage)
        
    def closemessage(self):
        self.message.config(text="")
        
    def countstaff(self):
        current_date = datetime.now().date()
        qry=f"SELECT * FROM staff_attendance where DATE(date_checkin)=DATE('{current_date}')  group by staff_id"
        cursor.execute(qry)
        row=cursor.fetchall()
        count=len(row) if row is not None else 0
        self.count.config(text=f"Staff attended today:{count}",fg="blue")
    
    def update_clock(self):
        current_time = time.strftime('%H:%M:%S')
        self.label.config(text=current_time)
        self.master.after(1000,self.update_clock)
        
    def submit(self,event):
        self.submit_attend()
    
    def submit_attend(self):
        staff_id=self.entry.get()
        current_date=datetime.now().date()
        current_time = time.strftime('%H:%M')
        current_hour, current_minute = map(int, current_time.split(':'))

        current_minutes = current_hour * 60 + current_minute
        morning_start = 7 * 60
        morning_end = 11 * 60+31
        afternoon_start = 13 * 60+15
        afternoon_end = 17 * 60

        if morning_start <= current_minutes < morning_end:
            time_section = "morning"
        elif afternoon_start <= current_minutes <= afternoon_end:
            time_section = "afternoon"
        elif 11 * 60 + 30 <= current_minutes < 13 * 60:
            time_section = "breaktime"
        else:
            time_section = "other"
        
        
        # current_hourminute = int(time.strftime('%H:%M'))
        # if 8<=current_hourminute<13:
        #     time_section="morning"
        # elif 13<=current_hourminute<=17:
        #     time_section="afternoon"
        
        
        qry_staff="SELECT * FROM staff_list where staff_id=?"
        cursor.execute(qry_staff,(staff_id,))
        row_staff=cursor.fetchone()
        
        current_time=time.strftime('%H:%M:%S')
        
        if row_staff is not None:
            select="SELECT MAX(time_checkin) FROM staff_attendance where staff_id=? AND date_checkin=?"
            cursor.execute(select,(staff_id,current_date))
            row=cursor.fetchone()
            if row[0] is not None:
                time_checkin_str=row[0]
                time_checkin=datetime.strptime(time_checkin_str, '%H:%M:%S')
                time_checkin_afterfive=time_checkin+timedelta(minutes=5)
                timenow=datetime.now().time()
                
                if timenow>time_checkin_afterfive.time():
                    query="INSERT INTO `staff_attendance`(`staff_id`,`time_checkin`,`time_section`,`date_checkin`) values(?,?,?,?)"
                    cursor.execute(query,(staff_id,current_time,time_section,current_date,))
                    mydb.commit()
                    self.messagelabel("success")
                    winsound.PlaySound("audio/success.wav", winsound.SND_FILENAME)
                    self.entry.delete(0, tk.END)
                    
                else:
                    self.messagelabel("fail")
                    winsound.PlaySound("audio/fail.wav", winsound.SND_FILENAME)
                    self.entry.delete(0, tk.END)
            else:
                query="INSERT INTO `staff_attendance`(`staff_id`,`time_checkin`,`time_section`,`date_checkin`) values(?,?,?,?)"
                cursor.execute(query,(staff_id,current_time,time_section,current_date,))
                mydb.commit()
                self.messagelabel("success")
                winsound.PlaySound("audio/success.wav", winsound.SND_FILENAME)
                self.entry.delete(0, tk.END) 
                
        else:
            messagebox.showinfo("fail","Your user account is not been register yet")
        
        export_excel()
        export_daily()
        self.countstaff()
        self.load_attendance()
        
    def load_attendance(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        current_date=datetime.now().date()
        # qry=f"select sa.staff_id,sl.staff_name,sa.time_checkin,sa.time_section,sa.date_checkin from staff_attendance sa INNER JOIN staff_list sl on sl.staff_id=sa.staff_id WHERE date_checkin='{current_date}' order by sa.id desc"
        qry=f"select sa.staff_id,sl.staff_name,sa.time_checkin,sa.time_section,sa.date_checkin from staff_attendance sa INNER JOIN staff_list sl on sl.staff_id=sa.staff_id  order by sa.id desc"
        cursor.execute(qry)
        rows=cursor.fetchall()
        count=0
        for row in rows:
            count+=1
            self.tree.insert("",tk.END,values=(count,row[0],row[1],row[2],row[3],row[4]))
    
    

def main():
    root = tk.Tk()
    app = clockApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
