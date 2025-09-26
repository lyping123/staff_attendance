import sqlite3, calendar, math ,openpyxl
from datetime import datetime,timedelta,date
from excel_style import border_alignCenter,BoldFont,FillColor


boder,alignment_center=border_alignCenter()
bold=BoldFont()
fill_green=FillColor("00FF00")
fill_red=FillColor("FF0000")


mydb = sqlite3.connect('attendance.db')

cursor=mydb.cursor()

def adapt_date(date_obj):
    return date_obj.isoformat()

sqlite3.register_adapter(date, adapt_date)

def export_excel():
    try:
        workbook = openpyxl.Workbook()
        query="SELECT * FROM staff_list"
        cursor.execute(query)
        staffs=cursor.fetchall()
        row_num=2
        for num_row,staff in enumerate(staffs):           
            sheet=workbook.active
            sheet["A1"]="Staff Name"
            sheet.column_dimensions["A"].width=40
            staff_name=staff[1]
            sheet.append([staff_name])
            sheet.cell(row=num_row+2, column=1).border=boder
            
           
            cal = calendar.Calendar()
            currect_month=datetime.today().month
            currect_year=date.today().year
            month_dates = cal.monthdatescalendar(currect_year, currect_month)
            col_num=2
            staff_timein=[]
            total_lateness=0
            for week in month_dates:
                
                for day in week:
                   
                    if day.month==currect_month:
                        if day.weekday() in [5, 6]: 
                            pass
                        else:
                            # print(f"{day.day:2}")
                            
                            date_today=day
                            qry=f"SELECT GROUP_CONCAT(time_checkin) as timecheckin FROM staff_attendance where DATE(date_checkin)=DATE('{date_today}') AND staff_id='{staff[3]}' AND (time_section='morning' or time_section='afternoon' or time_section='other')  group by date_checkin"
                            cursor.execute(qry)
                            datecheckin=cursor.fetchone()
                            
                            if datecheckin is not None:
                                # qry_break=f"SELECT GROUP_CONCAT(time_checkin) as timecheckin FROM staff_attendance where DATE(date_checkin)=DATE('{date_today}') AND staff_id='{staff[3]}' AND time_section='breaktime'  group by date_checkin"
                                # cursor.execute(qry_break)
                                # breaktimecheck=cursor.fetchone()
                                # if breaktimecheck is not None:
                                #     breaktimescan=breaktimecheck[0]
                                #     breaktimes=breaktimecheck[0].split(",")
                                #     time_breaks = [datetime.strptime(breaktime, "%H:%M:%S") for breaktime in breaktimes]
                                #     breaktimelen=len(time_breaks)
                                #     if breaktimelen==2:
                                #         breaktime=time_breaks[-1]-time_breaks[0]
                                #         breaktime_duration=math.floor(breaktime.total_seconds() / 60)
                                # else:
                                #     breaktimescan=""
                                #     breaktime_duration=60
                                
                                checktime=datecheckin[0]
                                times=checktime.split(",")
                                time_objects = [datetime.strptime(time, "%H:%M:%S") for time in times]                        
                                count_timeoff=len(time_objects)
                                timeatten=0
                                timeins = [times[i] for i in range(len(times)) if i % 2 == 0]
                                timeouts = [times[i] for i in range(len(times)) if i % 2 != 0]
        
                                timelateness=0
                                if count_timeoff>1:
                                    
                                    for i in range(len(time_objects)//2):
                                        start_timeoff = datetime.strptime("08:00:00","%H:%M:%S") if time_objects[i * 2]<=datetime.strptime("08:00:00","%H:%M:%S") else time_objects[i * 2]
                                        
                                        end_timeoff = datetime.strptime("17:00:00","%H:%M:%S") if time_objects[i * 2 + 1]>=datetime.strptime("17:00:00","%H:%M:%S") else time_objects[i * 2 + 1]
                                        lateness_duration= start_timeoff - datetime.strptime("08:00:00","%H:%M:%S") 
                                        timeoff_duration = end_timeoff - start_timeoff
                                        
                                        timelateness+=math.floor(lateness_duration.total_seconds() / 60)
                                        timeatten += math.floor(timeoff_duration.total_seconds() / 60)
                                else:
                                    start_timeoff = datetime.strptime("08:00:00","%H:%M:%S") if time_objects[0]<=datetime.strptime("08:00:00","%H:%M:%S") else time_objects[0]
                                    
                                    end_timeoff = datetime.strptime("17:00:00","%H:%M:%S") if time_objects[-1]>=datetime.strptime("17:00:00","%H:%M:%S") else time_objects[-1]
                                    lateness_duration= start_timeoff - datetime.strptime("08:00:00","%H:%M:%S") 
                                    timeoff_duration = end_timeoff - start_timeoff
                                    
                                    timelateness+=math.floor(lateness_duration.total_seconds() / 60)
                                    timeatten += math.floor(timeoff_duration.total_seconds() / 60)
                                
                                total_lateness+=timelateness
                                # morning_time=datetime.strptime("08:00:00","%H:%M:%S") if time_objects[0]<=datetime.strptime("08:00:00","%H:%M:%S") else time_objects[0]
                                # afternoon_time=datetime.strptime("17:00:00","%H:%M:%S") if time_objects[-1]>=datetime.strptime("17:00:00","%H:%M:%S") else time_objects[-1]
                                # time_difference = afternoon_time - morning_time
                                # working_time = math.floor(time_difference.total_seconds() / 60)
                                
                                # if time_objects[-1]<datetime.strptime("12:00:00","%H:%M:%S"):
                                #     breaktime_duration-=60
                                
                                
                                # totaltimeatten=timeatten-breaktime_duration if timeatten-breaktime_duration >0 else 0
                               
                                totaltimeoff=540-timeatten
                                
                                # sheet.append([day.day,timein,timeout,breaktimescan,totaltimeatten,totaltimeoff])
                                day_month = f"{day.day}-{day.strftime('%b')}"
                                
                                sheet.cell(row=1, column=col_num, value=day_month).border=boder
                                sheet.cell(row=1, column=col_num, value=day_month).alignment =alignment_center
                                # Convert timein (comma-separated string) to formatted time strings
                                
                                timein_formatted = ", ".join([datetime.strptime(t, "%H:%M:%S").strftime("%H:%M") for t in timeins if t])
                                timeout_formatted = ", ".join([datetime.strptime(t, "%H:%M:%S").strftime("%H:%M") for t in timeouts if t])
                                sheet.cell(row=row_num, column=col_num, value=f"{timein_formatted}")
                                sheet.cell(row=row_num, column=col_num).border=boder
                                sheet.cell(row=row_num, column=col_num).alignment =alignment_center
                                col_num += 1
                                
                                
                            else:
                                day_month = f"{day.day}-{day.strftime('%b')}"
                                sheet.cell(row=1, column=col_num, value=day_month).border=boder
                                sheet.cell(row=1, column=col_num, value=day_month).alignment =alignment_center
                                
                                # Try to read the value from the origin Excel file if it exists
                                current_date = datetime.now()
                                current_month_name = current_date.strftime('%B')
                                
                                origin_filename = f"staff_attendance{currect_year}({current_month_name}).xlsx"
                                origin_value = ""
                                try:
                                    origin_wb = openpyxl.load_workbook(origin_filename)
                                    origin_sheet = origin_wb.active
                                    origin_cell = origin_sheet.cell(row=row_num, column=col_num)
                                    if origin_cell.value not in (None, ""):
                                        origin_value = origin_cell.value
                                except FileNotFoundError:
                                    # File not found, skip reading origin value
                                    origin_value = ""
                                except Exception:
                                    pass  # Other errors, just leave blank

                                sheet.cell(row=row_num, column=col_num, value=origin_value).border = boder
                                sheet.cell(row=row_num, column=col_num).alignment = alignment_center
                                col_num += 1

            sheet.cell(row=1, column=col_num, value="Total lateness (mins)").border=boder
            sheet.cell(row=1, column=col_num).alignment =alignment_center                    
            sheet.cell(row=row_num, column=col_num, value=f"{total_lateness}").border=boder
            sheet.cell(row=row_num, column=col_num).alignment =alignment_center
            
            col_num += 1
            row_num+=1

        current_date = datetime.now()
        current_month_name = current_date.strftime('%B')  
        workbook.save(f"staff_attendance{currect_year}({current_month_name}).xlsx")
        
    except Exception as e:
        print(f"Error occur is {e}")
        
def export_daily():
    try:
        workbook = openpyxl.Workbook()
        query="SELECT * FROM staff_list order by staff_name ASC"
        cursor.execute(query)
        staffs=cursor.fetchall()
        sheet=workbook.active
        sheet["A1"]="STAFF ID"
        sheet["B1"]="STAFF NAME"
        sheet["C1"]="TIME IN"
        row_num = 2
        sheet.column_dimensions["A"].width=15
        sheet.column_dimensions["B"].width=40
        sheet.column_dimensions["C"].width=40
        current_date = datetime.now().date()
        send_status=0
        
        message="Today's Staff Attendance:\n"
        message+=f"Date: {current_date}\n"
        for staff in staffs:
            qry=f"SELECT  GROUP_CONCAT(time_checkin) as timecheckin FROM staff_attendance where DATE(date_checkin)=DATE('{current_date}') AND staff_id='{staff[3]}' AND (time_section='morning' or time_section='afternoon' or time_section='other') group by date_checkin"
            cursor.execute(qry)
            row=cursor.fetchone() 
            
            if row is None:
                message+=f"{staff[1]} : no attend today \n"
                send_status+=1
            
            attendance=row[0] if row is not None else "no attend today"
            sheet.append([staff[3],staff[1],attendance])
            
            cell = sheet[f"B{row_num}"]
            if attendance != "no attend today":
                cell.fill = fill_green  # Green if attended
            else:
                cell.fill = fill_red
            row_num += 1
        message+=f"to visit the attendance list, please visit this link: https://drive.google.com/drive/folders/1TvSvB0Kdda4mxJarqdVbBYNoY09NwRYx?usp=sharing\n"
        sendmessage=message
        
        # remind_whatapp(sendmessage,send_status)
        
        workbook.save(f"today staff attendance.xlsx")
        
        
    except Exception as e:
        print(f"Error occur is {e}")      

        
export_excel()
export_daily()


