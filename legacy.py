import csv
import re
from openpyxl import Workbook
import os
import datetime

with open('attendance.csv', newline='', encoding="UTF-16") as attendance:
    attendanceLog = csv.reader(attendance, delimiter = ' ', quotechar = '|')
    possibleLateComers = []
    onTime = []
    for row in attendanceLog:
        if re.search("\d{2}:\d{2}:\d{2}", row[len(row)-2]) :
            timestamp = row[len(row)-2] #Store time of individual
            if int(timestamp.split(':')[1])>30:
                possibleLateComers.append(' '.join(row).split('\t'))
            else:
                onTime.append(' '.join(row).split('\t'))

    save_location =  os.path.join(os.getcwd(), "attendance_"+str(datetime.datetime.now()))
    os.mkdir(save_location)

    #Save in Workbook for on time joiners

    presentWb = Workbook()
    ws = presentWb.active
    ws.append(['Name','Actions','Timestamp'])
    for row in onTime:
        ws.append(row)
    presentWb.save(os.path.join(save_location, "present.xlsx"))

    #Save in Workbook for late joiners

    absentWb = Workbook()
    ws = absentWb.active
    ws.append(['Name', 'Actions', 'Timestamp'])
    
    for i in range(0, len(possibleLateComers)):
        student=possibleLateComers[i]
        if student[1] == "Left":
            didJoinBack = False
            for j in range(i+1, len(possibleLateComers)):
               temp = possibleLateComers[j]
               if(temp[0]==student[0]):
                   didJoinBack = True
            if didJoinBack == False :
                pass
                #print(temp[0] + "left early") 
        else:
            #pass
            ws.append(student)

    absentWb.save(os.path.join(save_location, "late_comers.xlsx"))
