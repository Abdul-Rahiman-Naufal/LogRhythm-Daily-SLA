# Developed by abdulrahimannaufal@gmail.com

from email.mime.base import MIMEBase
import pyodbc 
import csv
import smtplib
import subprocess
import re
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import xml.etree.ElementTree as ET
from pathlib import Path
from os.path import basename
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=ServerName;'
                      'Database=LogRhythm_Alarms;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()


query="""
select AlarmMetrics.AlarmID,Name,
Coalesce( Convert(varchar(5),abs(DateDiff(day, (cast(OpenedOn as datetime)-cast(GeneratedOn as datetime)),'1900-01-01')))  + ':'
 + Convert(varchar(10),(cast(OpenedOn as datetime)-cast(GeneratedOn as datetime)), 108),'00:00:00:00') as 'Response Time',
 Coalesce( Convert(varchar(5),abs(DateDiff(day, (cast(ClosedOn as datetime)-cast(GeneratedOn as datetime)),'1900-01-01')))  + ':'
 + Convert(varchar(10),(cast(ClosedOn as datetime)-cast(GeneratedOn as datetime)), 108),'00:00:00:00') as 'Resolution Time',
GeneratedOn,OpenedOn,ClosedOn,
(Convert(nvarchar, Case When AlarmStatus = 4 Then 'AutoClosed' When AlarmStatus = 8 Then 'Reported' When AlarmStatus = 6 Then 'Resolved' When AlarmStatus = 5 Then 'FalsePositive' When AlarmStatus = 0 Then 'New' When AlarmStatus = 1 Then 'OpenAlarm' when AlarmStatus = 9 Then 'Monitor' End)) as AlarmStatus,FullName from AlarmMetrics 
join [LogRhythm_Alarms].[dbo].[Alarm] on Alarm.AlarmID=AlarmMetrics.AlarmID 
join LogRhythmEMDB.dbo.AlarmRule on Alarm.AlarmRuleID=AlarmRule.AlarmRuleID
left join [LogRhythmEMDB].[dbo].[Person] on Person.PersonID=LastPersonID
where Alarm.AlarmDate BETWEEN getdate()-DATEADD(hour,37,0) AND  getdate()-DATEADD(hour,13,0)
and AlarmStatus not in(4)
"""
cursor.execute(query)

fields = ['Alarm ID', 'Name', 'Response Time', 'Resolution Time', 'Generated On','Opened On', 'Closed On', 'Alarm Status', 'Analyst Name'] 

    
filename = "SLA.csv"
    
with open(filename, 'w') as csvfile: 
    csvwriter = csv.writer(csvfile) 
        
    csvwriter.writerow(fields) 
        
    csvwriter.writerows(cursor)

cursor.close()
conn.close()

message = MIMEMultipart()

message.add_header('Content-Disposition', "attachment; filename= %s" % filename)


attachment = open(filename, 'rb')

part = MIMEBase("application", "octet-stream")

part.set_payload(attachment.read())

encoders.encode_base64(part)

part.add_header("Content-Disposition",f"attachment; filename= {filename}")

message.attach(part)


sender_address = "sender@email.com"
receiver_address= "receipients@email.com"


receiver_address=list(receiver_address.split(",")) 

html = """\

<html>

  <head>
<style>
table, th, td {
  border: 2px solid black;
  border-collapse: collapse;
}
</style>
  </head>

  <body>

    Dear Team,
    </br>
    </br>
    Please find SLA report attached. Note that time is in UTC.

  </body>

</html>

"""

mail_content = html
message['From'] = sender_address
message['To'] = ", ".join(receiver_address)
message['Subject'] = "Daily SLA Report"
message.attach(MIMEText(mail_content, 'html'))

session = smtplib.SMTP("SMTPServername")
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Mail Sent')
