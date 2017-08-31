# Adrian Osuna
# Searches a file director for .sql files and Connects to a database
# Via ODBC, executes the .sql against the database and write content back
# To a EXCEL (.xlxs) file and send email to Office Manager to facilitate and review

import pyodbc #Used to Connect To DataBase
import os
import time

#Imports Used to Read And Write To XLSX Files.
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
########################################################

#Imports Used to Connect To Gmail Account And Send Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
########################################################

def get_files(dir_path):
  header = ''
  os.chdir("C:/Users/AdrianO/Desktop/Perfect Practice/Python" + dir_path)
  loc = os.getcwd()
  files = os.listdir(loc)
  queryReport = {}

  for file_ in files:
    print "looping the files in directory provided"
    if file_.endswith(".sql"): 
      #print file_
      with open(file_, 'r') as myfile:
        data = ''
        #print "File Name ", file_
        count = 0
        for line in myfile:
          #print line
          if count == 0:
            header = str(line)
            count += 1
          else:
            data = data + line.replace('\n',' ')
          
        queryReport[file_] = data
  #print "printing header: ", header
  report = [header,queryReport]
  return report
#end get_files functions

def get_data(header, file_name, query):
  os.chdir("C:/Users/AdrianO/Desktop/Perfect Practice/Python")
  current_time = time.strftime('%m-%d-%y_%H%M')
  dest_filename = file_name.replace('.sql', current_time + '.xlsx')
  wb = Workbook()
  ws = wb.active
  
  conn_str = (
  "DRIVER={PostgreSQL Unicode};"
  "DATABASE={datebase};"
  "UID={username};"
  "PWD={password};"
  "SERVER={server_address};"
  "PORT={port_number};")
  conn = pyodbc.connect(conn_str)
  results = conn.execute(query)

  rows = results.fetchall()

  #add column header to top of excel list
  ws.append(header.split(","))
  #generic list, used to write contents pulled from database to the excel file
  lst = []

  #loop through contents pulled from database and write to excel list
  for row in rows:
    #print 'printing from the database\n',row
    lst = list(row)
    #print 'printing a copy of the row as a list\n', lst
    ws.append(lst)

  print 'writing contents to file'
  #write the file
  wb.save(dest_filename)
  wb.close()

#end get_data function

#connects to google smtp server via port 587, grabs a specific excel file and email to Office
def send_mail(send_from,send_to,subject,text,file_name,isTls=True):
  
  #create message to send
  msg = MIMEMultipart()
  msg['From'] = send_from
  msg['To'] = send_to
  msg['Date'] = formatdate(localtime = True)
  msg['Subject'] = subject
  msg.attach(MIMEText(text))

  part = MIMEBase('application', "octet-stream")
  part.set_payload(open(str(filename), "rb").read())
  encoders.encode_base64(part)
  part.add_header('Content-Disposition', 'attachment; filename="' + filename + '"')
  msg.attach(part)

  smtp = smtplib.SMTP("smtp.gmail.com",587)
  if isTls:
      smtp.starttls()
  smtp.login("{login}","{password}")
  smtp.sendmail(send_from, send_to, msg.as_string())
  smtp.quit()

#end send email function

def main():
  os.chdir("C:/Users/AdrianO/Desktop/Perfect Practice/Python")
  reports = "/reports/"
  getQuery = get_files(reports)
  for k, v in getQuery[1].iteritems():
    get_data(getQuery[0], k, v)
    current_time = time.strftime('%m-%d-%y_%H%M')
    dest_filename = k.replace('.sql', current_time + '.xlsx')
    send_mail("{sender@example.com}", "{receiver@example.com}", "Auto Generated Report " + dest_filename, "Auto Generated Report " + dest_filename, dest_filename)
  
#end main

if __name__ == '__main__':
  main()