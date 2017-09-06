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
  queryReport = {} #key=file name, value=[header for xlsx, query to run]

  for file_ in files:
    print "looping the files in directory provided"
    if file_.endswith(".sql"): 
      lst = []
      #print file_
      with open(file_, 'r') as myfile:
        data = ''
        #print "File Name ", file_
        count = 0
        for line in myfile:
          #print line
          if count == 0:
            header = str(line)
            lst.append(header)
            count += 1
          else:
            data = data + line.replace('\n',' ')
          
        lst.append(data)
        queryReport[file_] = lst 
    header = ''
  
  return queryReport
#end get_files functions

def write_xlsx(header, file_name, query):
  current_time = time.strftime('%m-%d-%y_%H%M')
  dest_filename = file_name.replace('.sql', '_' + current_time + '.xlsx')
  wb = Workbook()
  ws = wb.active
  
  conn_str = (
  "DRIVER={ServerType};"
  "DATABASE={ServerName};"
  "UID={DBUserName};"
  "PWD={DBPass};"
  "SERVER={ServerAddress};"
  "PORT={ServerPort};")
  conn = pyodbc.connect(conn_str)
  results = conn.execute(query)

  rows = results.fetchall()
  
  if len(rows) > 0:
    #add column header to top of excel list
    ws.append(header.split(","))
    #generic list, used to write contents pulled from database to the excel file
    lst = []
    
    #loop through contents pulled from database and write to excel list
    for row in rows:
      convert_str = []
      lst = list(row)
      for item in lst:
        new_item = str(item)
        new_item = new_item.replace('\x1f','')
        convert_str.append(new_item)
      ws.append(convert_str)

    print 'writing contents to file: ', dest_filename
    #write the file
    wb.save(dest_filename)
    wb.close()
    return dest_filename

#end write_xlsx function

#connects to google smtp server via port 587, grabs a specific excel file and email to Office
def send_mail(send_from,send_to,subject,text,filename,isTls=True):
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
  smtp.login("{SenderUserName}","{SenderUserPass}")
  smtp.sendmail(send_from, send_to, msg.as_string())
  smtp.quit()

#end send email function

def main():
  os.chdir("C:/Users/AdrianO/Desktop/Perfect Practice/Python")
  reports = "/report_queries/"
  getQuery = get_files(reports)
  os.chdir("C:/Users/AdrianO/Desktop/Perfect Practice/Python/reports")
  loc = os.getcwd()
  xlsx_files = []
  #key=file name, value=[header for xlsx, query to run]
  for k, v in getQuery.iteritems():
    xlsx_files.append( write_xlsx(v[0], k, v[1]) )
  
  for files_ in xlsx_files:
    print type(files_), files_
    if files_ == None:
      pass
    elif '{SpecialClientReport}' in files_:
      print '{SpecialClientReport} Found'
      #{SpecialClientReport}, send email to {differentEmailReceiver} for review
    else: 
      send_mail("{senderEmail}", "{receiverEmail}", "Auto Generated Report " + files_, "Auto Generated Report " + files_, files_)
  
#end main

if __name__ == '__main__':
  main()
