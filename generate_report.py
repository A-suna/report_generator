# Adrian Osuna
# Reads information from a JSON file, file location where to read/search
# for .sql files, information to connect to a database. Execute the .sql script
# against the Database all trought ODBC. Writes the results in a (EXCEL) .xlsx file
# to a specific directory provided by the JSON file. It then emails the contexts 
# of the reports to a user provided by the JSON file.

import pyodbc #Used to Connect To DataBase
import os, time, calendar, string, json
from datetime import date

# Imports Used to Read And Write To XLSX Files.
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
########################################################

# Imports Used to Connect To Gmail Account And Send Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
########################################################

numberToAlpha = dict(zip(range(1, 27), string.ascii_uppercase))

def get_files(dir_path):
  header = ''
  os.chdir(dir_path)
  files = os.listdir(dir_path)
  queryReport = {} # key=file name, value=[header for xlsx, query to run]

  for file_ in files:
    #print 'looping the files in directory provided'
    if file_.endswith(".sql"): 
      lst = []
      # print file_
      with open(file_, 'r') as myfile:
        data = ''
        # print "File Name ", file_
        write_header = False
        for line in myfile:
          # print line
          if not write_header:
            header = str(line)
            lst.append(header)
            write_header = True
          else:
            data = data + line.replace('\n',' ')
          
        lst.append(data)
        queryReport[file_] = lst 
    header = ''
  
  return queryReport
# end get_files functions

def write_xlsx(header, file_name, query, conn_str):
  current_time = time.strftime('%m-%d-%y_%H%M')
  dest_filename = file_name.replace('.sql', '_' + current_time + '.xlsx')
  wb = Workbook()
  ws = wb.active
  conn = pyodbc.connect(conn_str)
  result = conn.execute(query).fetchall()
  
  if len(result) > 0:
    # add column header to top of excel list
    ws.append(header.split(','))
    # generic list, used to write contents pulled from database to the excel file
    lst = []
    
    # loop through contents pulled from database and write to excel list
    for row in result:
      convert_str = []

      lst = list(row)
      for item in lst:
        new_item = str(item)
        new_item = new_item.replace('\x1f','')
        new_item = new_item.replace('None','')
        convert_str.append(new_item)
      ws.append(convert_str)
    
    # print 'writing contents to file, with formating: ', dest_filename

    # Expand all Columns to match Larges Text in column
    for column_cells in ws.columns:
      length = max(len(str(cell.value)) for cell in column_cells) * 1.3
      ws.column_dimensions[column_cells[0].column].width = length

    # Add table formatting to the Excel List, Only work to letter Z
    # breaks if query is larger than 26 columns
    endColumn = numberToAlpha[len(list(ws.rows)[0])] + str(len(list(ws.rows)))
    tab = Table(displayName='Table', ref='A1:'+endColumn)
    style = TableStyleInfo(name="TableStyleLight11", showFirstColumn=False, showLastColumn=False,
                           showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    # write the file
    wb.save(dest_filename)
    wb.close()
    return dest_filename

# end write_xlsx function

# connects to google smtp server via port 587, grabs a specific excel file and email to Office
def send_mail(send_from,send_to,subject,text,filename,isTls=True):
  # create message to send
  msg = MIMEMultipart()
  msg['From'] = send_from['email']
  msg['To'] = send_to['email']
  msg['Date'] = formatdate(localtime = True)
  msg['Subject'] = subject
  msg.attach(MIMEText(text))

  part = MIMEBase('application', 'octet-stream')
  part.set_payload(open(str(filename), "rb").read())
  encoders.encode_base64(part)
  part.add_header('Content-Disposition', 'attachment; filename="' + filename + '"')
  msg.attach(part)

  smtp = smtplib.SMTP(send_from['host'],send_from['port'])
  smtp.ehlo()
  if isTls:
    smtp.starttls()
    smtp.ehlo()
  
  smtp.login(send_from['user'],send_from['pass'])
  smtp.sendmail(send_from['email'], send_to['email'], msg.as_string())
  smtp.quit()

# end send email function

def main():
  data = json.load(open('instruction.json'))
  main_dir = data['Directories']['MainDir']
  os.chdir(main_dir)
  dirs = {} # KEY = file directory Value = {Key = reportname.sql value }
  for k,v in data['Directories']['reports'].iteritems():
    if type(v) == list:
      try:
        for d in v:
          dirs[d] = get_files(main_dir+d)
      except:
        print("Directory not found")

  os.chdir(main_dir + data['Directories']['Write_Report'])
  xlsx_files = []
  # key=file name, value=[header for xlsx, query to run]
  for km, vm in dirs.iteritems():
    for ks, vs in vm.iteritems():
      xlsx_files.append( write_xlsx(vs[0], ks, vs[1], data['ConnectionString']['main']) )
  
  for files_ in xlsx_files:
    # print type(files_), files_
    if files_ != None:
      subject = 'Auto Generated Report ' + files_
      if 'Collection' in files_: 
        send_mail(data['SendFrom'], data['SendTo']['Adrian'], subject, subject, files_)
      else:
        send_mail(data['SendFrom'], data['SendTo']['Adrian'], subject, subject, files_)
    else:
      pass

# end main

if __name__ == '__main__':
  main()
