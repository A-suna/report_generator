''' Adrian Osuna
Reads information from a JSON file, file location where to read/search
for .sql files, information to connect to a database. Execute the .sql script
against the Database all trought ODBC. Writes the results in a (EXCEL) .xlsx file
to a specific directory provided by the JSON file. It then emails the contexts
of the reports to a user provided by the JSON file.
'''
import os, time, string, json

# Imports Used to Connect To Gmail Account And Send Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

# Imports Used to Read And Write To XLSX Files.
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import pyodbc #Used to Connect To DataBase

NUM_TO_ALPHA = dict(zip(range(1, 27), string.ascii_uppercase))

def get_files(dir_path):
    ''' Gets all Query Files and stores them in a dict
    with the Key being the filename and the value being
    a list of the content of the file. The list is at most
    composed of 2 elements and, the first element is the header
    for the .xlsx file and the second element is the query to
    execute against the DB
    '''
    header = ''
    os.chdir(dir_path)
    files = os.listdir(dir_path)
    query_report = {} # key=file name, value=[header for xlsx, query to run]

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
                        data = data + line.replace('\n', ' ')

                lst.append(data)
                query_report[file_] = lst
        header = ''

    return query_report
# end get_files functions

def write_xlsx(header, file_name, query, conn_str):
    '''
        Writes the results of the query to an .xlsx file,
        the contents is stored in a table, and has the columns width
        expanded for ease of use and makes it easier to read.
    '''
    print file_name
    dest_filename = file_name.replace('.sql', '_' + time.strftime('%m-%d-%y_%H%M') + '.xlsx')
    wb = Workbook()
    ws = wb.active
    conn = pyodbc.connect(conn_str)
    result = conn.execute(query).fetchall()

    if result:
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
                new_item = new_item.replace('\x1f', '') # \x1f breaks writing xlsx files with tables
                new_item = new_item.replace('None', '')
                convert_str.append(new_item)
            ws.append(convert_str)

        # Expand all Columns to match Larges Text in column
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) for cell in column_cells) * 1.3
            ws.column_dimensions[column_cells[0].column].width = length

        # Add table formatting to the Excel List, Only work to letter Z
        # breaks if query is larger than 26 columns
        end_column = NUM_TO_ALPHA[len(list(ws.rows)[0])] + str(len(list(ws.rows)))
        tab = Table(displayName='Table', ref='A1:'+end_column)
        style = TableStyleInfo(name="TableStyleLight11", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)

        # write the file
        wb.save(dest_filename)
        wb.close()
        return dest_filename

# end write_xlsx function

def send_mail(send_from, send_to, subject, text, filename):
    '''
    make connection to smtp server and port provided in the
    instruction.json file. Create message and send to speficied user
    '''
    msg = MIMEMultipart()
    msg['From'] = send_from['email']
    msg['To'] = send_to['email']
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(str(filename), "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="' + filename + '"')
    msg.attach(part)

    smtp = smtplib.SMTP(send_from['host'], send_from['port'])
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()

    smtp.login(send_from['user'], send_from['pass'])
    smtp.sendmail(send_from['email'], send_to['email'], msg.as_string())
    smtp.quit()
# end send email function

def main():
    '''
    Reads Json File with directory location,
    Gets email information and generates .xlsx
    files and emails provided emails accounts.
    '''
    f_json = None
    try:
        f_json = open('instruction.json')
        data = json.load(f_json)
        main_dir = data['Directories']['MainDir']
        os.chdir(main_dir)
        dirs = {} # KEY = file directory Value = {Key = reportname.sql value }
        for f_dir, v_dict in data['Directories']['reports'].iteritems():
            if isinstance(v_dict, list):
                try:
                    for d_dir in v_dict:
                        dirs[d_dir] = get_files(main_dir+d_dir)
                except:
                    print"Directory not found"

        os.chdir(main_dir + data['Directories']['Write_Report'])
        xlsx_files = []
        # key=file name, value=[header for xlsx, query to run]
        for k_main, v_main in dirs.iteritems():
            for k_sub, v_sub in v_main.iteritems():
                xlsx_files.append(write_xlsx(v_sub[0], k_sub, v_sub[1],
                                             data['ConnectionString']['setup']))

        for files_ in xlsx_files:
            # print type(files_), files_
            if files_ != None:
                # subject = 'Auto Generated Report ' + files_
                if 'Collection' in files_:
                    send_mail(data['SendFrom'], data['SendTo']['Adrian'], subject, subject, files_)
                else:
                    send_mail(data['SendFrom'], data['SendTo']['Adrian'], subject, subject, files_)
            else:
                pass
    finally:
        if f_json is not None:
            f_json.close()
        else:
            print "Cannot locate file \"instruction.json\""
# end main

if __name__ == '__main__':
    main()
