## Report Generator
generate_report.py connects uses pyodbc to connect to a database and query results and then send information to specific users via email provided in a .json file. The results queried from the database are then written to a EXCEL(.xlsx) file. Files are generated for each query maded against the Database.

Each query is read from a .sql file. These queries are written prior to executing the script. 

***
### JSON
JSON library is used to get database information:
- SMTP protocol 
  + (Connection port number, domain host, username, and password)
- a file directory and location where to read and write files. 
- Person of contact, who to send the reports to
- Special reports to be pulled for certain clients, and report is sent to a Manger to review.
