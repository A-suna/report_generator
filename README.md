## Report Generator
generate_report.py uses pyodbc to connect to a database and uses a JSON file as a set of instructions and a list of directories  of where to get the queries to run against the database. The data is then written to an EXCEL(.xlsx) file. Each query does generate it's own EXCEL(.xlsx) file. The JSON file also provides an email to each corresponding query that the script needs to submit the results to. Structure of the File below.

Each query is read from a .sql file. These queries are written prior to executing the script. 

***
### JSON
JSON library is used to get database information:
- SMTP protocol 
  + (Connection port number, domain host, username, and password)
- a file directory and location where to read and write files. 
- Person of contact, who to send the reports to
- Special reports to be pulled for certain clients, and report is sent to a Manger to review.
