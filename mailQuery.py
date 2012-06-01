#!/usr/bin/python

#Default connection params
DB_USER="andrew"
DB_PASSWORD="andrew"
DB_TNS="dev10g"

import cx_Oracle
import sys
import smtplib
import ConfigParser
from datetime import date
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders


__version__ = "0.1"
__date__ = "2012/05/01"
__updated__ = "2012/05/01"
__author__ = "Andrew McIntosh (github.com/amcintosh)"
__copyright__ = "Copyright 2012, Andrew McIntosh"
__license__ = "GPL"


def usage():
    print "Usage: mailQuery <query_file> Optional: <user> <password> <tns_name> "


def getConnection(user=DB_USER,password=DB_PASSWORD,tns=DB_TNS):
    '''Connect to Oracle database'''
    connectStr = user+"/"+password+"@"+tns
    return cx_Oracle.connect(connectStr)


def constructParams(paramStr):
    '''Takes a string of parameters and constructs a list from them.
       Date placeholders are replaced with date objects. All other 
       strings are copied as is.
    '''
    params = paramStr.split(",")
    newParams = []
    for param in params:
        if param.strip()=="FIRSTOFTHISMONTH":
            today = date.today()
            newParams.append(date(today.year, today.month, 1))
        elif param.strip()=="FIRSTOFLASTMONTH":
            today = date.today()
			#TODO Make this work for january
            newParams.append(date(today.year, today.month-1, 1))
        else:
            newParams.append(param.strip())
    return newParams


def writeCsv(cursor,outfilename):
    '''Write data from a cursor to the specified file in csv format'''
    outfile = open(outfilename, 'w')
    outfile.write('"')
    for row in cursor:
        outfile.write('","'.join(map(str,row)))
    outfile.write('"')
    outfile.close()

	
def writeQueriesToCSV(connection, fileConfig):
    '''Iterates over and executes queries specified in config file.
    Each query result set is writen to a csv file with the file name 
    again specified in the config file.
    '''
    outFileNames = []
    for section in fileConfig.sections():
        if section=="MailConfig": 
            continue
        cursor = connection.cursor()
        cursor.arraysize = 256
        params = constructParams(fileConfig.get(section,"Params"))
        cursor.execute(fileConfig.get(section,"Query"),params)
        outFileName = fileConfig.get(section,"Filename")
        writeCsv(cursor,outFileName)
		outFileNames.append(outFileName)
        cursor.close()
    return outFileNames
	
		
def mailCsv(mailConfig, filesToMail):
    fp = open(outfilename, 'rb')
    msg = MIMEMultipart()
    msg['From'] = "amcintosh"
    msg['To'] = "amcintosh@something.com" #COMMASPACE.join(send_to)
    msg['Subject'] = "test"
    text = "This is my test"
    msg.attach( MIMEText(text) )
    part = MIMEBase('application', "octet-stream")
    part.set_payload(fp.read())
    part.add_header('Content-Disposition', 'attachment', filename=outfilename)
    msg.attach(part)
    smtp = smtplib.SMTP("server",25)
    smtp.sendmail("amcintosh@isomething.com", "amcintosh@something.com", msg.as_string())
    smtp.close()
    
	
def main(argv):
    if len(argv)<2 or argv[1]=="usage":
        usage()
        sys.exit(2)

    fileConfig = ConfigParser.ConfigParser()
    fileConfig.read(argv[1])

    if len(argv)>5:
        con = getConnection(argv[2],argv[3],argv[4])
    else:
        con = getConnection()
    filesToMail = writeQueriesToCSV(con, fileConfig)
    con.close()

    mailCsv(fileConfig.items("MailConfig"),filesToMail)
    

if __name__ == "__main__":
    main(sys.argv)
