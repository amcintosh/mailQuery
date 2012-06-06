#!/usr/bin/python

import cx_Oracle
import os
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
    print "Usage: mailQuery <query_file>"


def getConnection(dbConfig):
    '''Connect to Oracle database'''
    connectStr = ""
    try:
        connectStr = dbConfig.get("DBConfig","dbUser") \
                +"/"+dbConfig.get("DBConfig","dbPassword") \
                +"@"+dbConfig.get("DBConfig","dbTNS")
    except ConfigParser.NoOptionError as err:
        print "Required database configuration options missing from configuration file:",err
        sys.exit(2)
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
    colNames = []
    for i in range(0, len(cursor.description)):
        colNames.append(cursor.description[i][0])
    outfile.write(",".join(colNames)+"\n")
    
    for row in cursor:
        outfile.write('"')
        outfile.write('","'.join(map(str,row)))
        outfile.write('"\n')
    outfile.close()

	
def writeQueriesToCSV(connection, fileConfig):
    '''Iterates over and executes queries specified in config file.
    Each query result set is writen to a csv file with the file name 
    again specified in the config file.
    '''
    outFileNames = []
    for section in fileConfig.sections():
        if section=="MailConfig" or section=="DBConfig": 
            continue
        cursor = connection.cursor()
        cursor.arraysize = 256
        try:
            params = constructParams(fileConfig.get(section,"Params"))
            cursor.execute(fileConfig.get(section,"Query"),params)

            outFileName = fileConfig.get(section,"Filename")
            writeCsv(cursor,outFileName)
            outFileNames.append(outFileName)
        except ConfigParser.NoOptionError as err:
            print "Query configuration options missing from config file:", err
            sys.exit(2)
        cursor.close()
    return outFileNames
	
		
def mailCsv(mailConfig, filesToMail):
    fromEmail = mailConfig.get("MailConfig","mailFrom")
    toEmail = mailConfig.get("MailConfig","mailTo")
    msg = MIMEMultipart()
    msg['From'] = fromEmail
    msg['To'] = toEmail
    msg['Subject'] = mailConfig.get("MailConfig","mailSubject")
    text = mailConfig.get("MailConfig","mailBody")
    msg.attach( MIMEText(text) )
    for file in filesToMail:
        fp = open(file, 'rb')
        part = MIMEBase('application', "octet-stream")
        part.set_payload(fp.read())
        part.add_header('Content-Disposition', 'attachment', filename=file)
        msg.attach(part)
    smtp = smtplib.SMTP(mailConfig.get("MailConfig","mailServer"),mailConfig.get("MailConfig","mailPort"))
    smtp.sendmail(fromEmail, toEmail, msg.as_string())
    smtp.close()


def cleanUpFiles(filesToRemove):
    for file in filesToRemove:
        os.remove(file)
	
	
def main(argv):
    if len(argv)<2 or argv[1]=="usage":
        usage()
        sys.exit(2)
    
    fileConfig = ConfigParser.ConfigParser()
    fileConfig.read(argv[1])
    
    con = getConnection(fileConfig)
    filesToMail = writeQueriesToCSV(con, fileConfig)
    #mailCsv(fileConfig,filesToMail)
    try:
        if fileConfig.get("MailConfig","attachmentCleanup")=="true":
		    cleanUpFiles(filesToMail)
    except ConfigParser.NoOptionError:
	    #Do nothing. attachmentCleanup not required
		pass
    con.close()


if __name__ == "__main__":
    main(sys.argv)