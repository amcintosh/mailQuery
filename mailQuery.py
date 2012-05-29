#!/usr/bin/python

#Default connection params
DB_USER="andrew"
DB_PASSWORD="andrew"
DB_TNS="dev10g"

#import cx_Oracle
import sys
import datetime
import smtplib
import ConfigParser
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


def getQueryDataFromFile(filepath):
    '''Read config file and parse all queries and their associated config.
    Returns a list of dictionaries, where each dictionary is a separate 
    query to run containing the parameters for that query.
    '''
    Config = ConfigParser.ConfigParser()
    Config.read(filepath)
    queryData = []
    for section in Config.sections():
        thisQuery = {}
        for name,value in Config.items(section):
            if name!="params":
                thisQuery[name] = value
            else:
                thisQuery[name] = constructParams(value)
        queryData.append(thisQuery)
    return queryData


def constructParams(paramStr):
    '''Takes a string of parameters and constructs a list from them.
       Date placeholders are replaced with date objects. All other 
       strings are copied as is.
    '''
    params = paramStr.split(",")
    newParams = []
    for param in params:
        if param.strip()=="FIRSTOFTHISMONTH":
            pass
        elif param.strip()=="FIRSTOFLASTMONTH":
            pass
        else:
            newParams.append(param.strip())
    return newParams


def writeCsv(cursor,outfilename):
    outfile = open(outfilename, 'w')
    outfile.write('"')
    for row in cursor:
        outfile.write('","'.join(map(str,row)))
    outfile.write('"')
    outfile.close()


def mailCsv(outfilename):
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

    query = getQueryDataFromFile(argv[1])
    print query
    if len(argv)>5:
        con = getConnection(argv[2],argv[3],argv[4])
    else:
        con = getConnection()
    cursor = con.cursor()
    cursor.arraysize = 256
    cursor.execute(query)
    outfilename = "output.txt"
    writeCsv(cursor,outfilename)
    mailCsv(outfilename)
    
    cursor.close()
    con.close()

if __name__ == "__main__":
    main(sys.argv)
