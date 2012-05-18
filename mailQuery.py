#!/usr/bin/python

#Default connection params
DB_USER="andrew"
DB_PASSWORD="andrew"
DB_TNS="dev10g"

import cx_Oracle
import sys
import datetime
import smtplib
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


'''Connect to Oracle database'''
def getConnection(user=DB_USER,password=DB_PASSWORD,tns=DB_TNS):
    connectStr = user+"/"+password+"@"+tns
    return cx_Oracle.connect(connectStr)


def getQueryFromFile(filepath):
    infile = open(filepath, 'r')
    query = infile.read()
    infile.close()
    return query

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
    msg['To'] = "amcintosh@otn.ca" #COMMASPACE.join(send_to)
    msg['Subject'] = "test"
    text = "This is my test"
    msg.attach( MIMEText(text) )
    part = MIMEBase('application', "octet-stream")
    part.set_payload(fp.read())
    part.add_header('Content-Disposition', 'attachment', filename=outfilename)
    msg.attach(part)
    smtp = smtplib.SMTP("exchange.otn.local",25)
    smtp.sendmail("amcintosh@otn.ca", "amcintosh@otn.ca", msg.as_string())
    smtp.close()

    
def main(argv):
    if len(argv)<2 or argv[1]=="usage":
        usage()
        sys.exit(2)

    query = getQueryFromFile(argv[1])
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
    #mailCsv(outfilename)
    
    cursor.close()
    con.close()

if __name__ == "__main__":
    main(sys.argv)
