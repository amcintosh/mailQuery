#!/usr/bin/python

import cx_Oracle
import os
import sys
import smtplib
import ConfigParser
from datetime import date, timedelta
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText


__version__ = "0.1"
__date__ = "2012/05/01"
__updated__ = "2013/02/01"
__author__ = "Andrew McIntosh (github.com/amcintosh)"
__copyright__ = "Copyright 2012, Andrew McIntosh"
__license__ = "GPL"


def usage():
    '''Print usage if no config file provided.'''
    print "Usage: mailQuery <query_file>"


def get_connection(db_config):
    '''Connect to Oracle database.
       Connection string can either use TNS or full server/service 
       path as specified in the config file.
    '''
    connect_str = ""
    try:
        if db_config.has_option("DBConfig","dbTNS"):
            connect_str = (db_config.get("DBConfig","dbUser") +
                          "/"+db_config.get("DBConfig","dbPassword") +
                          "@"+db_config.get("DBConfig","dbTNS"))
        else:
            connect_str = (db_config.get("DBConfig","dbUser") +
                          "/"+db_config.get("DBConfig","dbPassword") +
                          "@"+db_config.get("DBConfig","dbHost") +
                          ":"+db_config.get("DBConfig","dbPort") +
                          "/"+db_config.get("DBConfig","dbService"))
    except ConfigParser.NoOptionError as err:
        print ("Required database configuration options missing "
               "from configuration file:"), err
        sys.exit(2)
    return cx_Oracle.connect(connect_str)


def construct_params(param_str):
    '''Takes a string of parameters and constructs a list from them.
       Date placeholders are replaced with date objects. All other 
       strings are copied as is.
    '''
    params = param_str.split(",")
    new_params = []
    for param in params:
        if param.strip()=="FIRSTOFTHISMONTH":
            first = date.today().replace(day=1)
            new_params.append(first)
        elif param.strip()=="FIRSTOFLASTMONTH":
            last_month = (date.today().replace(day=1) - timedelta(days=1)).replace(day=1)
            new_params.append(last_month)
        elif param.strip()=="FIRSTOFNEXTMONTH":
            next_month = (date.today().replace(day=1) + timedelta(days=32)).replace(day=1)
            new_params.append(next_month)
        else:
            new_params.append(param.strip())
    return new_params


def write_csv(cursor, out_file_name):
    '''Write data from a cursor to the specified file in csv format'''
    outfile = open(out_file_name, 'w')
    col_names = []
    for i in range(0, len(cursor.description)):
        col_names.append(cursor.description[i][0])
    outfile.write(",".join(col_names)+"\n")
    
    for row in cursor:
        outfile.write('"')
        outfile.write('","'.join(map(str, row)))
        outfile.write('"\n')
    outfile.close()

	
def write_queries_to_csv(connection, file_config):
    '''Iterates over and executes queries specified in config file.
    Each query result set is writen to a csv file with the file name 
    again specified in the config file.
    '''
    out_file_names = []
    for section in file_config.sections():
        if section == "MailConfig" or section == "DBConfig": 
            continue
        cursor = connection.cursor()
        cursor.arraysize = 256
        try:
            params = construct_params(file_config.get(section, "Params"))
            cursor.execute(file_config.get(section, "Query"), params)

            out_file_name = file_config.get(section, "Filename")
            write_csv(cursor, out_file_name)
            out_file_names.append(out_file_name)
        except ConfigParser.NoOptionError as err:
            print "Query configuration options missing from config file:", err
            sys.exit(2)
        cursor.close()
    return out_file_names
	
		
def mail_csv(mail_config, files_to_mail):
    '''Sends email with specified files as attachments.
       Mail server, from, to, subject, etc. should all
       be defined in the configuration file.
    '''
    from_email = mail_config.get("MailConfig","mailFrom")
    to_email = mail_config.get("MailConfig","mailTo")
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = mail_config.get("MailConfig","mailSubject")
    text = mail_config.get("MailConfig","mailBody")
    msg.attach( MIMEText(text,'html') )
    for file_to_mail in files_to_mail:
        file_payload = open(file_to_mail, 'rb')
        part = MIMEBase('application', "octet-stream")
        part.set_payload(file_payload.read())
        part.add_header('Content-Disposition', 'attachment', 
                        filename=file_to_mail)
        msg.attach(part)
    smtp = smtplib.SMTP(mail_config.get("MailConfig","mailServer"), 
                        mail_config.get("MailConfig","mailPort"))
    smtp.sendmail(from_email, to_email.split(", "), msg.as_string())
    smtp.close()


def clean_up_files(files_to_remove):
    '''Deletes the specified files for the local directory.'''
    for file_to_remove in files_to_remove:
        os.remove(file_to_remove)
	
	
def main(argv):
    if len(argv)<2 or argv[1]=="usage":
        usage()
        sys.exit(2)
    
    file_config = ConfigParser.ConfigParser()
    file_config.read(argv[1])
    
    con = get_connection(file_config)
    files_to_mail = write_queries_to_csv(con, file_config)
    mail_csv(file_config, files_to_mail)
    try:
        if file_config.get("MailConfig","attachmentCleanup")=="true":
            clean_up_files(files_to_mail)
    except ConfigParser.NoOptionError:
        #Do nothing. attachmentCleanup not required
        pass
    con.close()


if __name__ == "__main__":
    main(sys.argv)