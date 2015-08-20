'''
Scheduled task to automate inserting values from csv to google spreedhseets

Installaiton notes:
1. pip install gspread
2. pip install python-dateutil
4. follow these instructions for OAuth2: http://gspread.readthedocs.org/en/latest/oauth2.html
5. update config.py to match your environment

Author: NVolpe
Date: 8/16/2015
'''

import json, sys, smtplib, logging, time, datetime, csv
import gspread, dateutil
import config as config
#import gdata.gauth
#import gdata
#import gdata.spreadsheet.service

from dateutil.tz import tzlocal
#from oauth2client.client import SignedJwtAssertionCredentials

import json
#import webbrowser
import httplib2
from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client.tools import run
import gdata.docs.service
import gdata.spreadsheet.service
import gdata.spreadsheet.text_db
#from pysqlite2 import dbapi2 as sqlite3
#import pymssql
import sys
import datetime
import os
#import argparse
#from collections import defaultdict


def getGdataCredentials():

    client_secrets = "credentials.json"
    storedCreds = "creds.dat"
    scope = ["https://spreadsheets.google.com/feeds"]
    force = False

    storage = Storage(storedCreds)
    credentials = storage.get()

    if credentials is None or credentials.invalid or force:
        credentials = run(flow_from_clientsecrets(client_secrets, scope=scope), storage)

    if credentials.access_token_expired:
        credentials.refresh(httplib2.Http())

    return credentials

def getAuthorizedSpreadsheetClient():

    client_secrets = "client_secrets.json"
    storedCreds = "creds.dat"
    force = False

    credentials = getGdataCredentials()
    client = gdata.spreadsheet.service.SpreadsheetsService(additional_headers={'Authorization' : 'Bearer %s' % credentials.access_token})

    return client


#def getCSVData():
#    """open a csv file and grab the value we care about"""

#    # open csv file
#    with open(config.csv_file_on_disk, 'rb') as file:
#        data = list(csv.reader(file)) # cache the data
        
#    #close the file
#    file.close()

#    # get the value we care about and return it 
#    # example: data[2][3]
#    return data[config.row][config.column]

#def googleLogin(value, timeStamp):
#    """log in to google and edit our google sheet"""

#    Client_id = '83927516433-bsgcm239afepee31la0sj2p9l9kk7ann.apps.googleusercontent.com'
#    Client_secret = '-vXiR6e5u38J0uI-h4boYKQ8'
#    Scope = 'https://spreadsheets.google.com/feeds/'
#    User_agent = 'python-sheets-project'

#    token = gdata.gauth.OAuth2Token(
#        client_id = Client_id,
#        client_secret = Client_secret,
#        scope = Scope,
#        user_agent = User_agent)

#    ##url = token.generate_authorize_url(redirect_uri='urn:ietf:wg:oauth:2.0:oob', approval_prompt='force')
#    print token.generate_authorize_url(redirect_uri='urn:ietf:wg:oauth:2.0:oob')

#    code = raw_input("what is the code? ")
#    token.get_access_token(code) 

#    spr_client = gdata.spreadsheet.service.SpreadsheetsService()


#    print 'yes'

   
#def sendEmail(exceptionMsg):
#    """send email because i dont want the team to be unaware of mistakes"""

#    # Gmail Login
#    username = config.username
#    password = config.password

#    # message for the email
#    FROM = config.username
#    TO = config.recipients
#    SUBJECT = config.subject
#    TEXT = exceptionMsg

#    # Prepare actual message
#    message = """\From: %s\nTo: %s\nSubject: %s\n\n%s
#    """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

#    # Sending the mail  
#    server = smtplib.SMTP("smtp.gmail.com", 587)
#    server.ehlo()
#    server.starttls()
#    server.login(username, password) # https://myaccount.google.com/security need to turn on Allow less secure apps: ON
#    server.sendmail(FROM, TO, message)
#    server.close()


def main():
    """main function"""

    try:
        #for development
        #raise Exception('something')

        print 'script started'

        # instantiate logger
        logging.basicConfig(filename='logger.log', level=logging.DEBUG)

        # Get the current date/time with the timezone.
        startTimeStamp = datetime.datetime.now(tzlocal())
        startTimeFormat = startTimeStamp.strftime("%Y-%m-%d:%S %I:%M %p %Z")

        # log start of script
        logging.info('Script Started at:  ' + startTimeFormat)

        #get/set storage
        storage = Storage("creds.dat")
        credentials = storage.get()

        if credentials is None or credentials.invalid:
          credentials = run(flow_from_clientsecrets("credentials.json", scope=["https://spreadsheets.google.com/feeds"]), storage)

        # Get authorized client
        client = getAuthorizedSpreadsheetClient()
        sheet = client.GetWorksheetsFeed('1CJjBsF2rVYhkrfTMZ2yBNzNXrd8t-7ex9WlrKa04f7c')

        print 'yay'
        # get the value we care about
        #csvValue = getCSVData()

        # send data to google
        #editGoogleSheet(123, startTimeFormat)

    except Exception, e:
        # if an exception occurs, we should email an alert and log it
        # -----------------------------------------------------------
        logging.warning('Script threw an execption: ' + str(e))
        #sendEmail(e.message)
        
        print 'Script threw an execption'
        print str(e.message)

    finally: 
        # Get the current date/time with the timezone.
        endTimeStamp = datetime.datetime.now(tzlocal())
        endTimeFormat = endTimeStamp.strftime("%Y-%m-%d %I:%M:%S %p %Z")

        # log the finish
        logging.info('Script Finished at: ' + endTimeFormat)
        logging.info('------------------------------------------------------------------------')
        logging.info('------------------------------------------------------------------------')

        print 'script finished'


# this calls the main() function
if __name__ == '__main__':
    main()

