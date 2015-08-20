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

import json, sys, smtplib, logging, time, datetime, csv, os
import gspread, dateutil
import json
import httplib2

import gdata.docs.service
import gdata.spreadsheet.service
import gdata.spreadsheet.text_db

import config as config

from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client.tools import run
from dateutil.tz import tzlocal


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


def getCSVData():
    """open a csv file and grab the value we care about"""

    # open csv file
    with open(config.csv_file_on_disk, 'rb') as file:
        data = list(csv.reader(file)) # cache the data
        
    #close the file
    file.close()

    # get the value we care about and return it 
    # example: data[2][3]
    return data[config.row][config.column]

def editGoogleSheet(client, data, timeStamp):
    """edit our google sheet"""
    #this will probably change to more data objects rather than 1 value and mulitple rows/colums to keep track of

    #get the current worksheet
    feed = client.GetWorksheetsFeed(config.speedsheet_id)
    id_parts = feed.entry[0].id.text.split('/')
    curr_wksht_id = id_parts[len(id_parts) - 1]

    row = config.cell_to_update[0]
    col = config.cell_to_update[1]

    date_row = config.cell_for_date[0]
    date_col = config.cell_for_date[1]

    client.UpdateCell(row, col, data, config.speedsheet_id, curr_wksht_id)
    client.UpdateCell(date_row, date_col, timeStamp, config.speedsheet_id, curr_wksht_id)

   
def sendEmail(exceptionMsg):
    """send email because i dont want the team to be unaware of mistakes"""

    # Gmail Login
    username = config.username
    password = config.password

    # message for the email
    FROM = config.username
    TO = config.recipients
    SUBJECT = config.subject
    TEXT = exceptionMsg

    # Prepare actual message
    message = """\From: %s\nTo: %s\nSubject: %s\n\n%s
    """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

    # Sending the mail  
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.ehlo()
    server.starttls()
    server.login(username, password) # https://myaccount.google.com/security need to turn on Allow less secure apps: ON
    server.sendmail(FROM, TO, message)
    server.close()


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

        # get the data we care about
        csvValue = getCSVData()

        # send data to google
        editGoogleSheet(client, csvValue, startTimeFormat)


    except Exception, e:
        # if an exception occurs, we should email an alert and log it
        # -----------------------------------------------------------
        logging.warning('Script threw an execption: ' + str(e))
        sendEmail(e.message)
        
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

