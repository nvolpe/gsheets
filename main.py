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
import gdata.gauth
import gdata
import gdata.spreadsheet.service

from dateutil.tz import tzlocal
from oauth2client.client import SignedJwtAssertionCredentials


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

def editGoogleSheet(value, timeStamp):
    """log in to google and edit our google sheet"""

    Client_id = '83927516433-bsgcm239afepee31la0sj2p9l9kk7ann.apps.googleusercontent.com'
    Client_secret = '-vXiR6e5u38J0uI-h4boYKQ8'
    Scope = 'https://spreadsheets.google.com/feeds/'
    User_agent = 'python-sheets-project'

    token = gdata.gauth.OAuth2Token(
        client_id = Client_id,
        client_secret = Client_secret,
        scope = Scope,
        user_agent = User_agent)

    ##url = token.generate_authorize_url(redirect_uri='urn:ietf:wg:oauth:2.0:oob', approval_prompt='force')
    print token.generate_authorize_url(redirect_uri='urn:ietf:wg:oauth:2.0:oob')

    code = raw_input("what is the code? ")
    token.get_access_token(code) 

    spr_client = gdata.spreadsheet.service.SpreadsheetsService()

    # get list of spreadsheets i am authorized to access
    #gd_client = gdata.spreadsheets.client.SpreadsheetsClient()
    #gd_client.auth_token = gdata.gauth.OAuth2TokenFromCredentials(credentials)



    #spr_client.email = 'navolpe@gmail.com'
    #spr_client.password ='fuckalex'
    #spr_client.source = 'python-sheets-project'
    #spr_client.ProgrammaticLogin()

    #documents_feed = spr_client.GetSpreadsheetsFeed('1CJjBsF2rVYhkrfTMZ2yBNzNXrd8t-7ex9WlrKa04f7c')

    print 'yes'

    # Prepare the dictionary to write
    #dict = {}
    #dict['date'] = time.strftime('%m/%d/%Y')
    #dict['time'] = time.strftime('%H:%M:%S')
    #dict['weight'] = '123'
    #print dict

    #entry = spr_client.InsertRow(dict, '1CJjBsF2rVYhkrfTMZ2yBNzNXrd8t-7ex9WlrKa04f7c', 'Sheet1')

    #if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
    #  print "Insert row succeeded."
    #else:
    #  print "Insert row failed."

    # Get the list of spreadsheets
    #feed = spr_client.GetSpreadsheetsFeed('1CJjBsF2rVYhkrfTMZ2yBNzNXrd8t-7ex9WlrKa04f7c')
    #self._PrintFeed(feed)
    #input = raw_input('\nSelection: ')
    #id_parts = feed.entry[string.atoi(input)].id.text.split('/')
    #self.curr_key = id_parts[len(id_parts) - 1]


    #spreadsheet_feed = spr_client.GetFeed('http://spreadsheets.google.com/feeds/spreadsheets/private/full')
    #first_entry = spreadsheet_feed.entry[0]
    #key = first_entry.id.text.rsplit('/')[-1]
    #worksheets_feed = spr_client.GetWorksheetsFeed(key)
    #for entry in worksheets_feed.entry:
    #    print entry.title.text

    #doc_name = 'Questions for CDOT meeting'

    #q = gdata.spreadsheet.service.DocumentQuery()
    #q['title'] = doc_name
    #q['title-exact'] = 'true'
    #feed = spr_client.GetSpreadsheetsFeed(query=q)
    #spreadsheet_id = feed.entry[0].id.text.rsplit('/',1)[1]
    #feed = spr_client.GetWorksheetsFeed(spreadsheet_id)
    #worksheet_id = feed.entry[0].id.text.rsplit('/',1)[1]


    #print spreadsheet_id

    #for document_entry in documents_feed:
    #    print 'uh'
    print 'shit'



    #print "Refresh token\n"
    #print token.refresh_token
    #print "Access Token\n"
    #print token.access_token

    ## Access the Google Documents List API.
    #docs_client = gdata.docs.client.DocsClient(source='python-sheets-project')
    ## This is the token instantiated in the first section.
    #docs_client = token.authorize(docs_client)
    #docs_feed = client.GetDocumentListFeed()
    #for entry in docs_feed.entry:
    #  print entry.title.text

    # get credentials from json file stored on disk.
    #json_key = json.load(open(config.credentials_file_on_disk))

    ## whatever this is, connects to google sheets
    #scope = ['https://spreadsheets.google.com/feeds']

    ## pluck credentials from file
    #credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'], scope)

    ## Login with your Google account
    #gc = gspread.authorize(credentials)

    ## Open a worksheet from spreadsheet with one shot
    #wks = gc.open(config.spreadsheet_name).get_worksheet(config.sheet_index)

    ## update cells
    #wks.update_acell(config.cell_to_update, value)
    #wks.update_acell(config.cell_for_date, timeStamp)


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

        # get the value we care about
        #csvValue = getCSVData()

        # send data to google
        editGoogleSheet(123, startTimeFormat)

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

