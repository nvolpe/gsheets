#from oauth2client.file import Storage
#from oauth2client.client import flow_from_clientsecrets
#from oauth2client.tools import run
#import httplib2

#import gdata.spreadsheets.client
#import gdata.spreadsheet.service
#import gdata.gauth

#storage = Storage("creds.dat")
#credentials = storage.get()
#if credentials is None or credentials.invalid:
#        credentials = run(flow_from_clientsecrets("credentials.json", scope=["https://spreadsheets.google.com/feeds"]), storage)
#if credentials.access_token_expired:
#        credentials.refresh(httplib2.Http())

## get list of spreadsheets i am authorized to access
#gd_client = gdata.spreadsheets.client.SpreadsheetsClient()
#gd_client.auth_token = gdata.gauth.OAuth2TokenFromCredentials(credentials)
#print gd_client.get_spreadsheets()

## access a specific spreadsheet
#client = gdata.spreadsheet.service.SpreadsheetsService(additional_headers={'Authorization' : 'Bearer %s' % credentials.access_token})
#entry = client.GetSpreadsheetsFeed('1CJjBsF2rVYhkrfTMZ2yBNzNXrd8t-7ex9WlrKa04f7c')
#print entry.title



import json
import webbrowser
import httplib2
from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client.tools import run
import gdata.docs.service
import gdata.spreadsheet.service
import gdata.spreadsheet.text_db
from pysqlite2 import dbapi2 as sqlite3
import pymssql
import sys
import datetime
import os
import argparse
from collections import defaultdict

storage = Storage("creds.dat")
credentials = storage.get()
if credentials is None or credentials.invalid:
  credentials = run(flow_from_clientsecrets("client_secrets.json", scope=["https://spreadsheets.google.com/feeds"]), storage)

def getGdataCredentials(client_secrets="client_secrets.json", storedCreds="creds.dat", scope=["https://spreadsheets.google.com/feeds"], force=False):
  storage = Storage(storedCreds)
  credentials = storage.get()
  if credentials is None or credentials.invalid or force:
    credentials = run(flow_from_clientsecrets(client_secrets, scope=scope), storage)
  if credentials.access_token_expired:
    credentials.refresh(httplib2.Http())
  return credentials

def getAuthorizedSpreadsheetClient(client_secrets="client_secrets.json", storedCreds="creds.dat", force=False):
  credentials = getGdataCredentials(client_secrets=client_secrets, storedCreds=storedCreds, scope=["https://spreadsheets.google.com/feeds"], force = force)
  client = gdata.spreadsheet.service.SpreadsheetsService(
    additional_headers={'Authorization' : 'Bearer %s' % credentials.access_token})
  return client


# Get authorized client
client = getAuthorizedSpreadsheetClient()
s = client.GetWorksheetsFeed(key)