from oauth2client.file import Storage
from oauth2client.client import flow_from_clientsecrets
from oauth2client.tools import run
import httplib2

import gdata.spreadsheets.client
import gdata.spreadsheet.service
import gdata.gauth

storage = Storage("creds.dat")
credentials = storage.get()
if credentials is None or credentials.invalid:
        credentials = run(flow_from_clientsecrets("credentials.json", scope=["https://spreadsheets.google.com/feeds"]), storage)
if credentials.access_token_expired:
        credentials.refresh(httplib2.Http())

# get list of spreadsheets i am authorized to access
gd_client = gdata.spreadsheets.client.SpreadsheetsClient()
gd_client.auth_token = gdata.gauth.OAuth2TokenFromCredentials(credentials)
print gd_client.get_spreadsheets()

# access a specific spreadsheet
client = gdata.spreadsheet.service.SpreadsheetsService(additional_headers={'Authorization' : 'Bearer %s' % credentials.access_token})
entry = client.GetSpreadsheetsFeed('1CJjBsF2rVYhkrfTMZ2yBNzNXrd8t-7ex9WlrKa04f7c')
print entry.title