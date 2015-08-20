'''
configuration for google sheet tool

Author: NVolpe
Date: 8/16/2015
'''

# google account credentials json file on disk
credentials_file_on_disk = 'credentials.json'

# name of the spreadsheet we want to update
spreadsheet_name = 'Questions for CDOT meeting'

# worksheet index, Zero-based numbering
sheet_index = 0 

# google sheet cells to update, Alphanumeric ie: B1
cell_to_update = 'E5' 
cell_for_date = 'F5'

# csv cell location for the data
row = 1
column = 1

# google email account for sending warning email
username = 'navolpe@gmail.com'
password = 'fuckalex'

# warning email recepients
recipients = ['navolpe@gmail.com'] # comma delimited string array

# warning email subject line
subject = 'WARNING: CSV harvester failed'

# csv location on disk
csv_file_on_disk = r'C:\projects\github\personal\allston\allston-gdrive\allston-gdrive\sheet.csv'