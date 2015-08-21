'''
configuration for google sheet tool

Author: NVolpe
Date: 8/16/2015
'''

# google account credentials json file on disk
credentials_file_on_disk = 'credentials.json'

# name of the spreadsheet we want to update
spreadsheet_name = 'Questions for CDOT meeting'

speedsheet_id = '1CJjBsF2rVYhkrfTMZ2yBNzNXrd8t-7ex9WlrKa04f7c'

# worksheet index, Zero-based numbering
sheet_index = 0 

# cell for date
cell_for_date = [4, 4] 
cell_for_date_worksheet = 'suckadick'


data_lookup =  [
    {
        'sheet_name': 'suckadick',
        'csv_location': [1, 1],
        'gsheet_location': [10, 2]
    },
    {
        'sheet_name': 'eatpussy',
        'csv_location': [1, 2],
        'gsheet_location': [11, 2]
    },
    {
        'sheet_name': 'lickclit',
        'csv_location': [2, 3],
        'gsheet_location': [12, 2]
    }
]


# google email account for sending warning email
username = 'navolpe@gmail.com'
password = 'fuckalex'

# warning email recepients
recipients = ['navolpe@gmail.com'] # comma delimited string array

# warning email subject line
subject = 'WARNING: CSV harvester failed'

# csv location on disk
csv_file_on_disk = r'C:\projects\github\personal\allston\allston-gdrive\gsheets\sheet.csv'