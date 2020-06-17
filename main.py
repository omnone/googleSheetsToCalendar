from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/calendar']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1Y9_6T0J97OvNPozKt1aNTlYR-QD_uMLVNfB6yxQywMs'
SAMPLE_RANGE_NAME = 'Εξεταστική!A2:E'

################################################################################################################
def get_credentials():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds

################################################################################################################
def create_events(creds, events):
    service = build('calendar', 'v3', credentials=creds)

    for row in events:
        try:
            start_time = row[3].split('-')[0]
            end_time = row[3].split('-')[1]
        except:
            continue

        event = {
            'summary': f'{row[1]}',
            'location': 'sample_location',
            'description': 'sample_descr',
            'start': {
                'dateTime': f'{row[2]}T{start_time}:00',
                'timeZone': 'Europe/Athens',
            },
            'end': {
                'dateTime': f'{row[2]}T{end_time}:00',
                'timeZone': 'Europe/Athens',
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 24 * 60},
                    {'method': 'popup', 'minutes': 30},
                ],
            },
        }
        
        # save event to your google calendar 
        event = service.events().insert(calendarId='primary', body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))

################################################################################################################
def get_sheet_data(creds):
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        for row in values:
            if len(row) > 0:
                print(row)

    return values

################################################################################################################
def main():
    events = get_sheet_data(get_credentials())
    create_events(get_credentials(), events)

################################################################################################################

if __name__ == '__main__':
    main()
