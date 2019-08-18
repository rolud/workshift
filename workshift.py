import datetime 
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/calendar']

def main():

    creds = None
    # the file toke.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # if there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            # save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials = creds)

    event = {
        'summary': 'Pomeriggio',
        'start': {
            'dateTime': '2019-08-19T15:00:00+02:00',
            'timeZone': 'Europe/Rome',
        },
        'end': {
            'dateTime': '2019-08-19T23:00:00+02:00',
            'timeZone': 'Europe/Rome',
        },
    }

    event = service.events().insert(calendarId='rr1g9ige1sm67rgqgv1ria19o4@group.calendar.google.com', body = event).execute()
    print('event created:')
    print(event.get('htmlLink'))

if __name__ == '__main__':
    main()