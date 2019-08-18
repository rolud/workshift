import datetime 
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import xlrd

SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = 'rr1g9ige1sm67rgqgv1ria19o4@group.calendar.google.com'
SHIFT = {'1':'Mattina', '2':'Pomeriggio', '3':'Notte', 'R*':'Riposo', 'NL':'Non lavoro', 
         '1/F.':'Mattina (FERIE)', '2/F.':'Pomeriggio (FERIE)', '3/F.':'Notte (FERIE)', 'R*/F.':'Riposo (FERIE)', 'NL/F.':'Non lavoro (FERIE)',
         '1/#F':'Mattina (FERIE RICHIESTE)', '2/#F':'Pomeriggio (FERIE RICHIESTE)', '3/#F':'Notte (FERIE RICHIESTE)', 'R*/#F':'Riposo (FERIE RICHIESTE)', 'NL/#F':'Non lavoro (FERIE RICHIESTE)',}
WORKTIME = {
    'Mattina':{
        'start': '07:00:00',
        'end': '15:00:00'
    },
    'Pomeriggio':{
        'start': '15:00:00',
        'end': '23:00:00'
    },
    'Notte':{
        'start': '23:00:00',
        'end': '07:00:00'
    },
}

def auth():
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

    return creds

def worksheet(path):
    events = []
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    for row in range(sheet.nrows):
        cell = sheet.cell(row,0)
        if cell.value == "SCARCELLA":
            break
    # print(row)
    for col in range(1,sheet.ncols):
        y = '2019'
        m = '08'
        d = sheet.cell(2,col).value
        shift = SHIFT[sheet.cell(row,col).value]
        # print(shift)
        if shift in WORKTIME:
            #print(col)
            # print(sheet.cell(2,col).value, 'agosto 2019 -> ', SHIFT[sheet.cell(row,col).value])
            if shift == 'Notte':
                ed = int(d) + 1
            else:
                ed = int(d)
            event = {
                'summary': shift,
                'start': {
                    'dateTime': y + '-' + m + '-' + "{:0>2d}".format(int(d)) + 'T' + WORKTIME[shift]['start'] + '+02:00',
                    'timeZone': 'Europe/Rome',
                },
                'end' : {
                    'dateTime': y + '-' + m + '-' + "{:0>2d}".format(ed) + 'T' + WORKTIME[shift]['end'] + '+02:00',
                    'timeZone': 'Europe/Rome',
                },
            }
            events.append(event)
    return events

def main():

    creds = auth()
    service = build('calendar', 'v3', credentials = creds)

    events = worksheet("Agosto 19.xlsx")

    for event in events:
        event = service.events().insert(calendarId='rr1g9ige1sm67rgqgv1ria19o4@group.calendar.google.com', body = event).execute()
        # print('------------------')
        # print(event['summary'])
        # print(event['start']['dateTime'])
        # print(event['end']['dateTime'])

    # event = {
    #     'summary': 'Pomeriggio',
    #     'start': {
    #         'dateTime': '2019-08-19T15:00:00+02:00',
    #         'timeZone': 'Europe/Rome',
    #     },
    #     'end': {
    #         'dateTime': '2019-08-19T23:00:00+02:00',
    #         'timeZone': 'Europe/Rome',
    #     },
    # }

    # event = service.events().insert(calendarId='rr1g9ige1sm67rgqgv1ria19o4@group.calendar.google.com', body = event).execute()
    # print('event created:')
    # print(event.get('htmlLink'))




if __name__ == '__main__':
    main()
