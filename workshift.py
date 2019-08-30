import datetime 
import pickle
import os.path
import sys

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

def worksheet(input, month, year):
    events = []
    workbook = xlrd.open_workbook(input)
    sheet = workbook.sheet_by_index(0)
    daycell = -1
    for row in range(sheet.nrows):
        cell = sheet.cell(row,0)
        if (cell.ctype == 1):
            if cell.value.upper() == 'COGNOME':
                daycell = row
            if cell.value.upper() == "SCARCELLA":
                break
    if daycell < 0: 
        return
    print(daycell, row)
    for col in range(1,sheet.ncols):
        day = int(sheet.cell(daycell,col).value)
        shift = SHIFT[sheet.cell(row,col).value]
        # print(shift)
        if shift in WORKTIME:
            #print(col)
            # print(sheet.cell(2,col).value, 'agosto 2019 -> ', SHIFT[sheet.cell(row,col).value])
            if shift == 'Notte':
                if ((month == 1 or month == 3 or month == 5 or month == 7 or month == 8 or month == 10) and day == 31):
                    end_day = 1
                    end_month = month + 1
                    end_year = year
                elif (month == 12 and day == 31):
                    end_day = 1
                    end_month = 1
                    end_year = year + 1
                elif ((month == 4 or month == 6 or month == 9 or month == 11) and day == 30):
                    end_day = 1
                    end_month = month + 1
                    end_year = year
                elif (month == 2):
                    if (year % 4 == 0 and day == 29):
                        end_day = 1
                        end_month = month + 1
                        end_year = year
                    elif (year % 4 != 0 and day == 28):
                        end_day = 1
                        end_month = month + 1
                        end_year = year
                else:
                    end_day = day + 1
                    end_month = month
                    end_year = year
            else:
                end_day = day
                end_month = month
                end_year = year
            print("{:0>2d}".format(day),"{:0>2d}".format(month),year)
            event = {
                'summary': shift,
                'start': {
                    'dateTime': str(year) + '-' + "{:0>2d}".format(month) + '-' + "{:0>2d}".format(day) + 'T' + WORKTIME[shift]['start'] + '+02:00',
                    'timeZone': 'Europe/Rome',
                },
                'end' : {
                    'dateTime': str(end_year) + '-' + "{:0>2d}".format(end_month) + '-' + "{:0>2d}".format(end_day) + 'T' + WORKTIME[shift]['end'] + '+02:00',
                    'timeZone': 'Europe/Rome',
                },
            }
            events.append(event)
    return events

def main():

    if (len(sys.argv) != 6):
        print('Error! Usage: python workshift.py <input_file> -m <mm> -y <yyyy>')
        return
    script, inputFile, mc, month, yc, year = sys.argv
    if (mc != '-m' or yc != '-y' or not month.isnumeric() or not year.isnumeric()):
        print('Error!! Usage: python workshift.py <input_file> -m <mm> -y <yyyy>')
        return
    if (int(month) < 1 or int(month) > 12):
        print('Error! month should be between 1 and 12')
        return
    if (int(year) < 2000 or int(year) > 2030):
        print('Error! Year should be between 2000 and 2030')
        return
    # print(sys.argv)

    # creds = auth()
    # service = build('calendar', 'v3', credentials = creds)

    events = worksheet(inputFile, int(month), int(year))

    # for event in events:
    #    event = service.events().insert(calendarId='rr1g9ige1sm67rgqgv1ria19o4@group.calendar.google.com', body = event).execute()
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
