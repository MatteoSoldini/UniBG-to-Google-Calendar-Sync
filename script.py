from __future__ import print_function
from datetime import date, datetime, timedelta
import pickle
import os.path
import requests
import sys
from icalendar import Calendar
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

calendarId = '' #INSERT YOUR CALENDAR ID HERE

class PostToGoogleCalendar:
    def __init__(self):
        self.creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                self.creds = flow.run_local_server()
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('calendar', 'v3', credentials=self.creds)

    def get_events(self):
        now = getReferenceDate().isoformat() + 'T00:00:00Z'
        events_result = self.service.events().list(calendarId=calendarId, timeMin=now,
                                                   maxResults=500, singleEvents=True,
                                                   orderBy='startTime').execute()

        """for event in events_result.get('items'):
            print(event.get('summary'), event.get('start').get('dateTime'))"""
        return events_result.get('items')

    def create_event(self, new_event):
        if not self.already_exists(new_event):
            event = self.service.events().insert(calendarId=calendarId, body=new_event).execute()
            return event.get('htmlLink')
            print('Event Does Not Exists')
        else:
            print('Event Already Exists')

    def already_exists(self, new_event):
        events = self.get_date_events(new_event['start']['dateTime'], self.get_events())
        event_list = [new_event['summary'] for new_event in events]
        if new_event['summary'] not in event_list:
            return False
        else:
            return True

    def get_date_events(self, date, events):
        lst = []
        date = date
        for event in events:
            #print(date, event.get('start').get('dateTime'))
            if event.get('start').get('dateTime'):
                d1 = event.get('start').get('dateTime')
                if d1 == date:
                    lst.append(event)
        return lst

def getReferenceDate():
    now = date.today()
    return now - timedelta(days=now.weekday())

def printCalendarList(service):
   # Call the Calendar API
   print('Getting list of calendars')
   calendars_result = service.calendarList().list().execute()

   calendars = calendars_result.get('items', [])

   if not calendars:
       print('No calendars found.')
   for calendar in calendars:
       summary = calendar['summary']
       id = calendar['id']
       primary = "Primary" if calendar.get('primary') else ""
       print("%s\t%s\t%s" % (summary, id, primary))

def main():
    reference_date = getReferenceDate()

    for w in range(0, 2):
        url = 'https://logistica.unibg.it/PortaleStudenti/ec_download_ical_grid.php?view=easycourse&form-type=corso&include=corso&txtcurr=1+anno+-+PERCORSO+COMUNE&anno=' + str(reference_date.year) + '&scuola=ScuoladiIngegneria&corso=21-270&anno2%5B%5D=PDS0-2012%7C1&visualizzazione_orario=cal&date=' + (reference_date + timedelta(weeks=w)).strftime('%d-%m-%Y') + '&periodo_didattico=&_lang=it&list=0&week_grid_type=-1&ar_codes_=&ar_select_=&col_cells=0&empty_box=0&only_grid=0&highlighted_date=0&all_events=0&faculty_group=0&_lang=it&ar_codes_=ECCCLENGB1%20A|EC65741|EC41298%20ES%20B|EC41298%20ES%20A|EC41298%20ES%20C|EC41287%20ES%20B|EC41287%20ES%20Info|EC41294%20ES%20A|EC41294%20ES%20B|EC41294%20ES%20C|EC65744%20ES%20A|EC65744%20ES%20B|EC65747|EC65744|EC65738&ar_select_=true|true|true|true|true|true|true|true|true|true|true|true|true|true|true&txtaa=2020/2021&txtcorso=INGEGNERIA%20INFORMATICA%20(Laurea)&txtanno=&docente=&attivita=&txtdocente=&txtattivita='
        r = requests.get(url, allow_redirects=True)

        open('calendar.ics', 'wb').write(r.content)

        g = open('calendar.ics','rb')
        gcal = Calendar.from_ical(g.read())
        for component in gcal.walk():
            if component.name == "VEVENT":
                s_dt = component.decoded('dtstart')
                e_dt = component.decoded('dtend')
                event = {
                      'summary': component.get('summary'),
                      'start': {
                        'dateTime': s_dt.strftime('%Y-%m-%dT%H:%M:%S') + '+02:00',
                        'timeZone': 'Europe/Rome',
                      },
                      'end': {
                        'dateTime': e_dt.strftime('%Y-%m-%dT%H:%M:%S') + '+02:00',
                        'timeZone': 'Europe/Rome',
                      },
                      'reminders': {
                        'useDefault': False,
                        'overrides': [
                          {'method': 'popup', 'minutes': 30},
                        ],
                      },
                }

                g_calendar = PostToGoogleCalendar()
                g_calendar.create_event(event)

                #event = service.events().insert(calendarId='dnfoduh0ci32vab2nj6tpv9218@group.calendar.google.com', body=event).execute()
                #print('Adding ' + component.get('summary'))
        g.close()

if __name__ == '__main__':
    main()
