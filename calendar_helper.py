from __future__ import print_function
from datetime import datetime, timedelta
import pytz 
import pickle
import os.path
import json
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def get_service():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def create_event(summary, location, start_date_time, end_date_time):
    return {
        'summary': summary,
        'location': location,
        'start': {
            'dateTime': start_date_time,
            'timeZone': 'Europe/Stockholm'

        },
        'end' : {
            'dateTime': end_date_time,
            'timeZone': 'Europe/Stockholm'
        }
    }


def get_all_events():
    min_time = datetime.now().utcnow().isoformat() + 'Z'
    events = get_service().events().list(calendarId='primary', timeMin=min_time, singleEvents=True).execute()
    return  events.get('items', [])
    
def get_mt_event(start_date):
    events = get_all_events()

    for event in events:
        if event['summary'] == 'Muay Thai' and str(start_date.isoformat()) in event['start']['dateTime']:
            return event

    return None


def delete_event(start_time):
    print('Deleteing: ' + str(start_time))
    event = get_mt_event(start_time)
    if event is not None:
        eid = event['id']
        get_service().events().delete(calendarId='primary', eventId=eid).execute()



def book_muay_thai(date_time):
    print('Booking: ' + str(date_time))
    start_time = date_time.strftime('%Y-%m-%dT%H:%M:%S')
    if get_mt_event(date_time) is None:
        end_time = (date_time + timedelta(hours=1, minutes=30)).strftime('%Y-%m-%dT%H:%M:%S')
        event = create_event('Muay Thai', 'Fighcenter', start_time, end_time)
        events = get_service().events()

        events.insert(calendarId='primary', body=event).execute()
    return

