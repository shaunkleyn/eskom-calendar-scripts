# Copyright 2018 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# [START calendar_quickstart]
from __future__ import print_function
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime
import os.path
import pandas as pd
import configparser

# You need to enable Google Calendar API in a Google Cloud project: see this 
#       https://developers.google.com/calendar/api/quickstart/python


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']
config = configparser.ConfigParser()
config.read('configuration.ini')

url = config['EskomCalendar']['csv_url']
area_name = config['EskomCalendar']['area_name']
calendar_name = config['GoogleCalendar']['name']

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        loadSheddingCalendarId = ''
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        page_token = None
        csvEventDates = []
        calendarEventDates = []
        calendarEvents = {}

        # Find the load shedding calendar ID
        if calendar_name != '':
            page_token = None
            while True:
                calendar_list = service.calendarList().list(pageToken=page_token).execute()
                for calendar_list_entry in calendar_list['items']:
                    if calendar_list_entry['summary'] == calendar_name:
                        loadSheddingCalendarId = calendar_list_entry['id']
                        print('Using calendar: ' + calendar_name)
                        break
                page_token = calendar_list.get('nextPageToken')
                if not page_token:
                    break
        else:
            loadSheddingCalendarId = 'primary'
            print('Using primary calendar as default')

        # Get all future events to prevent inserting duplicate calendar entries
        if loadSheddingCalendarId != '':
            df = pd.read_csv(url, delimiter = ',', names = ['area_name', 'start', 'finsh', 'stage', 'source'])
            filter1 = df[df.area_name==area_name]
            events_result = service.events().list(calendarId=loadSheddingCalendarId, timeMin=now,
                                                singleEvents=True,
                                                orderBy='startTime').execute()
            events = events_result.get('items', [])
            if events:
                for event in events:
                    calendarEventDate = event['start'].get('dateTime', event['start'].get('date'))
                    date = pd.to_datetime(calendarEventDate)
                    calendarEventDates.append(date)
                    calendarEvents[calendarEventDate] = event['id']

            # loop through CSV items and create future events if it doesn't already exist
            for index, row in filter1.iterrows():
                csvEventDate = pd.to_datetime(row['start'])
                csvEventDates.append(csvEventDate)
                if csvEventDate not in calendarEventDates:
                    # Only create future events
                    if pd.to_datetime(row['start']).tz_localize(None) > pd.Timestamp.now():
                        event = { 'summary': 'Stage ' + row['stage'],
                            'location': row['area_name'],
                            'description': row['source'],
                            'start': { 'dateTime': row['start'] },
                            'end': { 'dateTime': row['finsh'] },
                            'reminders': { 'useDefault': False } }

                        event = service.events().insert(calendarId=loadSheddingCalendarId, body=event).execute()
                        calendarEventDates.append(pd.to_datetime(row['start']))
                        print('Event created: %s' % (event.get('htmlLink')))

            # Loop though calendar events and remove the events that are no longer in the csv events list
            for key, value in calendarEvents.items():
                if pd.to_datetime(key) not in csvEventDates:
                    print('Event deleted for ' + key)
                    service.events().delete(calendarId=loadSheddingCalendarId, eventId=value).execute()

    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()
# [END calendar_quickstart]