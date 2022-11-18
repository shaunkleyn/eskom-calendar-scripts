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
import sys
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
calendar_name = config['GoogleCalendar']['calendar_name']

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
            if not os.path.exists('credentials.json'):
                print ('Please download your credentials file from Google Cloud and place it in the same directory as this script')
                sys.exit()
                
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        print('Starting...')
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        page_token = None
        csvEventDates = []
        calendarEventDates = []
        calendarEvents = {}
        calendar_id = config['GoogleCalendar'].get('calendar_id')

        # Find the load shedding calendar ID
        if calendar_name != '' and (calendar_id is None or calendar_id == ''):
            print('Finding calendar ID for ' + calendar_name)
            page_token = None
            while True:
                calendar_list = service.calendarList().list(pageToken=page_token).execute()
                for calendar_list_entry in calendar_list['items']:
                    if calendar_list_entry['summary'] == calendar_name:
                        calendar_id = calendar_list_entry['id']
                        print('Using calendar: ' + calendar_name)

                        # Add the ID to the config
                        config["GoogleCalendar"].update({"calendar_id":calendar_list_entry['id']})

                        # Save changes to the config
                        with open("configuration.ini","w") as file_object:
                            config.write(file_object)

                        break
                page_token = calendar_list.get('nextPageToken')
                if not page_token:
                    break
        else:
            if calendar_id == '':
                calendar_id = 'primary'
                print('Using primary calendar as default')

        # Get all future events to prevent inserting duplicate calendar entries
        if calendar_id != '':
            df = pd.read_csv(url, delimiter = ',', names = ['area_name', 'start', 'finsh', 'stage', 'source'])
            filter1 = df[df.area_name==area_name]
            events_result = service.events().list(calendarId=calendar_id, timeMin=now,
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
            print('Looking for new events to add...')
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

                        event = service.events().insert(calendarId=calendar_id, body=event).execute()
                        calendarEventDates.append(pd.to_datetime(row['start']))
                        print('Event created: %s' % (event.get('htmlLink')))

            # Loop though calendar events and remove the events that are no longer in the csv events list
            print('Looking for events that should be deleted...')
            for key, value in calendarEvents.items():
                if pd.to_datetime(key) not in csvEventDates:
                    print('Event deleted for ' + key)
                    service.events().delete(calendarId=calendar_id, eventId=value).execute()

        print('Complete!')
    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()
# [END calendar_quickstart]