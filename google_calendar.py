import os.path
from time import struct_time

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request


class GoogleCalendar:
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    calendar_id = None
    credentials = None
    service = None

    def __init__(
            self,
            calendar_id='primary',
            authorized_user_file='token.json',
            client_secrets_file='credentials.json'
    ):
        self.calendar_id = calendar_id
        self.authorized_user_file = authorized_user_file
        self.client_secrets_file = client_secrets_file

        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            self.credentials = Credentials.from_authorized_user_file('token.json', self.SCOPES)

        # If there are no (valid) credentials available, let the user log in.
        if not self.credentials or not self.credentials.valid:
            if self.credentials and self.credentials.expired and self.credentials.refresh_token:
                self.credentials.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', self.SCOPES)
                self.credentials = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(self.credentials.to_json())

        self.service = build('calendar', 'v3', credentials=self.credentials)

    def insert_event(self, start, end, summary, location, description):
        body = event = {
            'summary': summary,
            'location': location,
            'description': description,
            'start': {
              'dateTime': start,
              'timeZone': 'America/Los_Angeles',
            },
            'end': {
              'dateTime': end,
              'timeZone': 'America/Los_Angeles',
            },
            # 'recurrence': [
            #   'RRULE:FREQ=WEEKLY;COUNT=2'
            # ],
          }
        event = self.service.events().insert(calendarId=self.calendar_id, body=body).execute()
        return event

    @staticmethod
    def format_time(time):
        return time.strftime('%Y-%m-%dT%H:%M:%S') + '+03:00'

    def __del__(self):
        self.service.close()
