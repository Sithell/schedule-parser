import os.path
from libs.event import Event

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

    class Decorators:
        @classmethod
        def return_event(cls, decorated):
            def wrapper(*args, **kwargs):
                event = decorated(*args, **kwargs)
                return Event(
                    event.get('id'),
                    event.get('status'),
                    event.get('htmlLink'),
                    event.get('summary'),
                    event.get('description'),
                    event.get('location'),
                    event.get('colorId'),
                    event.get('start'),
                    event.get('end'),
                )

            return wrapper

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
        if os.path.exists(self.authorized_user_file):
            self.credentials = Credentials.from_authorized_user_file(self.authorized_user_file, self.SCOPES)

        # If there are no (valid) credentials available, let the user log in.
        if not self.credentials or not self.credentials.valid:
            if self.credentials and self.credentials.expired and self.credentials.refresh_token:
                self.credentials.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(client_secrets_file, self.SCOPES)
                self.credentials = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.authorized_user_file, 'w') as token:
                token.write(self.credentials.to_json())

        self.service = build('calendar', 'v3', credentials=self.credentials)

    @Decorators.return_event
    def insert_event(self,
                     start, end,
                     summary=None,
                     description=None,
                     location=None,
                     frequency='WEEKLY',
                     interval=1,
                     timezone='Europe/Moscow') -> Event:
        body = event = {
            'summary': summary,
            'location': location,
            'description': description,
            'start': {
              'dateTime': start,
              'timeZone': timezone,
            },
            'end': {
              'dateTime': end,
              'timeZone': timezone,
            },
            'recurrence': [
              'RRULE:FREQ={};INTERVAL={};'.format(frequency, interval)
            ],
        }
        event = self.service.events().insert(calendarId=self.calendar_id, body=body).execute()
        return event

    @Decorators.return_event
    def get_events(self) -> list[Event]:
        return self.service.events().list(
            calendarId=self.calendar_id,
            orderBy='updated',
            singleEvents=True,
            maxResults=100
        ).execute().get('items', [])

    def delete_event(self, event_id):
        self.service.events().delete(calendarId=self.calendar_id, eventId=event_id).execute()

    @staticmethod
    def format_time(time):
        return time.strftime('%Y-%m-%dT%H:%M:%S') + '+03:00'

    def __del__(self):
        self.service.close()
