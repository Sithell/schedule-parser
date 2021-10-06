from libs.google_calendar import GoogleCalendar
from libs.event import Event
import configparser
import json

# Удалить все созданные раннее мероприятия

config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

with open('events.txt', 'r', encoding='utf-8') as file:
    events = [Event(**json.loads(x)) for x in file.readlines()]

calendar = GoogleCalendar(config.get('GOOGLE', 'calendar_id'))
for event in events:
    calendar.delete_event(event.id)
    print('Deleted {}: {}'.format(event.id, event.summary))

open('events.txt', 'w').close()
