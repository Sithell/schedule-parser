import xlrd
import re
import json
from time import strptime
import configparser
from datetime import datetime, date, timedelta
from libs.google_calendar import GoogleCalendar

# Это либо гениально, либо очень тупо, но эта функция записывает значение в глобальную переменную
# когда аргумент передается и возвращает последнее записанное значение, когда аргумента нет
# Нужно для того, чтобы объявить переменную и сразу использовать ее значение в той же строке
# В PHP это работает из коробки: a = b = [] -> a = []
temp = None
def _(*args):
    global temp
    if len(args) == 1:
        temp = args[0]

    if len(args) > 1:
        temp = args

    return temp


def pretty_print(obj):
    print(json.dumps(obj, indent=4).encode().decode('unicode_escape'))


# Лень было искать как проверить есть ли в массиве непустые значения, поэтому велосипед
def is_empty(array):
    for item in array:
        if item:
            return False

    return True


def is_upper_week(date):
    # TODO По дате определять верхняя это неделя или нижняя
    return True

# С помощью регулярок вытаскивает из описания пары инфу:
# * Место проведения
# * Преподаватель
# * Время
# * Название предмета
# week - None, 'upper', 'lower'
# Предполагаем, что сегодня верхняя неделя
def parse_class(dayOfWeek, time, content, week=None):
    content_str = ' '.join(content)
    day = {
        'понедельник': 0,
        'вторник': 1,
        'среда': 2,
        'четверг': 3,
        'пятница': 4,
        'суббота': 5,
    }[dayOfWeek.lower()]

    location = ''
    description = ''
    teacher = ''

    # Аудитория
    if _(re.match(r'.*(ауд.? *(А?-?\d\d\d))', content_str, re.UNICODE | re.IGNORECASE)) is not None:
        location = config.get('STRINGS', 'room') + _().groups()[1]
        is_remote = False
        content_str = content_str.replace(_().groups()[0], '')

    # Дистант
    if _(re.match(r'.*(дистанционно)', content_str, re.UNICODE | re.IGNORECASE)) is not None:
        location = config.get('STRINGS', 'distant')
        is_remote = True
        content_str = content_str.replace(_().groups()[0], '')

    # Преподаватель
    if _(re.match(r'.* ([А-Яа-я\.]+ ?[А-Я][а-я]+ [А-Я]\.[А-Я]\.)', content_str, re.UNICODE)) is not None or \
            _(re.match(r'.*([А-Я][а-я]+ [А-Я]\.[А-Я]\.)', content_str, re.UNICODE)) is not None:
        teacher = _().groups()[0]
        content_str = content_str.replace(teacher, '')

    # Время начала и конца
    today = date.today()
    today = datetime(today.year, today.month, today.day)
    today += timedelta(days=day - today.weekday())

    if week is not None:
        if (is_upper_week(today) and week == 'lower') or (not is_upper_week(today) and week == 'upper'):
            today += timedelta(days=7)

        interval = 2
    else:
        interval = 1

    start = strptime(time.split('-')[0], '%H.%M')
    start = timedelta(hours=start.tm_hour, minutes=start.tm_min)
    end = strptime(time.split('-')[1], '%H.%M')
    end = timedelta(hours=end.tm_hour, minutes=end.tm_min)

    # Убрать слово "Авиамоторная"
    if _(re.match(r'.*(авиамоторная)', content_str, re.UNICODE | re.IGNORECASE)) is not None:
        content_str = content_str.replace(_().groups()[0], '')

    # Убрать слово "ауд."
    if _(re.match(r'.*( *ауд\.?)', content_str, re.UNICODE | re.IGNORECASE)) is not None:
            content_str = content_str.replace(_().groups()[0], '')

    title = ' '.join(content_str.strip().split(' '))
    if teacher:
        description += "{}: {}\n".format(config.get('STRINGS', 'teacher'), teacher)

    if config.getboolean('DEFAULT', 'add_full_text_to_description'):
        description += '\n'.join(content)

    result = {
        'title': title,
        'description': description,
        'location': location,
        'start': today + start,
        'end': today + end
    }
    event = google_calendar.insert_event(
        google_calendar.format_time(result['start']),
        google_calendar.format_time(result['end']),
        result['title'],
        result['description'],
        result['location'],
        interval=interval
    )
    print(time, event.summary)
    with open('events.txt', 'a', encoding='utf-8') as file:
        json.dump(event.asdict(), file, ensure_ascii=False)
        print(file=file)

    return result


config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

book = xlrd.open_workbook(config.get('DEFAULT', 'filename'))  # TODO Принимать путь к файлу из аргументов
sh = book.sheet_by_index(0)

col_start = config.getint('EXCEL', 'row_begin')  # Строка, с которой начинается расписание (включая название группы)
row_count = 5 * 6 * 4  # 5 пар в день, 6 дней в неделю, 4 строки на пару

schedule = {}
day = ''
time = ''
for row in range(col_start + 1, col_start + 1 + row_count):
    _day = sh.cell_value(rowx=row, colx=config.getint('EXCEL', 'col_day'))
    _time = sh.cell_value(rowx=row, colx=config.getint('EXCEL', 'col_time'))
    value = sh.cell_value(rowx=row, colx=config.getint('EXCEL', 'col_class'))

    if not (_day.isspace() or len(_day) == 0):
        day = _day

    if not (_time.isspace() or len(_time) == 0):
        time = _time

    if schedule.get(day) is None:
        schedule[day] = {}

    if schedule.get(day).get(time) is None:
        schedule[day][time] = []

    schedule[day][time].append(value.strip())

del _day, day, _time, time, value

google_calendar = GoogleCalendar(config.get('GOOGLE', 'calendar_id'))

for day in schedule.keys():
    print(day + ': ')
    for time in schedule.get(day).keys():
        lessons = schedule.get(day).get(time)
        if is_empty(lessons):  # Пары нет
            continue

        if len(lessons[0]) > 0:  # Если первая строка заполнена, то есть верхняя неделя
            parse_class(day, time, lessons[:2], week='upper')
            if not is_empty(lessons[2:]):
                parse_class(day, time, lessons[2:], week='lower')

            continue

        if is_empty(lessons[:2]) and not is_empty(lessons[2:]):  # Только нижняя неделя
            parse_class(day, time, lessons[2:], week='lower')
            continue

        parse_class(day, time, [_ for _ in lessons if _])

