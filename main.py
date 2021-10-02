import xlrd
import re
import json
from time import strftime, strptime
import configparser


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


# С помощью регулярок вытаскивает из описания пары инфу:
# * Место проведения
# * Преподаватель
# * Время
# * Название предмета
def parse_class(day, time, content):
    content_str = ' '.join(content)

    location = ''
    title = ''
    description = ''
    teacher = ''
    start = ''
    end = ''
    is_remote = False

    # Аудитория
    if _(re.match(r'.*(ауд.? *(А?-?\d\d\d))', content_str, re.UNICODE | re.IGNORECASE)) is not None:
        location = _().groups()[1]
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
    start = strftime('%H:%M', strptime(time.split('-')[0], '%H.%M'))
    end = strftime('%H:%M', strptime(time.split('-')[1], '%H.%M'))

    # Убрать слово "Авиамоторная"
    if _(re.match(r'.*(авиамоторная)', content_str, re.UNICODE | re.IGNORECASE)) is not None:
        content_str = content_str.replace(_().groups()[0], '')

    title = ' '.join(content_str.strip().split(' '))
    description = "{}: {}\n\n{}".format(config.get('STRINGS', 'teacher'), teacher, '', '\n'.join(content))
    result = {
        'title': title,
        'description': description,
        'location': location,
        'start': start,
        'end': end
    }

    pretty_print(result)
    return result


config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

book = xlrd.open_workbook(config.get('DEFAULT', 'filename'))  # TODO Принимать путь к файлу из аргументов
sh = book.sheet_by_index(0)

col_start = config.get('EXCEL', 'row_begin')  # Строка, с которой начинается расписание (включая название группы)
row_count = 5 * 6 * 4  # 5 пар в день, 6 дней в неделю, 4 строки на пару

schedule = {}
day = ''
time = ''
for row in range(col_start + 1, col_start + 1 + row_count):
    _day = sh.cell_value(rowx=row, colx=config.get('EXCEL', 'col_day'))
    _time = sh.cell_value(rowx=row, colx=config.get('EXCEL', 'col_time'))
    value = sh.cell_value(rowx=row, colx=config.get('EXCEL', 'col_class'))

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

for day in schedule.keys():
    for time in schedule.get(day).keys():
        lessons = schedule.get(day).get(time)
        if is_empty(lessons):  # Пары нет
            continue

        if len(lessons[0]) > 0:  # Если первая строка заполнена, то сложный кейс
            # TODO Разделение на верхнюю и нижнюю неделю
            parse_class(day, time, lessons[:2])
            if not is_empty(lessons[2:]):
                parse_class(day, time, lessons[2:])

            continue

        parse_class(day, time, [_ for _ in lessons if _])
