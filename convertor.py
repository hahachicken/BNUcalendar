import xlrd
import re
from icalendar import Event
from datetime import date, datetime, time, timedelta, timezone
from uuid import uuid1
from dateutil.relativedelta import relativedelta
from icalendar import Calendar, Event

def extract(filename):
    wb = xlrd.open_workbook(filename=filename)#打开文件
    sheet = wb.sheet_by_index(0)
    n_class = int(sheet.cell(5,0).value.split(' ')[0].split('：')[1])
    return [ [sheet.cell(row_n, i).value for i in [0,5]] for row_n in range(7, 7+n_class) ]

def CNweekday_to_digit(char):
    weekday = '一二三四五六日'
    for i in range(1,7):
        if char == weekday[i]:
            return i
    return 0


def format_info(info):
    '''
    info : ['[BQ10029451]计算机网络', '1-16周 一[1-2] 西操场(200)']
    '''
    index = info[0].find(']')
    class_name = info[0][index+1:]

    for lecture in info[1].split(','):

        week_info, day_info, place = tuple(lecture.split())

        #week_info[0]: start week, week_info[1]: end week
        week_info = re.findall(r'(\d+)-(\d+)周', week_info, re.S | re.M)[0]
        week_info = [int(i) for i in week_info]

        #day_info[0]: 周几，day_info[1]: start, day_info[2]: end
        day_info = list(re.findall(r'(.*?)\[(\d+)-(\d+)\]', day_info, re.S | re.M)[0])
        day_info[0] = CNweekday_to_digit(day_info[0])
        day_info = [int(i) for i in day_info]

        #place:
        try:
            place = re.findall(r'(.*?)\(.*', place, re.S | re.M)[0]
        except IndexError:
            pass
        
    return class_name, week_info, day_info, place

def creat_event(info, date_start):
    '''
    info[0] = '形势与政策Ⅲ'
    info[1] = [6, 8]
    info[2] = [3, 9, 10]
    info[3] = '七103'

    ~ 6-8周 三[9-10] 七103(145)
    '''
    time_start = {
    1: time(8,0),   2: time(8,55),  3: time(10,0),  4: time(10,55),
    5: time(13,30), 6: time(14,25), 7: time(15,30), 8: time(16,25),
    9: time(18,0), 10: time(18,55),11: time(19,50),12: time(20,45)
    }

    ev_start_date = date_start + relativedelta(weeks= info[1][0] - 1, weekday = info[2][0])
    ev_start_datetime = datetime.combine(ev_start_date, time_start[info[2][1]])  #上课时间
    ev_end_datetime = datetime.combine(ev_start_date, time_start[info[2][2]]) + relativedelta(minutes=45)   #下课时间
    ev_count = info[1][1] - info[1][0] + 1
    ev_interval = 1

    # generate event
    event = Event()
    event.add('uid', str(uuid1()) + '@CQUT')
    event.add('summary', info[0])
    event.add('dtstamp', datetime.now())
    event.add('dtstart', ev_start_datetime)
    event.add('dtend', ev_end_datetime)
    event.add('location', info[3])
    event.add('rrule', {'freq': 'weekly', 'interval': ev_interval, 'count': ev_count})
    return event

def process_CAL(filename, date_start = date(2021, 2, 21)):
    cal = Calendar()
    cal['version'] = '2.0'
    cal['prodid'] = '-//CQUT//Syllabus//CN'  # *mandatory elements* where the prodid can be changed, see RFC 5445

    for info in extract(filename):
        course = creat_event(format_info(info), date_start)
        cal.add_component(course)

    with open('output.ics', 'w+', encoding='utf-8') as file:
        file.write(cal.to_ical().decode('utf-8'.replace('\r\n', '\n').strip()))

#by defalut, using starting date of sp2021
process_CAL('test.xls')