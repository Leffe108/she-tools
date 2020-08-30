"""
This script can be used to produce a CSV file from IOF XML ResultFile from eg. UsynligO.

The CSV file contains a subset of all data in the XML file.

Written 2020 Leif Linse
Python: 3.8.5. Developed on Windows 10
"""
from xml.dom.minidom import parse, Element
from typing import List,Union
import sys
import os
import csv
import math
import datetime
import dateutil.parser

def findall(node: Element, node_name: str) -> List:
    """Find all DOM elements from node by node name"""
    return node.getElementsByTagName(node_name)

def find(node: Element, node_name: str) -> [Element, None]:
    """Find first DOM element from node by node name"""
    r = findall(node, node_name)
    if len(r) == 0:
        return None
    return r[0]
def get_inner_text(node: [Element, None]) -> str:
    """
    From an element eg. <Given>Family Name</Given>
    return the inner text 'Family Name'
    """
    if node is None or not node.hasChildNodes():
        return ''
    return node.childNodes[0].data
def to_int(s: str) -> [int, None]:
    """Convert string to int|None"""
    if s == '':
        return None
    return int(s, 10)
def to_s(value: [int, None]) -> str:
    """Convert int|None to string"""
    if value is None:
        return ''
    return str(value)
def name_to_str(name: Element) -> str:
    """
    Read Given and Family name from name Element
    and return a string.
    """
    name_parts = []
    given = find(name, 'Given')
    family = find(name, 'Family')
    if given is not None and given.hasChildNodes():
        name_parts.append(get_inner_text(given).strip())
    if family is not None and family.hasChildNodes():
        name_parts.append(get_inner_text(family).strip())
    return ' '.join(name_parts)
def human_time(time: int) -> str:
    """
    Convert time in seconds to format [hh:]:mm:ss
    where hour part is ommitted if it is zero.
    """
    h = math.floor(time / 3600)
    m = math.floor(time / 60) - h * 60
    s = time - h * 3600 - m * 60

    r = ''
    if h > 0:
        r += str.zfill(h, 2) + ':'
    r += str.zfill(m, 2) + ':' + str.zfill(s, 2)

    return r
def iso_time_to_date_time(iso_time: str) -> Union[datetime.datetime,None]:
    """
    Convert ISO 8601 time to datetime object
    """
    print('iso time: ' + iso_time)
    if iso_time == '':
        return None
    return dateutil.parser.isoparse(iso_time)


def print_class_info(class_result: Element) -> None:
    """
    Print class id, name and course name to console
    """
    cls = find(class_result, 'Class')
    class_id = get_inner_text(find(cls, 'Id'))
    class_name = get_inner_text(find(cls, 'Name'))
    course = find(class_result, 'Course')
    course_name = get_inner_text(find(course, 'Name'))
    print('Class id: ' + class_id)
    print('Class name: ' + class_name)
    print('Course name: ' + course_name)


def parse_xml(file_name: str) -> List:
    """
    Parse IOF XML ResultList file to Python data structure
    """
    dom = parse(file_name)

    class_result = dom.getElementsByTagName('ClassResult')[0]
    print_class_info(class_result)

    persons = findall(class_result, 'PersonResult')

    data = []
    for p_result in persons:
        p = find(p_result, 'Person')
        r = {}

        r['name'] = name_to_str(find(p, 'Name'))
        r['team'] = ''
        org = find(p_result, 'Organisation')
        if org is not None:
            print(r['name'] + ' has org')
            r['team'] = get_inner_text(find(org, 'Name')).strip()
        result = find(p_result, 'Result')
        r['time'] = to_int(get_inner_text(find(result, 'Time')))
        r['start_datetime'] = iso_time_to_date_time(get_inner_text(find(result, 'StartTime')))
        r['finish_datetime'] = iso_time_to_date_time(get_inner_text(find(result, 'FinishTime')))
        r['position'] = to_int(get_inner_text(find(result, 'Position')))
        r['status'] = get_inner_text(find(result, 'Status'))
        r['split_times'] = []
        for split_time in findall(result, 'SplitTime'):
            r['split_times'].append({
                'time': to_int(get_inner_text(find(split_time, 'Time'))),
                'control_code': to_int(get_inner_text(find(split_time, 'ControlCode'))),
            })
        r['split_times'].sort(key=lambda x: x['time'])

        data.append(r)

    max_pos = max(x['position'] if x['position'] is not None else 0 for x in data)
    data.sort(key=lambda x: x['position'] if x['position'] is not None else max_pos + 1)
    return data


def format_split_times(split_times: List) -> str:
    """
    Format split times list to string
    <control code>: time, <control code>: time, ...
    """
    return ', '.join([to_s(st['control_code']) + ': ' + to_s(st['time']) for st in split_times])


def control_str(split_times: List) -> str:
    """
    Create comma separated list of visited control codes from
    split_times data. (ommits time)
    Format:
    <control code>, <control code>, ...
    """
    return ', '.join([to_s(st['control_code']) for st in split_times])


def csv_out(file_name: str, data: List) -> None:
    """
    Produces CSV file from parse_csv data

    Start date/time is in CEST (UTC + 2)
    """
    f = open(file_name, 'w', newline='')
    w = csv.writer(f, dialect='excel-variant')

    w.writerow([
        'Position',
        'Name',
        'Team',
        'Time',
        'Status',
        'Controls',
        'Split Times',
        'Start Date',
        'Start Time',
    ])
    for person in data:
        start_dt = person['start_datetime']
        if start_dt is not None:
            start_dt = start_dt.astimezone(datetime.timezone(datetime.timedelta(hours=2), 'CEST'))
        w.writerow([
            to_s(person['position']),
            person['name'],
            person['team'],
            to_s(person['time']),
            person['status'],
            control_str(person['split_times']),
            format_split_times(person['split_times']),
            start_dt.date().isoformat() if start_dt is not None else '',
            start_dt.time().isoformat(timespec='seconds') if start_dt is not None else '',
        ])


# -- Main ---
en_csv = len(sys.argv) >= 2 and sys.argv[1] == '--en'
if en_csv:
    files = sys.argv[2:]
else:
    files = sys.argv[1:]

if len(files) < 1:
    print('Extracts results from IOF 3.0 XML file to CSV')
    print()
    print(sys.argv[0] + ' [--en] <file1.xml> [<file2.xml> ... ]')
    print('Flag --en: Changes CSV dialect to match Excel for English. By default it matches')
    print('           Excel for most other latin languages like German, French, Swedish etc.')
    print()
    print('Provide one or more file names to IOF 3.0 XML files to read.')
    print()
    print('Output:')
    print('  One maching .csv file for each input XML file')
    sys.exit(1)

# Register CSV variant
if en_csv:
    # Using , as separator when 'en' parameter is included as arg 2
    csv.register_dialect('excel-variant', delimiter=',', quoting=csv.QUOTE_NONNUMERIC)
else:
    # Using ; as separator which is used by Excel in most non-English latin languages
    csv.register_dialect('excel-variant', delimiter=';', quoting=csv.QUOTE_NONNUMERIC)

for input_file_name in files:
    # parse XML
    data = parse_xml(input_file_name)
    # output to CSV
    csv_file_name = os.path.splitext(input_file_name)[0] + '.csv'
    csv_out(csv_file_name, data)
    print('Saved ' + csv_file_name)
