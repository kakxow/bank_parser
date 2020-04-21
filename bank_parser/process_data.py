import csv
import datetime as dt
import re
from typing import Set, List, Tuple

from config import STRICT_PATTERN
from pin_code_map import pin_code_map
from .utils import change_cyr


def process_result(data: Set[str]):
    result: Set[Tuple[str, str, str, str]] = set()
    exceptions: List[str] = []

    date_pattern_s = r'(\d{1,2}[.]\d{2}[.]((\d{2}\b)|(\d{4}\b)))'
    date_pattern_l = rf'инк от {date_pattern_s}'

    compiled_date_pattern_s = re.compile(date_pattern_s, re.I)
    compiled_date_pattern_l = re.compile(date_pattern_l, re.I)

    for el in data:
        # Skip Vozrozhdenie bank.  they do not specify event sources.
        if '062-2016' in el:
            continue
        # Long for alpha, short for other.
        date_matches = compiled_date_pattern_l.findall(el) or \
            compiled_date_pattern_s.findall(el)

        try:
            sap_code = find_sap_code(el)
            event_name, event_sum = get_event(el)
            date = max(
                date_matches,
                key=lambda x: dt.datetime(
                    int(x[0][6:]) if int(x[0][6:]) > 2000 else int(x[0][6:]) + 2000,
                    int(x[0][3:5]),
                    int(x[0][:2])
                )
            )[0]
            if int(date[-2:]) < 20:
                print(date)
        except (AttributeError, ValueError) as e:
            exceptions.append(el)
        else:
            result.add((sap_code, date, event_name, event_sum))

    today = dt.datetime.today().date()
    report_file = rf'C:\Max\report_{today}.csv'
    with open(report_file, 'w', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(('Код ТТ', 'Дата инкассации', 'Событие', 'Сумма'))
        writer.writerows(result)
    print(*exceptions, sep='\n')
    return report_file


def find_sap_code(s: str) -> str:
    """
    Raises AttributeError if sap_code is None.
    """
    # Delete spaces between letters and digits in SAP code.
    s = re.sub(r'((?<=[A-Za-z]{2})|(?<=[A-Za-z]))\s', '', s)

    sap_code_pattern = r'([A-Z]\d{3}\b)|([A-Z]{2}\d{2}\b)'

    sap_code_match = re.search(sap_code_pattern, s, re.I) \
        or re.search(sap_code_pattern, change_cyr(s), re.I)
    # Check for AlphaBanks's alternative object_codes.
    if 'сумк' in s and not sap_code_match:
        sap_code_match = re.match(r'\w+', s, re.I)

    sap_code = sap_code_match.group()
    result = pin_code_map.get(sap_code, sap_code)
    return result


def get_event(s: str) -> Tuple[str, str]:
    """
    Raises AttributeError if event is None.
    Raises ValueError is word in event is not in the events.
    """
    event_match = re.search(STRICT_PATTERN, s, re.I)

    event_name_draft = re.match(r'^\w+', event_match.group()).group()
    event_sum = re.search(r'\d+', event_match.group()).group()
    events = {
        'нед': 'недостача',
        'изл': 'излишек',
        'сом': 'сомнительная',
        'фал': 'фальшивая',
        'расхожд': 'расхождение',
        'неплатеж': 'неплетежеспособная',
    }
    for short, event_name in events.items():
        if short in event_name_draft.lower():
            return event_name, event_sum
    raise ValueError('Wrong event!')
