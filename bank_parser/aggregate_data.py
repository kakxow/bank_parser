import datetime as dt
from functools import partial
import os
from typing import Set

from config import PATTERN
from .mail_parser import save_registries
from . import parsers
from .registry_parser import parse_registries
from .utils import flatten, paths_for_date


def aggregate_data(only_today: bool, save_location: str) -> Set[str]:
    today = dt.datetime.today()
    ts = today.timestamp()
    allowed_diff = 12 * 60 * 60  # 12 hours in seconds.
    if today.weekday == 0:
        # Include weekend files, if today is monday.
        allowed_diff *= 5
    filter_today = partial(creation_time_filter, only_today, ts, allowed_diff)

    paths = paths_for_date(today)

    alpha = (parsers.alpha(f.path, PATTERN) for f in os.scandir(paths.alpha)
             if filter_today(f))
    sber = (parsers.sber(f.path, PATTERN) for f in os.scandir(paths.sber)
            if filter_today(f))
    sber_spec = (parsers.sber(f.path, PATTERN) for f in os.scandir(paths.sber_spec)
                 if filter_today(f))
    vtb = (parsers.vtb(f.path, PATTERN) for f in os.scandir(paths.vtb)
           if filter_today(f))

    banks = (alpha, sber, sber_spec, vtb)

    print('Parsing bank statements')
    result = set(flatten(banks)) | data_from_mail(only_today, save_location)

    if __debug__:
        formatted_result = set(f'{line}\n' for line in result)
        with open('new_file.txt', 'w') as f:
            f.writelines(formatted_result)

    return result


def data_from_mail(only_today: bool, save_location: str) -> Set[str]:
    print('Saving e-mails')
    save_registries(only_today, save_location)
    print('Parsing e-mails')
    return parse_registries(save_location)


def creation_time_filter(
    only_today: bool,
    ts: float,
    allowed_difference: int,
    file: os.DirEntry
) -> bool:
    if only_today:
        diff = ts - file.stat().st_ctime
        return diff < allowed_difference
    return True
