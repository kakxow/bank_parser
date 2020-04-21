from dataclasses import dataclass
import datetime as dt
import locale
from pathlib import Path
import time
from typing import Iterable, List

from config import REPORT_MAIL_TEXT, REPORT_MAIL_SUBJECT
from config import ALPHA_PATH, SBER_PATH, VTB_PATH
from .outlook import Mail


def timer(fn):
    def g(*args, **kwargs):
        start = time.time()
        res = fn(*args, **kwargs)
        print(time.time() - start)
        return res
    return g


def change_cyr(txt: str) -> str:
    to_be_replaced = 'ЕНХВАРОКСМТ'
    replace_with = 'EHXBAPOKCMT'
    translation = str.maketrans(to_be_replaced, replace_with)
    result = txt.translate(translation)
    return result


def send_report(attachment_path: str, recepients: List[str]):

    report_mail = Mail(
        recepients=recepients,
        subject=REPORT_MAIL_SUBJECT,
        body=REPORT_MAIL_TEXT,
        attachments_paths=[attachment_path]
    )
    report_mail.send_mail(send=True)


def flatten(items: Iterable):
    for element in items:
        if isinstance(element, Iterable) and not isinstance(element, (str, bytes)):
            yield from flatten(element)
        else:
            yield element


@dataclass
class Paths:
    alpha: Path
    sber: Path
    sber_spec: Path
    vtb: Path


def paths_for_date(date: dt.datetime) -> Paths:
    locale.setlocale(locale.LC_ALL, 'ru-RU')

    month_name = date.strftime('%B')
    month_num = date.strftime('%m')
    year = date.strftime('%Y')

    alpha = ALPHA_PATH / year / month_num
    sber = SBER_PATH / f'{month_name} {year}'
    sber_spec = SBER_PATH / f'{sber} спец счет'
    vtb = VTB_PATH / year / f'{month_name} {year}'

    paths = Paths(alpha, sber, sber_spec, vtb)
    return paths
