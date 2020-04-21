from pathlib import Path

ALPHA_PATH = Path(r'\\fileserver\Архив выписок банка\Альфа Банк\2020\рубли')
# Structure - /2020/01, /2019/01 Январь
VTB_PATH = Path(r'\\fileserver\Архив выписок банка\ВТБ\Ритейл')
# Structure - /2020/Март 2020, /2019/01 Январь
SBER_PATH = Path(r'\\fileserver\Архив выписок банка\Сбербанк\Евросеть-Ритейл')
# Structure - /2020/Январь 2020, /2020/Январь 2020 спец счет, /Январь 2020

PATTERNS = (
    r'\bизл.*?[1-9]\d+',
    r'\bсом.*?[1-9]\d+',
    r'\bнед.*?[1-9]\d+',
    r'\bфал.*?[1-9]\d+',
)
PATTERN = '|'.join(f'(.*{ptrn})' for ptrn in PATTERNS)
STRICT_PATTERN = '|'.join(f'({ptrn})' for ptrn in PATTERNS)

r'^(.*\bизл.*[1-9]\d+)|(.*\bсом.*[1-9]\d+)|(.*\bнед.*[1-9]\d+)|(.*\bфал.*[1-9]\d+)'

REPORT_MAIL_TEXT = (
    '<p>Привет!</p>'
    '<p>Во вложении отчёт по недостачам, излишкам и '
    'сомнительным за вчерашний день.</p>'
)
REPORT_MAIL_SUBJECT = 'Недостачи, излишки, сомнительные'
