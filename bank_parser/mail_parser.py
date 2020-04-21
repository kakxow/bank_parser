import datetime as dt
import os
from pathlib import Path
import shutil
import zipfile

import win32com.client as win32


def save_attachments(date_start: dt.datetime, date_end: dt.datetime, save_location: str, ext_pattern: str = ''):
    """
    Saves attachments from specific subfolder in MS Outlook.

    Parameters
    ----------
    start_date
        Date to start search from.
    end_date
        Date to stop search where.
    save_location
        Where to save attachments.

    Returns
    -------
    None

    """
    save_path = Path(save_location)
    start_date = date_start.strftime('%d/%m/%Y')
    end_date = date_end.strftime('%d/%m/%Y')

    sender_ban_list = (
        'it_acq@sberbank.ru',
        'evyudin@sberbank.ru',
        'vasaparkina@sberbank.ru',
        'naavoronina@sberbank.ru',
        'noreply_vyp@qiwi.com',
        'Kartcentr@psbank.ru',
    )
    attachments_ban_list = ('pdf', 'gif', 'jpg', 'png', 'tif', 'ocx', 'htm', 'msg')

    sender_ban_list_in_dasl = \
        ' AND '.join([f"urn:schemas:httpmail:fromemail <> '{item}'" for item in sender_ban_list])
    filter_ = (
        "@SQL="
        f"{sender_ban_list_in_dasl} "
        f"AND (urn:schemas:httpmail:date >= '{start_date} 12:00 AM') "
        f"AND (urn:schemas:httpmail:date <= '{end_date} 12:00 AM') "
        "AND urn:schemas:httpmail:hasattachment = 1"
    )
    if ext_pattern:
        filter_ += f' AND ({ext_pattern})'
    folder_accountants = "Бухгалтерия (Москва): Реестры эквайринг"
    folder_inbox = "Входящие"

    outlook = win32.Dispatch('outlook.application')
    namespace = outlook.GetNamespace("MAPI")
    folder = namespace.Folders(folder_accountants).Folders(folder_inbox)
    items = folder.Items.restrict(filter_)

    for index, item in enumerate(items):
        for attachment in item.Attachments:
            file_name = attachment.filename
            if file_name[-3:].lower() in attachments_ban_list:
                continue
            indexed_file_name = str(index) + file_name
            attachment.SaveAsFile(save_path / indexed_file_name)


def save_registries(only_today: bool, save_location, ext_pattern: str = ''):
    today = dt.datetime.today()
    if only_today:
        date_start = today
        date_end = today + dt.timedelta(days=1)
    else:
        # From the start of the month.
        date_start = today - dt.timedelta(days=(today.day - 1))
        date_end = today + dt.timedelta(days=1)
    if os.path.exists(save_location):
        shutil.rmtree(save_location)
    os.mkdir(save_location)
    save_attachments(date_start, date_end, save_location, ext_pattern)
    for file in os.scandir(save_location):
        if file.name.endswith('zip'):
            unzip(file.path)


def unzip(file_path: str):
    zf = zipfile.ZipFile(file_path)
    unzip_path = Path(file_path).parent
    zf.extractall(unzip_path)


if __name__ == '__main__':
    today = dt.datetime.today()
    tomorrow = today + dt.timedelta(days=1)
    path = r'C:\Max\my scripts\bank_parser\Examples\MAIL'
    pattern = "urn:schemas:httpmail:fromemail LIKE '%voz.ru'"
    save_attachments(today, tomorrow, path, pattern)
