from contextlib import contextmanager
import os
import re
from typing import Set
import zipfile


from bs4 import BeautifulSoup
import pandas as pd
import win32com.client as win32


def path_printer(fn):
    def g(path, *args, **kwargs):
        print(path)
        return fn(path, *args, **kwargs)
    return g


@path_printer
def alpha(path: str, pattern: str) -> Set[str]:
    with open(path, encoding='utf-8') as f:
        raw_text = f.read()
    soup = BeautifulSoup(raw_text, 'lxml')
    cells = soup.find_all('td', text=re.compile(f'^{pattern}', re.M | re.I))
    return set(cell.text.strip() for cell in cells)


def sber(path: str, pattern: str) -> Set[str]:
    try:
        with open(path, errors='ignore') as f:
            raw_text = f.read()
    except Exception as e:
        print(path, e)
    else:
        soup = BeautifulSoup(raw_text, 'xml')
        search_attrs = {'ss:Index': '41' if 'спец' in path else '36'}
        cells = soup.find_all('Cell', attrs=search_attrs)
        data = [cell.find('Data', string=re.compile(pattern, re.I)) for cell in cells]
        result = set([item.text for item in data if item])
        return result
    return set()


@path_printer
def vtb(path: str, pattern: str, temp_dir: str = None):
    """
    Converts
    """
    if not temp_dir:
        temp_dir = r'C:\Max\vtb_convert_temp'
    if not os.path.isdir(temp_dir):
        os.mkdir(temp_dir)
    if path.endswith('zip'):
        archive = zipfile.ZipFile(path)
        archive.extractall(temp_dir)
        for file in os.scandir(temp_dir):
            if file.is_dir():
                continue
            new_path = rtf_to_html_one(file.path, temp_dir)
            result = _vtb_parser(new_path, pattern)
            os.remove(new_path)
            os.remove(file.path)
            yield result
    else:
        new_path = rtf_to_html_one(path, temp_dir)
        result = _vtb_parser(new_path, pattern)
        return result


@path_printer
def _vtb_parser(path: str, pattern: str) -> Set[str]:
    with open(path) as f:
        dfs = pd.read_html(f, attrs={'border': '1'}, flavor='lxml', encoding='utf-8')
    df = max(dfs, key=len)
    constraint = df[8].str.contains(pattern, case=False, na=False)
    df = df[constraint]
    return set(df[8])


def rtf_to_html_one(path: str, convert_dir: str) -> str:
    """
    Converts rtf file to html and puts it to the temp_dir.

    Parameters
    ----------
    path
        Path to rtf file to convert.
    temp_dir
        Path to the directory for converted file.
    Returns
    -------
    str
        Path to the converted html file.

    """
    with word() as word_app:
        doc = word_app.Documents.Open(path)
        *_, file_name = path.rsplit('\\')
        new_name = file_name[:-3] + 'html'
        new_path = os.path.join(convert_dir, new_name)
        wdFormatHTML = 10
        doc.SaveAs2(new_path, wdFormatHTML)
        doc.Close()
    return new_path


@contextmanager
def word():
    word_app = win32.Dispatch('word.application')
    try:
        yield word_app
    finally:
        word_app.Quit()
