from collections import defaultdict
import csv
import dataclasses
import json
import os
import re
import string
from typing import DefaultDict, Dict, Iterable, List, Set, Any

import xlrd


@dataclasses.dataclass
class Field:
    name: str
    column_index: int
    keywords: List[str]
    anti_keywords: List[str]


def find_keys(rows: Iterable[Iterable], name_field_map: Dict[str, Field]) -> List[Field]:
    for row in rows:
        # if file.path == 'C:\\Max\\my scripts\\bank_parser\\Examples\\MAIL\\15.04.2020.xls':
        #     breakpoint()
        for column, cell in enumerate(row):
            try:
                if cell.ctype != 1:
                    continue
                value = cell.value.lower()
            except AttributeError:
                if not cell:
                    continue
                value = cell.lower()
            for field in name_field_map.values():
                has_keywords = any(keyword in value for keyword in field.keywords)
                has_anti_keywords = any(keyword in value for keyword in field.anti_keywords)
                if not field.column_index and has_keywords and not has_anti_keywords:
                    field.column_index = column
        if ((name_field_map['Излишек'].column_index and name_field_map['Недостача'].column_index)
            or name_field_map['Расхождение'].column_index) and \
                (name_field_map['Сомнительная'].column_index
                    or name_field_map['Неплатежеспособная'].column_index
                    or name_field_map['Фальшивая'].column_index):
            break
    return list(name_field_map.values())


def parse_rows(sheet_rows, name_field_map: Dict[str, Field]) -> DefaultDict[str, Dict[str, float]]:
    result: DefaultDict[str, Dict[str, float]] = defaultdict(dict)

    fields = [field for field in find_keys(sheet_rows, name_field_map) if field.column_index > 0]
    if len(fields) < 2:
        raise RuntimeError(f'Keys not found in non-empty sheet of the book')
    last_column = min(fields, key=lambda f: f.column_index).column_index - 1
    for row_index, row in enumerate(sheet_rows):
        try:
            row_stringified = [str(cell.value) for cell in row[:last_column]]
        except AttributeError:
            row_stringified = row[:last_column]
        key = ';'.join(row_stringified)

        if '2;3;4' in key or '2.0;3.0;4.0' in key or 'итого' in key.lower():
            continue

        for field in fields:
            try:
                cell_value = str(row[field.column_index].value)
            except AttributeError:
                cell_value = row[field.column_index]
            cell_value = re.sub(r'\s', '', cell_value)
            cell_value = re.sub(f'[{string.punctuation}]', '.', cell_value)
            try:
                field_value = float(cell_value)
            except ValueError:
                continue
            if field_value and float(field_value) > 50.0:
                result[key].update({
                    field.name: float(field_value)})
    return result


def parse_xls(path: str, name_field_map: Dict[str, Field]) -> Dict[str, Dict[str, float]]:
    result: Dict[str, Dict[str, float]] = {}
    with xlrd.open_workbook(path, encoding_override='cp1251') as wb:
        for sh in wb.sheets():
            if sh.nrows < 2:
                # Empty sheet.
                continue
            sheet_rows = sh.get_rows()
            parsed_sheet = parse_rows(sheet_rows, name_field_map)
            result = {**result, **parsed_sheet}
    return result


def parse_csv(path: str, name_field_map: Dict[str, Field]) -> Dict[str, Dict[str, float]]:
    with open(path) as f:
        reader = csv.reader(f, delimiter=';')
        result = parse_rows(reader, name_field_map)
    return result


def parse_registry(path: str) -> Dict[str, Dict[str, float]]:
    fields = [
        Field('Излишек', 0, ['излиш', ], []),
        Field('Недостача', 0, ['недост', ], []),
        Field('Расхождение', 0, ['расхожд', ], []),
        Field('Сомнительная', 0, ['сомнит', ], []),
        Field('Неплатежеспособная', 0, ['неплатеж', ], []),
        Field('Фальшивая', 0, ['подд', 'фальш', ], ['неплатеж', ]),
    ]
    name_field_map = {field.name: field for field in fields}
    result: Dict[str, Dict[str, float]] = {}
    try:
        if path.endswith('xls') or path.endswith('xlsx'):
            result = parse_xls(path, name_field_map)
        elif path.endswith('csv'):
            result = parse_csv(path, name_field_map)
        else:
            print('not an excel file')
    except RuntimeError:
        print(locals())
        print(f'keys not found in non-empty sheet of the book {path}')
    except Exception as e:
        print(locals())
        print(e)
    return result


def parse_registries(save_location: str) -> Set[str]:
    """
    Parse all registries and convert to string for process_data.py module.
    """

    resulting_set = set()
    parsed_registries = ((file.path, parse_registry(file.path)) for file in os.scandir(save_location))
    for path, key_data_map in parsed_registries:
        for key, data in key_data_map.items():
            row = ' '.join([key, json.dumps(data, ensure_ascii=False), path])
            resulting_set.add(row)
    return resulting_set


def extract_SAP_codes(parsed_registry: Dict[str, Dict[str, Dict[str, float]]]):
    {'path': {'keys': {'keywords': 'value'}}}
    bad_results = []
    new_map: Dict[str, Dict[str, Any]] = {}
    for index, (path, key_data_map) in enumerate(parsed_registry.items()):
        for key, data in key_data_map.items():
            # Search for SAP code in key.
            date = re.search(r'(\d{1,2}[.]\d{2}[.]((\d{2}\b)|(\d{4}\b)))', key, re.I)
            if not date:
                date = re.search(r'(\d{1,2}[.]\d{2}[.]((\d{2}\b)|(\d{4}\b)))', path, re.I)
            clean_key = re.search(r'([A-Z]\d{3}\b)|([A-Z]{2}\d{2}\b)', key, re.I)
            if not clean_key or not date:
                bad_results.append((path, key, data))
            else:
                new_map.update({clean_key.group(): {**data, f'path': path, 'date': date.group()}})
    return new_map, bad_results


def main(path: str):
    total = {}
    for file in os.scandir(path):
        total.update({file.name: parse_registry(file.path)})
    clean_result, bad_results = extract_SAP_codes(total)
    print(*bad_results, sep='\n')
    return clean_result


if __name__ == '__main__':
    total = {}
    path = r'C:\Max\temp_for_registries\916.04.2020.xlsx'
    path = r'C:\Max\temp_for_registries\676806_Отчёт о результатах пересчёта денежной наличности по операциям инкассации_17.04.20 01.27.30.xlsx'
    total = {path: parse_registry(path)}
    # for file in os.scandir(path):
    #     total.update({file.name: parse_registry(file.path)})
    clean_result, bad_results = extract_SAP_codes(total)
    print(json.dumps(clean_result, indent=2, ensure_ascii=False))
    print(json.dumps(bad_results, indent=2, ensure_ascii=False))
