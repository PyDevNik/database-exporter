'''
This file replaces documents names
'''

import pandas as pd
import os
from typing import List, Dict, Tuple
from sprinter import load_excel_table

# Список разделителей
SEPARATORS = ['/', ', ', ' и ', ' или ']


def get_descriptions(filename, column) -> Tuple[List[list], List[str]]:
    """
    Получение списка Документов/Наименований
    :param filename: Название файла таблицы
    :param column: Название столбца таблицы с данными
    :return: Значение документов/наименований до/после создания списка    
    """
    descs = load_excel_table(filename, column)
    if filename == 'result.xlsx':
        result = [[desc.split(',')] for desc in descs]
    return result, descs


def separate(data: List[str], table: pd.DataFrame) -> None:
    """
    Замена разделителей
    :param data: Список с данными из таблицы
    :param table: DataFrame с данными
    """
    for row in data:
        for separator in SEPARATORS:
            if separator in row:
                repl_row = row.replace(separator, ',')
                table.replace(row, repl_row, True)
    table.to_excel('result.xlsx')


def excel_to_dataframe(filename: str) -> pd.DataFrame:
    """
    Конвертация Excel таблицы в DataFrame
    :param filename: Файл таблицы
    :return: DataFrame с данными
    """
    table = pd.read_excel(filename)
    return table


def replace_table_dict(filename: str, column: str, dictionary: Dict) -> None:
    """
    Замена значений в таблице
    :param filename: Файл таблицы
    :param column: Необходимый столбец с Наименованиями/Документами
    :param dictionary: Словарь соответствия
    """
    descs, clmns = get_descriptions(filename, column)
    table = excel_to_dataframe(filename)
    desc_keys = list(dictionary.keys())
    for desc_index, desc_list in enumerate(descs):
        for d_index, desc in enumerate(desc_list):
            for descr_index, descr in enumerate(desc):
                if descr in desc_keys:
                    desc_list[d_index][descr_index] = dictionary[desc[descr_index]]
                    for descript in desc_list:
                        if isinstance(descript, list):
                            new_desc = ",".join(descript)
                        else:
                            new_desc = ",".join(desc_list)
                        table.replace(clmns[desc_index+1], new_desc, True)
    os.remove('result.xlsx')
    table.to_excel('result.xlsx')


# Запускаем код
if __name__ == '__main__':
    our_file = "excel.xlsx"
    our_data = load_excel_table(our_file)
    our_table = excel_to_dataframe(our_file)
    separate(our_data, our_table)
