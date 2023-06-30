'''
This file generates dictionary
'''

import pandas as pd
import json
from typing import List
from pprint import pprint

def create_dict(data: List[str]) -> None:
    """
    Генерация словаря в dict.json
    :param data: Данные из бд.
    """

    # Перебираем ряды
    for row_index, row in enumerate(data):
        for document_name in row.split(','):

            # Читаем json
            with open('dict.json', 'r', encoding='utf8') as dict_file:
                documents_dict = json.load(dict_file)

            if document_name not in documents_dict:
                # Вводим соответствие в консоль
                i = input(f'Введите соответствие для "{document_name}" ({row_index+3}): ')

                # Сохраняем соответствие
                documents_dict[document_name] = i.strip()
                with open('dict.json', 'w', encoding='utf8') as dict_file:
                    json.dump(documents_dict, dict_file, ensure_ascii=False)


def load_excel_table(filename: str, column: str) -> List[str]:
    """
    Загрузка Excel таблицы
    :param filename: имя файла таблицы
    :return: данные из файла
    """
    xlsx = pd.read_excel(filename)
    df = pd.DataFrame(xlsx, columns=[column])
    data = df.iloc[1:-1, 0]
    return data


def sort(data_list, doc_type='documents') -> List[dict]:
    """
    Сортировка и упорядочивание данных в списке
    :param data_list: Список с данными
    :param doc_type: Тип документа
    :return: Отсортированные данные
    """
    print(f"Sorting {len(data_list)} data...")
    result = []
    codes = set()
    for item in data_list:
        code = item["code"]
        docs = item[doc_type]
        if code not in codes:
            codes.add(code)
            result.append({"code": code, doc_type: docs})
        else:
            for i, r in enumerate(result):
                if r["code"] == code:
                    result[i][doc_type].extend(docs)
                    break
    return result


# Запускаем код
if __name__ == '__main__':
    excel_data = load_excel_table("excel.xlsx")
    create_dict(excel_data)
