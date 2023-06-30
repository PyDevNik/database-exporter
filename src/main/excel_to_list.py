'''
This file converts excel_table to dataframe,
and then dataframe is converted to list of dicts
[view xlsx_to_list_example.py]
'''

from sprinter import load_excel_table, sort
from separator import replace_table_dict, SEPARATORS, get_descriptions
import json
from pprint import pprint
from typing import List
import os
from time import time
import start

SEPARATORS.append(',')

FILENAME1 = 'result.xlsx'
CODE_COLUMN1 = 'Код ТНВЭД ТС'
DESC_COLUMN1 = 'Документ о подтверждении соответствия'


def get_table_codes(filename: str, column: str) -> List[str]:
    """
    Получение кодов из таблицы
    :param filename: Название файла таблицы
    :param column: Название столбца с кодами
    :return: Отформатированные коды
    """
    codes = load_excel_table(filename, column)
    new_codes = []
    for code in codes:
        code = str(code)
        for separator in SEPARATORS:
            if separator in code:
                code = code.split(separator)
                for c in code:
                    code.append(c.strip().replace(" ", ""))
        new_codes.append(code)
    return new_codes


def create_list(filename: str, code_column: str, desc_column: str, desc_type="documents") -> List[dict]:
    """
    Создание списка словарей из таблицы
    :param filename: Название файла таблицы
    :param code_column: Название столбца с кодами
    :param desc_column: Название столбца с значениями по коду
    :param desc_type: Тип Документа (по умолчанию 'documents')
    """
    start_time = time()
    print("Loading data...")
    all_codes = get_table_codes(filename, code_column)
    with open('dict.json', 'r', encoding='utf8') as f:
        dictionary = json.load(f)
        replace_table_dict(filename, desc_column, dictionary)
    documents, columns = get_descriptions(filename, desc_column)
    final_result = []
    print("Creating list...")
    for code in all_codes:
        for dindex, docs in enumerate(documents):
            if isinstance(code, list):
                for c_value in code:
                    result = {}
                    result['code'] = c_value
                    result[desc_type] = docs[0] if desc_type == 'documents' else docs
                    documents.pop(dindex)
                    final_result.append(result)
                    break
            else:
                result = {}
                result['code'] = code
                result[desc_type] = docs[0] if desc_type == "documents" else docs
                documents.pop(dindex)
                final_result.append(result)
                break
    final_result = [i for n, i in enumerate(final_result) if i not in final_result[:n]]
    final_result = sort(final_result, desc_type) if desc_type == "documents" else print("No need to sort data...")
    end_time = time()
    print(f"List generated in {round(end_time-start_time)} seconds")
    return final_result


# Запускаем код
if __name__ == "__main__":
    if not os.path.exists('result.xlsx'):
        start.start()
    my_list = create_list(FILENAME1, CODE_COLUMN1, DESC_COLUMN1)
    pprint(my_list)