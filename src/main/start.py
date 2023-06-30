'''
Tis file runs main functions of this module
'''
from sprinter import create_dict
from separator import separate, load_excel_table, excel_to_dataframe
from time import time
from excel_to_list import DESC_COLUMN1

def start(use_dict: bool = False) -> None:
    """
    Запуск основных функций модуля
    :param use_dict: Создавать словарь соответствия? (по умолчанию False)
    """
    print('Запуск...')

    # Замена разделителей
    print('Замена разделителей...')
    start = time()
    data_list = load_excel_table('excel.xlsx', DESC_COLUMN1)
    table = excel_to_dataframe('excel.xlsx')
    separate(data_list, table)
    end = time()
    print(f'Заменено за {round(end-start)} секунд')
    if use_dict:
        # Генерация словаря
        print('Генерация словаря...')
        start = time()
        data_list = load_excel_table('result.xlsx', 'Документ о подтверждении соответствия')
        create_dict(data_list)
        end = time()
        print(f'Вы успешно создали словарь за {round(end-start)} секунд')

if __name__ == "__main__":
    start()