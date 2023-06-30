'''
This is an example of usage of 
excel_to_list.py module
'''

from excel_to_list import create_list

result = create_list('result.xlsx', 
    'Код ТНВЭД ТС', 
    'Документ о подтверждении соответствия', 
    desc_type = "documents")