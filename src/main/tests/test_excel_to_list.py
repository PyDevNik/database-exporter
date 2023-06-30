import unittest
from typing import List
import os
from ..excel_to_list import create_list
import start

class TestExcelToList(unittest.TestCase):
    def test_create_list(self):
        if not os.path.exists('result.xlsx'):
            start.start()
        result = create_list('result.xlsx', 
            'Код ТНВЭД ТС', 
            'Документ о подтверждении соответствия', 
            desc_type = "documents")
        self.assertIsNotNone(result)
        self.assertEqual(type(result), List[dict])
        self.assertTrue(len(result) > 0)
    

if __name__ == "__main__":
    unittest.main()