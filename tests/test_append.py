"""追加数据测试。"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import Workbook
from excelkit.errors import ExcelKitError


class TestAppend(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        self.ws = self.wb.active

    def test_append(self):
        self.ws.append(["姓名", "成绩"])
        self.ws.append(["张三", 95])
        self.assertEqual(self.ws.max_row, 1)
        self.assertEqual(self.ws.values, [["姓名", "成绩"], ["张三", 95]])

    def test_append_rows(self):
        self.ws.append_rows([["a", 1], ["b", 2], ["c", 3]])
        self.assertEqual(self.ws.values, [["a", 1], ["b", 2], ["c", 3]])

    def test_append_rows_mismatch(self):
        with self.assertRaises(ExcelKitError):
            self.ws.append_rows([["a", 1], ["b"]])


if __name__ == "__main__":
    unittest.main()
