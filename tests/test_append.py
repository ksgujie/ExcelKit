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

    def test_append_rows_mismatch_error(self):
        with self.assertRaises(ValueError):
            self.ws.append_rows([["a", 1], ["b"]], on_mismatch="error")
        # error 模式保持原子性，不写入
        self.assertEqual(self.ws.values, [])

    def test_append_rows_mismatch_warn(self):
        import warnings
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            self.ws.append_rows([["a", 1], ["b"]], on_mismatch="warn")
            self.assertTrue(any(issubclass(x.category, UserWarning) for x in w))
        # 允许参差行，按实际列数写入
        self.assertEqual(self.ws["A1"].value, "a")
        self.assertEqual(self.ws["B1"].value, 1)
        self.assertEqual(self.ws["A2"].value, "b")
        self.assertIsNone(self.ws["B2"].value)

    def test_append_rows_mismatch_ignore(self):
        import warnings
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            self.ws.append_rows([["a", 1], ["b"]], on_mismatch="ignore")
            self.assertFalse(any(issubclass(x.category, UserWarning) for x in w))
        self.assertEqual(self.ws["A2"].value, "b")

    def test_append_rows_default_warn(self):
        import warnings
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            self.ws.append_rows([["a", 1], ["b", 2, 3]])
            self.assertTrue(len(w) >= 1)
        self.assertEqual(self.ws["C2"].value, 3)

    def test_append_rows_invalid_on_mismatch(self):
        with self.assertRaises(ValueError):
            self.ws.append_rows([["a"]], on_mismatch="bad")


if __name__ == "__main__":
    unittest.main()
