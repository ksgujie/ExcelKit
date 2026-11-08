"""单元格读写与类型转换测试。"""

import os
import sys
import unittest
from datetime import date, datetime

sys.path.insert(0, r'W:\eyes\Temp\opencode')

from excelkit import Workbook
from excelkit.style import Font, Style


class TestCellReadWrite(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        self.ws = self.wb.active

    def test_value_write_read(self):
        self.ws["A1"] = "你好"
        self.assertEqual(self.ws["A1"].value, "你好")
        self.ws["B1"] = 42
        self.assertEqual(self.ws["B1"].value, 42)

    def test_cell_index_address(self):
        cell = self.ws.cell(2, 2)
        self.assertEqual(cell.index, (2, 2))
        self.assertEqual(cell.address, "C3")
        self.assertEqual(cell.row, 2)
        self.assertEqual(cell.column, 2)

    def test_formula(self):
        self.ws["A1"] = 10
        self.ws["A2"] = 20
        self.ws["A3"].formula = "=SUM(A1:A2)"
        self.assertEqual(self.ws["A3"].formula, "=SUM(A1:A2)")
        self.ws["A3"].formula = None
        self.assertIsNone(self.ws["A3"].formula)

    def test_clear_value_overwrites_formula(self):
        self.ws["A1"].formula = "=1+1"
        self.ws["A1"] = 5
        self.assertIsNone(self.ws["A1"].formula)
        self.assertEqual(self.ws["A1"].value, 5)

    def test_copy_style(self):
        self.ws["A1"].style = Style(font=Font(bold=True))
        self.ws["B1"].copy_style(self.ws["A1"])
        self.assertTrue(self.ws["B1"].style.font.bold)

    def test_read_snapshot(self):
        self.ws["A1"] = 10
        snap = self.ws["A1"].read()
        self.ws["A1"] = 99
        self.assertEqual(snap.value, 10)
        self.assertEqual(self.ws["A1"].value, 99)

    def test_as_int_float_bool_date(self):
        self.ws["A1"] = "123"
        self.assertEqual(self.ws["A1"].read().as_int(), 123)
        self.ws["A2"] = "3.14"
        self.assertEqual(self.ws["A2"].read().as_float(), 3.14)
        self.ws["A3"] = "是"
        self.assertTrue(self.ws["A3"].read().as_bool())
        self.ws["A4"] = "#2026-8-1"
        self.assertEqual(self.ws["A4"].read().as_date(), date(2026, 8, 1))
        self.ws["A5"] = "#2026/8/1 12:33"
        self.assertEqual(self.ws["A5"].read().as_datetime(), datetime(2026, 8, 1, 12, 33))


if __name__ == "__main__":
    unittest.main()
