"""区域操作测试。"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import Workbook
from excelkit.errors import ExcelKitError


class TestRange(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.append_rows([[1, 2, 3], [4, 5, 6], [7, 8, 9]])

    def test_values(self):
        rng = self.ws.range("A1:C3")
        self.assertEqual(rng.values, [[1, 2, 3], [4, 5, 6], [7, 8, 9]])

    def test_set_values(self):
        rng = self.ws.range("A1:B2")
        rng.set_values([[10, 20], [30, 40]])
        self.assertEqual(self.ws["A1"].value, 10)
        self.assertEqual(self.ws["B2"].value, 40)

    def test_set_values_shape_mismatch(self):
        rng = self.ws.range("A1:B2")
        with self.assertRaises(ExcelKitError):
            rng.set_values([[1], [2, 3]])

    def test_address_rows_columns(self):
        rng = self.ws.range("B2:C4")
        self.assertEqual(rng.address, "B2:C4")
        self.assertEqual(rng.rows, 3)
        self.assertEqual(rng.columns, 2)

    def test_clear(self):
        self.ws.range("A1:C3").clear(values=True, styles=False)
        self.assertIsNone(self.ws["A1"].value)

    def test_merge_unmerge(self):
        rng = self.ws.range("A1:B2")
        rng.merge()
        self.assertIn((0, 0, 1, 1), self.ws.merged_ranges)
        rng.unmerge()
        self.assertNotIn((0, 0, 1, 1), self.ws.merged_ranges)

    def test_copy_to(self):
        self.ws.range("A1:A2").set_values([[1], [2]])
        self.ws.range("A1:A2").copy_to(self.ws.range("C1:C2"))
        self.assertEqual(self.ws["C1"].value, 1)
        self.assertEqual(self.ws["C2"].value, 2)

    def test_to_records(self):
        ws2 = self.wb.add_sheet("Rec")
        ws2.append_rows([["姓名", "成绩"], ["张三", 95], ["李四", 88]])
        records = ws2.range("A1:B3").to_records()
        self.assertEqual(records, [
            {"姓名": "张三", "成绩": 95},
            {"姓名": "李四", "成绩": 88},
        ])

    def test_used_range(self):
        self.assertEqual(self.ws.used_range.address, "A1:C3")
        empty_ws = Workbook().active
        self.assertIsNone(empty_ws.used_range)


if __name__ == "__main__":
    unittest.main()
