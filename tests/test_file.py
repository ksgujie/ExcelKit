"""文件读写往返测试。"""

import os
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import Workbook


class TestFileRoundTrip(unittest.TestCase):
    def test_xlsx_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "文本"
        ws["B1"] = 123
        ws["C1"] = 3.14
        ws["D1"] = True
        ws["E1"] = date(2026, 8, 1)
        ws.append(["第二行", 999])
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save(path)
            wb2 = Workbook.load(path)
            ws2 = wb2.active
            self.assertEqual(ws2["A1"].value, "文本")
            self.assertEqual(ws2["B1"].value, 123)
            self.assertEqual(ws2["C1"].value, 3.14)
            self.assertEqual(ws2["D1"].value, True)
            self.assertEqual(ws2["A2"].value, "第二行")
            self.assertEqual(ws2["B2"].value, 999)
        finally:
            os.unlink(path)

    def test_csv_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws.append_rows([["姓名", "成绩"], ["张三", 95]])
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="w", encoding="utf-8") as f:
            path = f.name
        try:
            wb.save(path)
            wb2 = Workbook.load(path)
            self.assertEqual(wb2.active.values, [["姓名", "成绩"], ["张三", 95]])
        finally:
            os.unlink(path)


if __name__ == "__main__":
    unittest.main()
