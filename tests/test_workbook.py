"""工作簿基础功能测试。"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import Workbook
from excelkit.errors import InvalidWorksheetNameError, ValidationError


class TestWorkbookBasic(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        self.ws = self.wb.active

    def test_active_and_add_sheet(self):
        self.assertIsNotNone(self.wb.active)
        ws2 = self.wb.add_sheet("销售")
        self.assertEqual(ws2.name, "销售")
        self.assertIs(self.wb.get_sheet("销售"), ws2)
        self.assertIs(self.wb.get_sheet(0), self.ws)
        self.assertIs(self.wb.get_sheet(1), ws2)

    def test_sheet_name_case_insensitive(self):
        self.ws.name = "MySheet"
        self.assertIs(self.wb.get_sheet("mysheet"), self.ws)

    def test_remove_sheet(self):
        self.wb.add_sheet("Tmp")
        self.wb.remove_sheet("Tmp")
        with self.assertRaises(InvalidWorksheetNameError):
            self.wb.get_sheet("Tmp")

    def test_move_sheet(self):
        self.wb.add_sheet("B")
        self.wb.add_sheet("C")
        self.wb.move_sheet("C", 0)
        self.assertEqual(self.wb.sheets[0].name, "C")

    def test_copy_sheet(self):
        self.ws["A1"] = 1
        self.wb.copy_sheet("Sheet1", "副本")
        self.assertEqual(self.wb.get_sheet("副本")["A1"].value, 1)

    def test_invalid_sheet_name(self):
        with self.assertRaises(InvalidWorksheetNameError):
            self.wb.add_sheet("Bad/Name")

    def test_validate(self):
        self.assertEqual(self.wb.validate(), [])

    def test_save_validate_failure(self):
        ws = self.wb.active
        ws.range("A1:B2").merge()
        ws.range("A2:B3").merge()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            with self.assertRaises(ValidationError):
                self.wb.save(path, validate=True)
        finally:
            os.unlink(path)


class TestTabColor(unittest.TestCase):
    def test_normalize_and_clear(self):
        ws = Workbook().active
        ws.tab_color = "#4472c4"
        self.assertEqual(ws.tab_color, "4472C4")
        ws.tab_color = "FFFF0000"
        self.assertEqual(ws.tab_color, "FFFF0000")
        ws.tab_color = None
        self.assertIsNone(ws.tab_color)

    def test_invalid_color(self):
        ws = Workbook().active
        for bad in ("", "#", "12345", "GGGGGG", "1234567", 123):
            with self.assertRaises(ValueError):
                ws.tab_color = bad

    def test_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "x"
        ws.tab_color = "4472C4"
        wb.add_sheet("无颜色")
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save(path)
            wb2 = Workbook.load(path)
            self.assertEqual(wb2.get_sheet("Sheet1").tab_color, "FF4472C4")
            self.assertIsNone(wb2.get_sheet("无颜色").tab_color)
        finally:
            os.unlink(path)

    def test_copy_sheet_keeps_color(self):
        wb = Workbook()
        wb.active.tab_color = "00FF00"
        wb.copy_sheet("Sheet1", "副本")
        self.assertEqual(wb.get_sheet("副本").tab_color, "00FF00")


if __name__ == "__main__":
    unittest.main()
