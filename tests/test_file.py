"""文件读写往返测试。"""

import os
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
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



    def test_xlsx_merge_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "x"
        ws.range("A1:B2").merge()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save(path)
            wb2 = Workbook.load(path)
            self.assertEqual(wb2.active.merged_ranges, ((0, 0, 1, 1),))
        finally:
            os.unlink(path)


    def test_xlsx_freeze_and_autofilter_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "h1"
        ws["B1"] = "h2"
        ws["A2"] = 1
        ws["B2"] = 2
        ws.freeze_panes = "A2"
        ws.auto_filter.range = "A1:B2"
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save(path)
            wb2 = Workbook.load(path)
            self.assertEqual(wb2.active.freeze_panes, "A2")
            self.assertEqual(wb2.active.auto_filter.range, "A1:B2")
        finally:
            os.unlink(path)


    def test_xlsx_dimension_covers_merge_and_image_anchor(self):
        wb = Workbook()
        ws = wb.active
        # merge with no cell values written beyond default state
        ws.range("A1:C5").merge()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save(path)
            with zipfile.ZipFile(path) as zf:
                root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            dim = root.find("m:dimension", ns)
            self.assertIsNotNone(dim)
            self.assertEqual(dim.get("ref"), "A1:C5")
        finally:
            os.unlink(path)

    def test_xlsx_inline_string_uses_is_element(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "姓名"
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save(path)
            with zipfile.ZipFile(path) as zf:
                root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            cell = root.find(".//m:c[@r='A1']", ns)
            self.assertIsNotNone(cell)
            self.assertEqual(cell.get("t"), "inlineStr")
            self.assertIsNotNone(cell.find("m:is", ns))
            self.assertIsNone(cell.find("m:inlineStr", ns))
            t_el = cell.find("m:is/m:t", ns)
            self.assertIsNotNone(t_el)
            self.assertEqual(t_el.text, "姓名")
        finally:
            os.unlink(path)

    def test_xlsx_reader_supports_legacy_inline_str_element(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "旧格式"
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            wb.save(path)
            # 模拟旧版本写出的 <inlineStr> 元素，验证读取端兼容
            with zipfile.ZipFile(path) as zf:
                sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            sheet = sheet.replace("<is>", "<inlineStr>").replace("</is>", "</inlineStr>")
            tmp = path + ".tmp"
            with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
                for item in src.infolist():
                    data = src.read(item.filename)
                    if item.filename == "xl/worksheets/sheet1.xml":
                        data = sheet.encode("utf-8")
                    dst.writestr(item, data)
            os.replace(tmp, path)
            wb2 = Workbook.load(path)
            self.assertEqual(wb2.active["A1"].value, "旧格式")
        finally:
            os.unlink(path)


if __name__ == "__main__":
    unittest.main()
