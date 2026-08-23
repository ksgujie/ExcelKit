"""基础 XLSX 包、公式和数据类型写出的回归测试。"""

import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

from excelkit import Workbook
from excelkit.writer.xlsx import XlsxWriter

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"m": MAIN_NS}


class XlsxWriterTests(unittest.TestCase):
    """验证唯一 XlsxWriter 入口生成的 ZIP 和 SpreadsheetML。"""

    def test_minimum_workbook_and_return_value(self):
        """功能：验证空工作簿保存、返回值及最小 XLSX 部件集合。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "empty.xlsx"
            workbook = Workbook()
            self.assertIs(workbook.save(filename), workbook)
            with zipfile.ZipFile(filename) as package:
                self.assertEqual(
                    set(package.namelist()),
                    {
                        "[Content_Types].xml",
                        "_rels/.rels",
                        "xl/workbook.xml",
                        "xl/_rels/workbook.xml.rels",
                        "xl/styles.xml",
                        "xl/worksheets/sheet1.xml",
                    },
                )
                for member in package.namelist():
                    ET.fromstring(package.read(member))

    def test_formula_and_supported_values_are_written(self):
        """功能：验证公式、布尔、数字、日期时间和普通对象的 XML 表示。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "types.xlsx"
            workbook = Workbook()
            sheet = workbook.add_sheet("数据 & 公式")
            sheet.label_color = "4472C4"
            sheet.append([" text ", True, 42, 3.5, date(2026, 8, 22)])
            sheet.append([datetime(2026, 8, 22, 12, 30), object()])
            sheet["F1"].formula = "=SUM(C1:D1)"
            workbook.save(filename)

            with zipfile.ZipFile(filename) as package:
                root = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
                self.assertEqual(
                    root.find("m:sheetPr/m:tabColor", NS).attrib["rgb"],
                    "FF4472C4",
                )
                cells = {cell.attrib["r"]: cell for cell in root.findall(".//m:c", NS)}
                self.assertEqual(cells["A1"].attrib["t"], "inlineStr")
                self.assertEqual(cells["B1"].attrib["t"], "b")
                self.assertEqual(cells["C1"].find("m:v", NS).text, "42")
                self.assertEqual(cells["E1"].attrib["t"], "d")
                self.assertEqual(cells["E1"].find("m:v", NS).text, "2026-08-22")
                self.assertEqual(
                    cells["A2"].find("m:v", NS).text,
                    "2026-08-22T12:30:00",
                )
                self.assertEqual(cells["F1"].find("m:f", NS).text, "SUM(C1:D1)")
                self.assertIsNone(cells["F1"].find("m:v", NS))

    def test_multiple_sheets_use_zero_based_iteration_internally(self):
        """功能：验证 0-based 遍历仍生成从 sheet1 开始的标准文件名。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "multi.xlsx"
            workbook = Workbook()
            workbook.add_sheet("第一张")["A1"] = 1
            workbook.add_sheet("第二张")["A1"] = 2
            XlsxWriter(workbook).write(filename)
            with zipfile.ZipFile(filename) as package:
                self.assertIn("xl/worksheets/sheet1.xml", package.namelist())
                self.assertIn("xl/worksheets/sheet2.xml", package.namelist())
                root = ET.fromstring(package.read("xl/workbook.xml"))
                names = [item.attrib["name"] for item in root.findall("m:sheets/m:sheet", NS)]
                self.assertEqual(names, ["第一张", "第二张"])

    def test_serialization_failure_keeps_existing_file(self):
        """功能：验证 XML 生成失败时已有目标文件不会被替换。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "safe.xlsx"
            filename.write_bytes(b"original")
            workbook = Workbook()
            workbook.active["A1"] = 1
            with patch("excelkit.writer.xlsx.sheet_xml", side_effect=RuntimeError("测试异常")):
                with self.assertRaises(RuntimeError):
                    workbook.save(filename)
            self.assertEqual(filename.read_bytes(), b"original")


if __name__ == "__main__":
    unittest.main()
