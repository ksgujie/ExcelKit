"""ExcelKit 0.6.0 报表可视化、排序和公式计算回归测试。"""

from __future__ import annotations

import base64
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from excelkit import Workbook, Worksheet
from excelkit.chart import ChartType
from excelkit.note import Note
from excelkit.sort import SortKey


class FormulaAndDataOperationTests(unittest.TestCase):
    """验证新增公式函数、命名区域、筛选和排序。"""

    def test_extended_formula_functions_and_named_range_references(self) -> None:
        """功能：验证 SUMIF、IFERROR、日期函数、查找函数和命名区域可由计算器求值。"""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["编号", "金额"])
        worksheet.append([1, 10])
        worksheet.append([2, 20])
        workbook.add_named_range("Amounts", worksheet.range("B2:B3"))
        worksheet["C1"].formula = "=SUM(Amounts)"
        worksheet["C2"].formula = "=SUMIF(B2:B3,\">10\")"
        worksheet["C3"].formula = "=IFERROR(1/0,99)"
        worksheet["C4"].formula = "=XLOOKUP(2,A2:A3,B2:B3)"
        worksheet["C5"].formula = "=VLOOKUP(1,A2:B3,2,FALSE)"
        worksheet["C6"].formula = "=YEAR(DATE(2026,8,1))"
        worksheet["C7"].formula = "=ROUNDUP(1.21,1)"
        worksheet["C8"].formula = "=ROUNDDOWN(1.29,1)"
        worksheet["C9"].formula = "=HLOOKUP(10,A2:B3,2,FALSE)"
        worksheet["C10"].formula = "=COUNTIF(B2:B3,20)"
        worksheet["C11"].formula = "=AVERAGEIF(B2:B3,\">10\")"

        workbook.calculate(strict=True)

        self.assertEqual(worksheet["C1"].cached_value, 30)
        self.assertEqual(worksheet["C2"].cached_value, 20)
        self.assertEqual(worksheet["C3"].cached_value, 99)
        self.assertEqual(worksheet["C4"].cached_value, 20)
        self.assertEqual(worksheet["C5"].cached_value, 10)
        self.assertEqual(worksheet["C6"].cached_value, 2026)
        self.assertEqual(worksheet["C7"].cached_value, 1.3)
        self.assertEqual(worksheet["C8"].cached_value, 1.2)
        self.assertEqual(worksheet["C9"].cached_value, 20)
        self.assertEqual(worksheet["C10"].cached_value, 1)
        self.assertEqual(worksheet["C11"].cached_value, 20)

    def test_sort_filter_visibility_and_note(self) -> None:
        """功能：验证排序会移动数据、筛选会隐藏行、可见性和批注可读写。"""
        worksheet = Workbook().active
        worksheet.append(["姓名", "成绩", "状态"])
        worksheet.append(["张三", 80, "通过"])
        worksheet.append(["李四", 95, "通过"])
        worksheet.append(["王五", 60, "不通过"])
        worksheet.sort("A2:C4", keys=[SortKey(1, descending=True)])
        self.assertEqual(worksheet["A2"].value, "李四")
        worksheet.auto_filter.range = "A1:C4"
        worksheet.auto_filter.set(2, ["通过"]).apply()
        self.assertFalse(worksheet.row(1).hidden)
        self.assertTrue(worksheet.row(3).hidden)
        worksheet.visibility = Worksheet.HIDDEN
        self.assertEqual(worksheet.visibility, Worksheet.HIDDEN)
        worksheet["A2"].note = Note("优秀", author="教务处")
        self.assertEqual(worksheet["A2"].note.author, "教务处")
        worksheet["A2"].note = None
        self.assertIsNone(worksheet["A2"].note)


class XlsxVisualTests(unittest.TestCase):
    """验证图片、图表和批注的 XLSX 写出及批注读取。"""

    def test_chart_image_note_and_visibility_are_written(self) -> None:
        """功能：验证生成的 XLSX 包包含图片、图表、DrawingML、批注和隐藏状态。"""
        png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/"
            "9kL7UAAAAABJRU5ErkJggg=="
        )
        with tempfile.TemporaryDirectory() as directory:
            image_file = Path(directory) / "pixel.png"
            image_file.write_bytes(png)
            filename = Path(directory) / "visual.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["月份", "销售额"])
            worksheet.append(["一月", 10])
            worksheet.append(["二月", 20])
            worksheet["A2"].note = Note("首月数据", author="财务部")
            worksheet.visibility = Worksheet.HIDDEN
            worksheet.add_image(str(image_file), anchor="E2").alt_text = "像素图"
            chart = worksheet.add_chart(ChartType.COLUMN, anchor="E5")
            chart.title = "月度销售"
            chart.add_series(values="B2:B3", categories="A2:A3", name="销售额")
            workbook.save(filename)

            with zipfile.ZipFile(filename) as package:
                names = set(package.namelist())
                self.assertIn("xl/media/image1.png", names)
                self.assertIn("xl/charts/chart1.xml", names)
                self.assertIn("xl/drawings/drawing1.xml", names)
                self.assertIn("xl/comments1.xml", names)
                for member in names:
                    if member.endswith((".xml", ".rels", ".vml")):
                        ET.fromstring(package.read(member))
            loaded = Workbook.load(filename)

        self.assertEqual(loaded.active.visibility, Worksheet.HIDDEN)
        self.assertEqual(loaded.active["A2"].note.text, "首月数据")
        self.assertEqual(loaded.active["A2"].note.author, "财务部")


if __name__ == "__main__":
    unittest.main()
