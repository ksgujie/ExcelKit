"""ExcelKit 0.7.0 数据查找、替换与文本导出回归测试。"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from excelkit import Workbook
from excelkit.errors import InvalidFileError
from excelkit.hyperlink import Hyperlink
from excelkit.note import Note
from excelkit.style import Font, Style


class SearchReplaceTests(unittest.TestCase):
    """验证工作表查找、替换和增强区域清除 API。"""

    def setUp(self) -> None:
        """功能：为每个查找替换测试建立包含值和公式的工作表。"""
        self.worksheet = Workbook().active
        self.worksheet.append(["Alpha", "alphabet", 2])
        self.worksheet.append(["ALPHA", "Beta", 3])
        self.worksheet["C3"].formula = "=SUM(C1:C2)"

    def test_find_values_and_formulas(self) -> None:
        """功能：验证大小写、完整匹配和公式查找语义。"""
        self.assertEqual(
            tuple(cell.address for cell in self.worksheet.find("alpha")),
            ("A1", "B1", "A2"),
        )
        self.assertEqual(
            tuple(cell.address for cell in self.worksheet.find("Alpha", whole=True)),
            ("A1", "A2"),
        )
        self.assertEqual(
            tuple(cell.address for cell in self.worksheet.find(
                "Alpha", whole=True, match_case=True
            )),
            ("A1",),
        )
        self.assertEqual(
            tuple(cell.address for cell in self.worksheet.find("SUM", in_formulas=True)),
            ("C3",),
        )
        self.assertEqual(
            tuple(cell.address for cell in self.worksheet.find(2)), ("C1",)
        )

    def test_replace_values_and_formulas(self) -> None:
        """功能：验证值、公式、限制次数和类型校验。"""
        self.assertEqual(self.worksheet.replace("alpha", "gamma", limit=2), 2)
        self.assertEqual(self.worksheet["A1"].value, "gamma")
        self.assertEqual(self.worksheet["B1"].value, "gammabet")
        self.assertEqual(self.worksheet["A2"].value, "ALPHA")
        self.assertEqual(
            self.worksheet.replace("SUM", "AVERAGE", in_formulas=True), 1
        )
        self.assertEqual(self.worksheet["C3"].formula, "=AVERAGE(C1:C2)")
        self.assertEqual(self.worksheet.replace(2, 20), 1)
        self.assertEqual(self.worksheet["C1"].value, 20)
        self.assertEqual(self.worksheet.replace("Beta", r"\\path", whole=True), 1)
        self.assertEqual(self.worksheet["B2"].value, r"\\path")
        with self.assertRaises(TypeError):
            self.worksheet.replace("A", 1)
        with self.assertRaises(ValueError):
            self.worksheet.find("")

    def test_clear_optional_hyperlinks_and_notes(self) -> None:
        """功能：验证 Range.clear 默认兼容行为和扩展清除开关。"""
        cell = self.worksheet["A1"]
        cell.style = Style(font=Font(bold=True))
        cell.hyperlink = Hyperlink("https://example.com")
        cell.note = Note("请核对", author="测试")
        self.worksheet.range("A1:A1").clear()
        self.assertIsNone(cell.value)
        self.assertIsNotNone(cell.hyperlink)
        self.assertIsNotNone(cell.note)
        self.worksheet.range("A1:A1").clear(
            values=False, styles=False, hyperlinks=True, notes=True
        )
        self.assertIsNone(cell.hyperlink)
        self.assertIsNone(cell.note)


class DelimitedExportTests(unittest.TestCase):
    """验证 CSV/TSV 导出以及 Workbook.save 分派。"""

    def test_worksheet_export_and_text_workbook_save(self) -> None:
        """功能：验证 CSV 往返、TSV 公式输出模式。"""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["姓名", "说明", "结果"])
        worksheet.append(["张三", "逗号,与引号\"", 2])
        worksheet["C3"].formula = "=SUM(C2:C2)"
        workbook.calculate(strict=True)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            csv_file = root / "sales.csv"
            formula_file = root / "formula.tsv"
            self.assertIs(worksheet.export(csv_file), worksheet)
            loaded = Workbook.load(csv_file)
            self.assertEqual(loaded.active["A2"].value, "张三")
            self.assertEqual(loaded.active["B2"].value, "逗号,与引号\"")
            self.assertEqual(loaded.active["C3"].value, "2")
            workbook.save(formula_file, formulas=True)
            self.assertIn("=SUM(C2:C2)", formula_file.read_text(encoding="utf-8-sig"))

    def test_text_save_requires_one_sheet_and_excel_rejects_text_options(self) -> None:
        """功能：验证文本导出不会静默丢弃额外工作表或错误选项。"""
        workbook = Workbook()
        workbook.active["A1"] = "数据"
        workbook.add_sheet("第二张")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with self.assertRaises(InvalidFileError):
                workbook.save(root / "many.csv")
            with self.assertRaises(ValueError):
                workbook.save(root / "book.xlsx", encoding="gb18030")


if __name__ == "__main__":
    unittest.main()
