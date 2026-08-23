"""ExcelKit 0.4.0 公式、复制、命名区域和数据表功能回归测试。"""

import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from excelkit import Workbook
from excelkit.errors import FormulaCalculationError
from excelkit.style import Fill, Font, Style


class FormulaCalculationTests(unittest.TestCase):
    """验证受控公式计算、依赖、状态和错误处理。"""

    def test_calculate_dependencies_functions_and_cross_sheet_references(self):
        """功能：验证常用函数、递归依赖和跨表引用可以写入缓存结果。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言计算结果和四种公式状态。
        """
        workbook = Workbook()
        source = workbook.add_sheet("原始 数据")
        source.append_rows([[1, 2], [3, 4]])
        result = workbook.add_sheet("结果")
        result["A1"].formula = "=SUM('原始 数据'!A1:B2)"
        result["A2"].formula = "=A1*2"
        result["A3"].formula = '=IF(A2>10,CONCAT("合计:",A2),0)'
        result["A4"].formula = '=CONCAT("")'

        self.assertIs(workbook.calculate(), workbook)
        self.assertEqual(result["A1"].cached_value, 10)
        self.assertEqual(result["A2"].cached_value, 20)
        self.assertEqual(result["A3"].cached_value, "合计:20")
        self.assertEqual(result["A3"].formula_status, "calculated")
        self.assertEqual(result["A3"].read().as_string(), "合计:20")
        self.assertIsNone(result["A3"].value)
        self.assertEqual(result["A4"].cached_value, "")
        self.assertEqual(result["A4"].formula_status, "calculated")

        source["A1"] = 10
        self.assertEqual(result["A1"].formula_status, "pending")
        self.assertIsNone(result["A1"].cached_value)
        workbook.calculate()
        self.assertEqual(result["A1"].cached_value, 19)

        before = result["A1"].cached_value
        with self.assertRaises(TypeError):
            result["B1"].formula = "="
        self.assertEqual(result["A1"].cached_value, before)

        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "formula.xlsx"
            workbook.save(filename)
            loaded = Workbook.load(filename)
            self.assertEqual(loaded.sheet("结果")["A4"].cached_value, "")
            self.assertEqual(
                loaded.sheet("结果")["A4"].formula_status, "calculated"
            )

    def test_calculation_errors_are_recorded_or_raised_in_strict_mode(self):
        """功能：验证非严格模式记录错误，严格模式抛出统一公式异常。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言错误状态、地址说明和循环引用处理。
        """
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"].formula = "=B1+1"
        sheet["B1"].formula = "=A1+1"
        workbook.calculate()
        self.assertEqual(sheet["A1"].formula_status, "error")
        self.assertIn("Sheet1!A1", sheet["A1"].calculation_error)
        with self.assertRaises(FormulaCalculationError):
            workbook.calculate(strict=True)
        with self.assertRaises(TypeError):
            workbook.calculate(strict=1)


class CopyAndRangeTests(unittest.TestCase):
    """验证单元格样式复制及区域清除、复制语义。"""

    def test_copy_style_and_range_copy_translate_relative_references(self):
        """功能：验证样式复制不复制值，区域复制会平移相对公式引用。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言绝对引用不移动且目标区域独立保存数据。
        """
        workbook = Workbook()
        sheet = workbook.active
        style = Style(font=Font(bold=True), fill=Fill("FFFF00"))
        sheet["A1"].value = 2
        sheet["A1"].style = style
        sheet["B1"].formula = "=A1+$A$1+A$1+$A1"

        target_cell = sheet["D1"]
        self.assertIs(target_cell.copy_style(sheet["A1"]), target_cell)
        self.assertEqual(target_cell.style, style)
        self.assertIsNone(target_cell.value)
        with self.assertRaises(TypeError):
            target_cell.copy_style(style)

        target = sheet.range("A3:B3")
        self.assertIs(sheet.range("A1:B1").copy_to(target), target)
        self.assertEqual(sheet["A3"].value, 2)
        self.assertEqual(sheet["B3"].formula, "=A3+$A$1+A$1+$A3")
        self.assertEqual(sheet["A3"].style, style)

    def test_clear_methods_keep_the_unselected_part(self):
        """功能：验证三种清除方法分别处理内容和样式。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言值、公式和样式按 API 开关保留或删除。
        """
        workbook = Workbook()
        sheet = workbook.active
        style = Style(font=Font(italic=True))
        sheet["A1"].value = 1
        sheet["A1"].style = style
        sheet.range("A1:A1").clear_styles()
        self.assertEqual(sheet["A1"].value, 1)
        self.assertEqual(sheet["A1"].style, Style())
        sheet["A1"].style = style
        sheet.range("A1:A1").clear_values()
        self.assertIsNone(sheet["A1"].value)
        self.assertEqual(sheet["A1"].style, style)
        sheet["A1"].value = 2
        sheet.range("A1:A1").clear()
        self.assertIsNone(sheet["A1"].value)
        self.assertEqual(sheet["A1"].style, Style())


class NamedRangeAndTableTests(unittest.TestCase):
    """验证工作簿命名区域和 XLSX 数据表生命周期。"""

    def test_named_range_lifecycle_follows_sheet_rename_and_removal(self):
        """功能：验证命名区域查询、重命名引用和随工作表删除。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言工作簿级名称大小写不敏感且区域动态引用工作表名。
        """
        workbook = Workbook()
        sheet = workbook.add_sheet("数据")
        named = workbook.add_named_range("SalesAmount", sheet.range("B2:B8"))
        self.assertIs(workbook.named_range("salesamount"), named)
        self.assertEqual(named.range.address, "B2:B8")
        sheet.name = "销售数据"
        self.assertEqual(named.worksheet.name, "销售数据")
        self.assertEqual(workbook.named_ranges, (named,))
        workbook.remove_sheet("销售数据")
        self.assertEqual(workbook.named_ranges, ())
        with self.assertRaises(KeyError):
            workbook.named_range("SalesAmount")

    def test_named_range_and_table_roundtrip_through_xlsx(self):
        """功能：验证命名区域及数据表的 XLSX 部件和读取往返。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言关系、内容类型以及数据表显示选项完整保留。
        """
        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "table.xlsx"
            workbook = Workbook()
            sheet = workbook.add_sheet("销售明细")
            sheet.append_rows([["名称", "金额"], ["甲", 10], ["乙", 20]])
            workbook.add_named_range("Amounts", sheet.range("B2:B3"))
            table = sheet.add_table(
                "A1:B3",
                name="SalesTable",
                style="TableStyleMedium9",
                show_row_stripes=False,
                show_column_stripes=True,
            )
            with self.assertRaises(ValueError):
                sheet.add_table("D1:D2", name="A1")
            self.assertIs(sheet.table("salestable"), table)
            workbook.save(filename)

            with zipfile.ZipFile(filename) as package:
                members = set(package.namelist())
                self.assertIn("xl/tables/table1.xml", members)
                self.assertIn("xl/worksheets/_rels/sheet1.xml.rels", members)
                table_root = ET.fromstring(package.read("xl/tables/table1.xml"))
                self.assertEqual(table_root.attrib["name"], "SalesTable")
                self.assertEqual(table_root.attrib["ref"], "A1:B3")

            loaded = Workbook.load(filename)
            loaded_sheet = loaded.sheet("销售明细")
            loaded_named = loaded.named_range("Amounts")
            loaded_table = loaded_sheet.table("SalesTable")
            self.assertEqual(loaded_named.range.address, "B2:B3")
            self.assertEqual(loaded_table.range.address, "A1:B3")
            self.assertEqual(loaded_table.style, "TableStyleMedium9")
            self.assertFalse(loaded_table.show_row_stripes)
            self.assertTrue(loaded_table.show_column_stripes)


if __name__ == "__main__":
    unittest.main()
