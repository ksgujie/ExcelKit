"""ExcelKit 0.5.0 新增工作表结构和 XLSX 功能回归测试。"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from excelkit import Workbook
from excelkit.hyperlink import Hyperlink


class StructureEditingTests(unittest.TestCase):
    """验证行列编辑及其相关元数据同步。"""

    def test_row_and_column_edits_update_references_and_rules(self) -> None:
        """功能：验证插入、删除会同步公式、验证规则和条件格式。"""
        workbook = Workbook()
        source = workbook.add_sheet("源数据")
        report = workbook.add_sheet("报表")
        source["A1"] = 10
        source["B1"] = 20
        report["A1"].formula = "='源数据'!A1+'源数据'!B1"
        source.add_validation("A1:B2", kind="whole", operator="greaterThan", formula1="0")
        source.add_conditional_format("A1:B2", operator="greaterThan", formula="10", fill="FFC7CE")

        source.insert_rows(0)
        self.assertEqual(report["A1"].formula, "='源数据'!A2+'源数据'!B2")
        self.assertEqual(source.validations[0].range, "A2:B3")
        self.assertEqual(source.conditional_formats[0].range, "A2:B3")

        source.delete_columns(0)
        self.assertEqual(report["A1"].formula, "=#REF!+'源数据'!A2")
        self.assertEqual(source.validations[0].range, "A2:A3")


class Release050RoundTripTests(unittest.TestCase):
    """验证 0.5.0 的核心 XLSX 往返能力。"""

    def test_xlsx_roundtrip_for_new_features(self) -> None:
        """功能：验证超链接、属性、表格、筛选、保护和规则均可读取回来。"""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["姓名", "成绩", "状态"])
        worksheet.append(["张三", 95, "通过"])
        worksheet["A2"].hyperlink = Hyperlink(
            target="https://example.com", display="张三主页", tooltip="打开链接"
        )
        worksheet["A3"].hyperlink = Hyperlink(location="Sheet1!A1", display="返回")
        worksheet.add_validation("C2:C20", values=["通过", "不通过"])
        worksheet.add_conditional_format(
            "B2:B20", operator="greaterThan", formula="90", fill="FFC7CE", font="006100"
        )
        worksheet.auto_filter.range = "A1:C20"
        worksheet.auto_filter.set(2, ["通过"])
        worksheet.protection.enabled = True
        worksheet.protection.password = "ABCD"
        workbook.protection.enabled = True
        workbook.properties.title = "成绩单"
        workbook.properties.author = "ExcelKit"
        workbook.properties.category = "测试"
        table = worksheet.add_table("A1:C2", name="Scores")
        table.append(["李四", 88, "不通过"])
        table.show_totals = True
        table.set_total("成绩", "average")

        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "release-050.xlsx"
            workbook.save(filename)
            loaded = Workbook.load(filename)

        loaded_sheet = loaded.active
        self.assertEqual(loaded.properties.title, "成绩单")
        self.assertEqual(loaded.properties.author, "ExcelKit")
        self.assertEqual(loaded.properties.category, "测试")
        self.assertEqual(loaded_sheet["A2"].hyperlink.target, "https://example.com")
        self.assertEqual(loaded_sheet["A3"].hyperlink.location, "Sheet1!A1")
        self.assertEqual(loaded_sheet.validations[0].values, ("通过", "不通过"))
        self.assertEqual(loaded_sheet.conditional_formats[0].fill, "FFFFC7CE")
        self.assertEqual(loaded_sheet.conditional_formats[0].font, "FF006100")
        self.assertEqual(loaded_sheet.auto_filter.filters[2], ("通过",))
        self.assertTrue(loaded_sheet.protection.enabled)
        self.assertTrue(loaded.protection.enabled)
        loaded_table = loaded_sheet.table("Scores")
        self.assertEqual(loaded_table.columns, ("姓名", "成绩", "状态"))
        self.assertTrue(loaded_table.show_totals)
        self.assertEqual(loaded_table.totals, {"成绩": "average"})

    def test_delimited_options_and_formula_relationships(self) -> None:
        """功能：验证 CSV 参数、表头元数据及 Cell 依赖和被依赖查询。"""
        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "scores.csv"
            filename.write_text("姓名;成绩\n张三;95\n", encoding="utf-8")
            workbook = Workbook.load(filename, delimiter=";", has_header=True)
        self.assertEqual(workbook.active.headers, ("姓名", "成绩"))
        self.assertEqual(workbook.active["B2"].value, "95")

        worksheet = Workbook().active
        worksheet["A1"] = 10
        worksheet["B1"].formula = "=A1*2"
        self.assertEqual([cell.address for cell in worksheet["B1"].dependencies], ["A1"])
        self.assertEqual([cell.address for cell in worksheet["A1"].dependents], ["B1"])


if __name__ == "__main__":
    unittest.main()
