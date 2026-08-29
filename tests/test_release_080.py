"""0.8.0 业务报表生产力 API 回归测试。"""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from excelkit import Workbook
from excelkit.autofill import AutoFillMode
from excelkit.conditional import IconSet
from excelkit.style import NumberFormat, ReportStyle
from excelkit.table import TotalFunction


class Release080Tests(unittest.TestCase):
    """验证新数据报表 API 的业务语义与 XLSX 写出结果。"""

    def test_write_read_table_and_totals(self) -> None:
        """功能：字典写入、读取和数据表汇总应使用一致字段顺序。"""
        worksheet = Workbook().active
        table = worksheet.write_table(
            0, 0,
            [{"订单号": "A-1", "金额": 20}, {"订单号": "A-2", "金额": 30}],
            name="Orders", freeze_header=True, auto_fit=True,
        )
        table.append_records([{"订单号": "A-3", "金额": 40}])
        table.set_total("金额", TotalFunction.SUM)

        self.assertEqual(
            worksheet.read_records("A1:B4"),
            [{"订单号": "A-1", "金额": 20}, {"订单号": "A-2", "金额": 30}, {"订单号": "A-3", "金额": 40}],
        )
        self.assertEqual(table.records[-1], {"订单号": "A-3", "金额": 40})
        self.assertTrue(table.show_totals)
        self.assertEqual(table.totals["金额"], TotalFunction.SUM)
        self.assertEqual(worksheet.freeze_panes, "A2")
        self.assertIsNotNone(worksheet.column(0).width)

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "totals.xlsx"
            worksheet._workbook.save(output)
            restored = Workbook.load(output).active.table("Orders")
        self.assertTrue(restored.show_totals)
        self.assertEqual(restored.records[-1], {"订单号": "A-3", "金额": 40})

    def test_formula_auto_fill_and_format_proxy(self) -> None:
        """功能：公式与数值自动填充、区域数字格式代理应生效。"""
        worksheet = Workbook().active
        worksheet.range("A1:A2").set_values([[1], [2]]).auto_fill(
            "A1:A5", mode=AutoFillMode.SERIES
        )
        worksheet["B1"].formula = "=A1*2"
        worksheet.range("B1:B1").auto_fill("B1:B5")
        worksheet.range("B1:B5").format.number = NumberFormat.CURRENCY
        worksheet.range("C1:C1").apply_style(ReportStyle.HEADER)

        self.assertEqual(worksheet.range("A1:A5").values, [[1], [2], [3], [4], [5]])
        self.assertEqual(worksheet["B5"].formula, "=A5*2")
        self.assertEqual(worksheet.range("B1:B5").format.number, NumberFormat.CURRENCY)
        self.assertTrue(worksheet["C1"].style.font.bold)

    def test_cleanup_and_batch_formula(self) -> None:
        """功能：去重、清空空白行和公式批量下拉应正确保留目标顺序。"""
        worksheet = Workbook().active
        worksheet.range("A1:C5").set_values([
            ["编号", "数量", "单价"], ["A", 2, 3], ["A", 2, 3], [None, None, None], ["B", 4, 5],
        ])
        self.assertEqual(worksheet.range("A1:C5").remove_duplicates([0], has_header=True), 1)
        self.assertEqual(worksheet.range("A1:C5").remove_blank_rows(), 2)
        worksheet.fill_formula("D2:D3", "=B2*C2")

        self.assertEqual(worksheet.range("A1:C3").values, [["编号", "数量", "单价"], ["A", 2, 3], ["B", 4, 5]])
        self.assertEqual(worksheet["D2"].formula, "=B2*C2")
        self.assertEqual(worksheet["D3"].formula, "=B3*C3")

    def test_pages_render_many_and_xlsx_extensions(self) -> None:
        """功能：分页、模板多表渲染、条件格式和分页符应全部写入 XLSX。"""
        workbook = Workbook()
        pages = workbook.export_pages(
            [{"编号": index} for index in range(5)], rows_per_sheet=2, sheet_name="数据"
        )
        self.assertEqual([sheet.name for sheet in pages], ["数据1", "数据2", "数据3"])

        template = workbook.add_sheet("模板")
        template["A1"] = "{name}"
        rendered = workbook.render_many(
            [{"name": "甲"}, {"name": "乙"}], sheet_name="模板", name_pattern="报表_{name}_{index}"
        )
        self.assertEqual([sheet["A1"].value for sheet in rendered], ["甲", "乙"])
        template.add_color_scale("A1:A10")
        template.add_data_bar("B1:B10")
        template.add_icon_set("C1:C10", style=IconSet.THREE_ARROWS)
        template.add_horizontal_page_break(1)

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "report.xlsx"
            workbook.save(output)
            with zipfile.ZipFile(output) as package:
                xml = package.read("xl/worksheets/sheet4.xml").decode("utf-8")
        self.assertIn("colorScale", xml)
        self.assertIn("dataBar", xml)
        self.assertIn("iconSet", xml)
        self.assertIn("rowBreaks", xml)


if __name__ == "__main__":
    unittest.main()
