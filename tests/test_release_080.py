"""0.8.2 业务报表生产力 API 回归测试。"""

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
            worksheet.range("A1:B4").to_records(),
            [{"订单号": "A-1", "金额": 20}, {"订单号": "A-2", "金额": 30}, {"订单号": "A-3", "金额": 40}],
        )
        self.assertEqual(
            worksheet.range("A2:B4").to_records(headers=False)[-1],
            {"Column1": "A-3", "Column2": 40},
        )
        self.assertEqual(table.to_records()[-1], {"订单号": "A-3", "金额": 40})
        self.assertTrue(table.show_totals)
        self.assertEqual(table.totals["金额"], TotalFunction.SUM)
        self.assertEqual(worksheet.freeze_panes, "A2")
        self.assertIsNotNone(worksheet.column(0).width)

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "totals.xlsx"
            worksheet._workbook.save(output)
            restored = Workbook.load(output).active.table("Orders")
        self.assertTrue(restored.show_totals)
        self.assertEqual(restored.to_records()[-1], {"订单号": "A-3", "金额": 40})

    def test_range_and_worksheet_to_records(self) -> None:
        """功能：验证外部字段行、显式字段和整表记录转换使用统一有效值。"""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet["A1"] = "学生成绩表"
        worksheet.range("A2:C2").set_values([["姓名", "年龄", "成绩"]])
        worksheet.range("A5:C6").set_values([
            ["张三", 18, 95],
            ["李四", 19, 88],
        ])
        worksheet["A8"] = "王五"
        worksheet["B8"] = 20
        worksheet["C8"].formula = "=B8+72"
        worksheet.range("A9:C9").set_values([["", None, ""]])
        workbook.calculate()

        self.assertEqual(
            worksheet.range("A5:C6").to_records(header_row=1),
            [
                {"姓名": "张三", "年龄": 18, "成绩": 95},
                {"姓名": "李四", "年龄": 19, "成绩": 88},
            ],
        )
        self.assertEqual(
            worksheet.range("A5:C6").to_records(headers=["name", "age", "score"])[0],
            {"name": "张三", "age": 18, "score": 95},
        )
        self.assertEqual(
            worksheet.to_records(header_row=1)[-1],
            {"姓名": "王五", "年龄": 20, "成绩": 92},
        )
        self.assertEqual(len(worksheet.to_records(header_row=1)), 3)
        self.assertEqual(worksheet["C8"].value, 92)
        self.assertEqual(worksheet.values[7][2], 92)
        self.assertEqual(worksheet.range("B8:C8").values, [[20, 92]])
        with self.assertRaises(ValueError):
            worksheet.range("A5:C8").to_records(header_row=5)
        with self.assertRaises(ValueError):
            worksheet.range("A5:C6").to_records(headers=["姓名"])
        with self.assertRaises(ValueError):
            worksheet.range("A5:C6").to_records(headers=False, header_row=1)
        with self.assertRaises(ValueError):
            worksheet.range("A5:C6").to_records(
                headers=["name", "age", "score"], header_row=1
            )

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
        worksheet["D2"].formula = "=B2*C2"
        worksheet.range("D2:D2").auto_fill("D2:D3")

        self.assertEqual(worksheet.range("A1:C3").values, [["编号", "数量", "单价"], ["A", 2, 3], ["B", 4, 5]])
        self.assertEqual(worksheet["D2"].formula, "=B2*C2")
        self.assertEqual(worksheet["D3"].formula, "=B3*C3")

    def test_xlsx_extensions(self) -> None:
        """功能：高级条件格式和分页符应全部写入 XLSX。"""
        workbook = Workbook()
        template = workbook.active
        template["A1"] = "{name}"
        template.add_color_scale("A1:A10")
        template.add_data_bar("B1:B10")
        template.add_icon_set("C1:C10", style=IconSet.THREE_ARROWS)
        template.add_horizontal_page_break(1)

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "report.xlsx"
            workbook.save(output)
            with zipfile.ZipFile(output) as package:
                xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn("colorScale", xml)
        self.assertIn("dataBar", xml)
        self.assertIn("iconSet", xml)
        self.assertIn("rowBreaks", xml)

    def test_outline_groups_and_removed_duplicate_apis(self) -> None:
        """功能：验证行列大纲可往返，并确保重复 API 已彻底删除。"""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet["A1"] = "分组"
        worksheet.group_rows(1, 3, collapsed=True)
        worksheet.group_rows(1, 3)
        worksheet.ungroup_rows(1, 3)
        self.assertTrue(worksheet.row(3).collapsed)
        self.assertTrue(worksheet.row(2).hidden)
        worksheet.group_columns(1, 2)

        self.assertEqual(worksheet.row(2).outline_level, 1)
        self.assertTrue(worksheet.row(2).hidden)
        self.assertTrue(worksheet.row(3).collapsed)
        self.assertEqual(worksheet.column(1).outline_level, 1)
        self.assertFalse(hasattr(worksheet, "auto_filter_range"))
        self.assertFalse(hasattr(worksheet, "fill_formula"))
        self.assertFalse(hasattr(worksheet, "read_records"))
        self.assertFalse(hasattr(worksheet["A1"], "cached_value"))
        self.assertFalse(hasattr(worksheet.range("A1:A1"), "clear_values"))
        self.assertFalse(hasattr(worksheet.auto_filter, "add"))
        self.assertFalse(hasattr(workbook, "render_many"))
        self.assertFalse(hasattr(workbook, "export_pages"))
        self.assertFalse(hasattr(worksheet.add_table("A1:A1", name="Audit"), "records"))

        with tempfile.TemporaryDirectory() as directory:
            xlsx = Path(directory) / "groups.xlsx"
            xls = Path(directory) / "groups.xls"
            workbook.save(xlsx)
            workbook.save(xls)
            restored = Workbook.load(xlsx).active
        self.assertEqual(restored.row(2).outline_level, 1)
        self.assertTrue(restored.row(3).collapsed)
        self.assertEqual(restored.column(2).outline_level, 1)

        worksheet.ungroup_rows(1, 3)
        worksheet.ungroup_columns(1, 2)
        self.assertEqual(worksheet.row(2).outline_level, 0)
        self.assertEqual(worksheet.column(1).outline_level, 0)


if __name__ == "__main__":
    unittest.main()
