"""Workbook.load、XLS 读写和 XLSX 样式往返测试。"""

import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from excelkit import (
    Alignment,
    Border,
    Fill,
    Font,
    HeaderFooter,
    PageMargins,
    Side,
    Style,
    Workbook,
)
from excelkit.errors import InvalidFileError


class LoadAndFormatTests(unittest.TestCase):
    """验证支持格式读取以及 XLS/XLSX 的数据和样式往返。"""

    def setUp(self):
        """功能：为每项文件测试创建临时目录。

        使用方法：由 unittest 在每个测试方法前自动调用。
        参数：无。
        返回：无；临时目录对象和路径保存到实例属性。
        """
        self.temporary = tempfile.TemporaryDirectory()
        self.directory = Path(self.temporary.name)

    def tearDown(self):
        """功能：删除当前测试创建的临时文件和目录。

        使用方法：由 unittest 在每个测试方法后自动调用。
        参数：无。
        返回：无。
        """
        self.temporary.cleanup()

    def test_xlsx_roundtrip_preserves_formula_datetime_and_style(self):
        """功能：验证 XLSX 加载可恢复公式、日期时间和基础样式。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        filename = self.directory / "roundtrip.xlsx"
        style = Style(
            font=Font(name="Arial", size=13, bold=True, color="336699"),
            fill=Fill("FFF2CC"),
            border=Border(bottom=Side("thin", "000000")),
            alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
            number_format="0.00",
        )
        workbook = Workbook()
        sheet = workbook.add_sheet("数据")
        sheet.label_color = "4472C4"
        sheet["A1"] = "标题"
        sheet["A1"].style = style
        sheet["B2"] = "#2026/8/1 12:33"
        sheet["C3"].formula = "=SUM(1,2)"
        workbook.save(filename)

        loaded = Workbook.load(filename)
        self.assertEqual(loaded.active["A1"].value, "标题")
        self.assertEqual(loaded.active["A1"].style, style)
        self.assertEqual(loaded.active["B2"].value, datetime(2026, 8, 1, 12, 33))
        self.assertEqual(loaded.active["C3"].formula, "=SUM(1,2)")
        self.assertEqual(loaded.active.label_color, "FF4472C4")

    def test_xls_roundtrip_preserves_values_dates_and_basic_style(self):
        """功能：验证 XLS 写出、读取、中文、日期时间和基础样式映射。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        filename = self.directory / "legacy.xls"
        style = Style(
            font=Font(bold=True, color="FF0000"),
            fill=Fill("FFFF00"),
            border=Border(left=Side("thin", "000000")),
            alignment=Alignment(horizontal="center", wrap_text=True),
            number_format="0.00",
        )
        workbook = Workbook()
        sheet = workbook.add_sheet("成绩")
        sheet.append(["姓名", "日期", "成绩"])
        sheet.append(["张三", "#2026-8-1 12:33", 95.5])
        sheet["A1"].style = style
        workbook.save(filename)

        self.assertEqual(filename.read_bytes()[:8], b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
        loaded = Workbook.load(filename)
        self.assertEqual(loaded.active.values[0], ["姓名", "日期", "成绩"])
        self.assertEqual(loaded.active["A2"].value, "张三")
        self.assertEqual(loaded.active["B2"].value, datetime(2026, 8, 1, 12, 33))
        self.assertEqual(loaded.active["C2"].value, 95.5)
        loaded_style = loaded.active["A1"].style
        self.assertEqual(loaded_style.font, style.font)
        self.assertEqual(loaded_style.fill, style.fill)
        self.assertEqual(loaded_style.border, style.border)
        self.assertEqual(loaded_style.alignment.horizontal, "center")
        self.assertTrue(loaded_style.alignment.wrap_text)
        self.assertEqual(loaded_style.number_format, "0.00")

    def test_xlsx_roundtrip_preserves_layout_and_print_settings(self):
        """功能：验证XLSX完整往返合并、尺寸、视图、筛选和打印设置。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言工作表和workbook定义名称中的页面设置均恢复。
        """
        filename = self.directory / "layout.xlsx"
        workbook = Workbook()
        sheet = workbook.add_sheet("销售 报表")
        sheet["A1"] = "标题"
        sheet.range("A1:C1").merge()
        sheet.row(0).height = 28
        sheet.row(4).hidden = True
        sheet.column(1).width = 20
        sheet.column(3).hidden = True
        sheet.freeze = "B2"
        sheet.filter_range = "A2:F100"
        sheet.show_gridlines = False
        page = sheet.page
        page.orientation = "landscape"
        page.paper_size = "A3"
        page.fit(width=1)
        page.area = "A1:F100"
        page.repeat_rows = (0, 1)
        page.repeat_columns = (0, 0)
        page.margins = PageMargins(1.5, 1.6, 2, 2.1, 0.8, 0.9)
        page.center_horizontal = True
        page.print_gridlines = True
        page.print_headings = True
        page.black_and_white = True
        page.draft = True
        page.order = "over_then_down"
        page.first_page_number = 3
        page.header = HeaderFooter(
            left="研发 && 销售",
            center="&B报表&B",
            right="&D &T",
        )
        page.footer = HeaderFooter(center="第 &P 页，共 &N 页")
        workbook.save(filename)

        loaded = Workbook.load(filename)
        result = loaded.active
        self.assertEqual([area.address for area in result.merged_ranges], ["A1:C1"])
        self.assertEqual(result.row(0).height, 28)
        self.assertTrue(result.row(4).hidden)
        self.assertEqual(result.column(1).width, 20)
        self.assertTrue(result.column(3).hidden)
        self.assertEqual(result.freeze, "B2")
        self.assertEqual(result.filter_range, "A2:F100")
        self.assertFalse(result.show_gridlines)
        self.assertEqual(result.page.orientation, "landscape")
        self.assertEqual(result.page.paper_size, "A3")
        self.assertEqual((result.page._fit_width, result.page._fit_height), (1, None))
        self.assertEqual(result.page.area, "A1:F100")
        self.assertEqual(result.page.repeat_rows, (0, 1))
        self.assertEqual(result.page.repeat_columns, (0, 0))
        self.assertAlmostEqual(result.page.margins.left, 1.5)
        self.assertTrue(result.page.center_horizontal)
        self.assertTrue(result.page.print_gridlines)
        self.assertTrue(result.page.print_headings)
        self.assertTrue(result.page.black_and_white)
        self.assertTrue(result.page.draft)
        self.assertEqual(result.page.order, "over_then_down")
        self.assertEqual(result.page.first_page_number, 3)
        self.assertEqual(result.page.header.left, "研发 && 销售")
        self.assertEqual(result.page.header.center, "&B报表&B")
        self.assertEqual(result.page.header.right, "&D &T")
        self.assertEqual(result.page.footer.center, "第 &P 页，共 &N 页")

    def test_xls_roundtrip_preserves_supported_layout_subset(self):
        """功能：验证XLS往返其格式后端可读取的合并、尺寸和冻结窗格。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；页面打印设置写出不报错，xlrd可读取部分正确恢复。
        """
        filename = self.directory / "legacy-layout.xls"
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "标题"
        sheet.range("A1:C1").merge()
        sheet.row(0).height = 28
        sheet.row(3).hidden = True
        sheet.column(1).width = 20
        sheet.column(2).hidden = True
        sheet.freeze = "B2"
        sheet.page.orientation = "landscape"
        sheet.page.fit(width=1)
        sheet.page.header = HeaderFooter(center="报表 &P/&N")
        workbook.save(filename)

        loaded = Workbook.load(filename)
        result = loaded.active
        self.assertEqual([area.address for area in result.merged_ranges], ["A1:C1"])
        self.assertAlmostEqual(result.row(0).height or 0, 28, delta=1)
        self.assertTrue(result.row(3).hidden)
        self.assertAlmostEqual(result.column(1).width or 0, 20, delta=0.1)
        self.assertTrue(result.column(2).hidden)
        self.assertEqual(result.freeze, "B2")

    def test_csv_and_tsv_loading(self):
        """功能：验证常用分隔文本格式通过 Workbook.load 读取。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        csv_file = self.directory / "data.csv"
        tsv_file = self.directory / "data.tsv"
        csv_file.write_text("姓名,日期\n张三,#2026-8-1\n", encoding="utf-8")
        tsv_file.write_text("姓名\t成绩\n李四\t92\n", encoding="utf-8")
        self.assertEqual(Workbook.load(csv_file).active["B2"].value, date(2026, 8, 1))
        self.assertEqual(Workbook.load(tsv_file).active.values, [["姓名", "成绩"], ["李四", "92"]])

    def test_invalid_file_format_raises_excelkit_error(self):
        """功能：验证未知文件格式得到统一 InvalidFileError。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        filename = self.directory / "data.bin"
        filename.write_bytes(b"not an excel workbook")
        with self.assertRaises(InvalidFileError):
            Workbook.load(filename)

    def test_save_rejects_unsupported_extensions_without_mutating_workbook(self):
        """功能：验证保存只接受 XLSX 和 XLS，错误扩展名不会创建工作表或文件。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言错误类型、文件状态和空工作簿状态。
        """
        workbook = Workbook()
        for suffix in (".csv", ".xlsm", ".bin", ""):
            filename = self.directory / f"invalid{suffix}"
            with self.subTest(suffix=suffix), self.assertRaises(InvalidFileError):
                workbook.save(filename)
            self.assertFalse(filename.exists())
            self.assertEqual(workbook.sheets, ())


if __name__ == "__main__":
    unittest.main()
