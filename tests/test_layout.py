"""工作表管理、结构、页面布局与打印设置对象测试。"""

import unittest

from excelkit import Workbook
from excelkit.page_setup import HeaderFooter, PageMargins


class WorkbookSheetManagementTests(unittest.TestCase):
    """验证工作表删除、移动和完整复制。"""

    def test_remove_and_move_sheet_use_zero_based_final_index(self):
        """功能：验证删除和移动使用名称或0-based索引并返回当前工作簿。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言顺序、链式返回及错误边界。
        """
        workbook = Workbook()
        for name in ("甲", "乙", "丙"):
            workbook.add_sheet(name)
        self.assertIs(workbook.move_sheet("丙", 0), workbook)
        self.assertEqual([sheet.name for sheet in workbook.sheets], ["丙", "甲", "乙"])
        self.assertIs(workbook.remove_sheet(1), workbook)
        self.assertEqual([sheet.name for sheet in workbook.sheets], ["丙", "乙"])
        with self.assertRaises(IndexError):
            workbook.move_sheet("乙", 2)
        with self.assertRaises(TypeError):
            workbook.move_sheet("乙", True)

    def test_copy_sheet_preserves_independent_layout_and_page_settings(self):
        """功能：验证复制工作表保留全部状态且可独立修改。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言值、样式之外的结构和打印设置均复制且互不共享。
        """
        workbook = Workbook()
        source = workbook.add_sheet("模板")
        source["A1"] = "标题"
        source.range("A1:C1").merge()
        source.row(0).height = 28
        source.column(1).width = 20
        source.freeze_panes = "B2"
        source.auto_filter.range = "A1:C3"
        source.show_gridlines = False
        source.page.orientation = "landscape"
        source.page.fit(width=1)
        source.page.header = HeaderFooter(center="报表")

        copied = workbook.copy_sheet("模板", "副本")
        self.assertEqual(copied.values, source.values)
        self.assertEqual([area.address for area in copied.merged_ranges], ["A1:C1"])
        self.assertEqual(copied.row(0).height, 28)
        self.assertEqual(copied.column(1).width, 20)
        self.assertEqual(copied.freeze_panes, "B2")
        self.assertEqual(copied.auto_filter.range, "A1:C3")
        self.assertFalse(copied.show_gridlines)
        self.assertEqual(copied.page.orientation, "landscape")
        self.assertEqual(copied.page.header.center, "报表")

        copied.row(0).height = 40
        copied.page.header = HeaderFooter(center="副本")
        self.assertEqual(source.row(0).height, 28)
        self.assertEqual(source.page.header.center, "报表")

        source["D4"] = {"items": [1]}
        independent = workbook.copy_sheet("模板", "独立值副本")
        independent["D4"].value["items"].append(2)
        self.assertEqual(source["D4"].value, {"items": [1]})


class WorksheetLayoutTests(unittest.TestCase):
    """验证合并区域、行列尺寸、冻结、筛选和网格线。"""

    def setUp(self):
        """功能：为每项布局测试创建独立空工作表。

        使用方法：由 unittest 在测试方法前自动调用。
        参数：无。
        返回：无。
        """
        self.worksheet = Workbook().active

    def test_merge_unmerge_and_non_anchor_write_rules(self):
        """功能：验证合并幂等、重叠检查、写入锚点和取消合并。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言合并区域的安全写入规则。
        """
        area = self.worksheet.range("A1:C1")
        self.assertIs(area.merge(), area)
        area.merge()
        self.assertEqual(area.address, "A1:C1")
        self.assertEqual([item.address for item in self.worksheet.merged_ranges], ["A1:C1"])
        self.worksheet["A1"] = "标题"
        with self.assertRaises(ValueError):
            self.worksheet["B1"] = "非法"
        with self.assertRaises(ValueError):
            self.worksheet.range("B1:D1").merge()
        self.assertIs(area.unmerge(), area)
        self.worksheet["B1"] = "恢复"
        self.assertEqual(self.worksheet["B1"].value, "恢复")

    def test_merge_rejects_existing_secondary_content_atomically(self):
        """功能：验证合并不会静默删除非左上角值或公式。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；失败后内容和合并集合保持不变。
        """
        self.worksheet["B1"] = "保留"
        with self.assertRaises(ValueError):
            self.worksheet.range("A1:C1").merge()
        self.assertEqual(self.worksheet["B1"].value, "保留")
        self.assertEqual(self.worksheet.merged_ranges, ())

    def test_dimensions_freeze_filter_and_gridlines(self):
        """功能：验证0-based行列尺寸以及工作表显示属性。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言规范化、恢复默认及类型验证。
        """
        row = self.worksheet.row(0)
        column = self.worksheet.column(1)
        row.height = 28
        row.hidden = True
        column.width = 20
        column.hidden = True
        self.assertEqual((row.index, row.height, row.hidden), (0, 28.0, True))
        self.assertEqual((column.index, column.width, column.hidden), (1, 20.0, True))
        self.assertIs(self.worksheet.row(0), row)
        self.assertIs(self.worksheet.column(1), column)

        self.worksheet.freeze_panes = "b2"
        self.worksheet.auto_filter.range = "a1:c20"
        self.worksheet.show_gridlines = False
        self.assertEqual(self.worksheet.freeze_panes, "B2")
        self.assertEqual(self.worksheet.auto_filter.range, "A1:C20")
        self.assertFalse(self.worksheet.show_gridlines)
        self.worksheet.freeze_panes = "A1"
        self.worksheet.auto_filter.range = None
        self.assertIsNone(self.worksheet.freeze_panes)
        self.assertIsNone(self.worksheet.auto_filter.range)


class PageSettingsTests(unittest.TestCase):
    """验证页面值对象、适应页数和打印选项。"""

    def test_page_settings_validate_and_fit_is_one_operation(self):
        """功能：验证页面枚举、缩放互斥及一次性fit配置。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言合法赋值和错误输入。
        """
        page = Workbook().active.page
        page.orientation = "landscape"
        page.paper_size = "letter"
        page.fit(width=1)
        self.assertEqual(page.orientation, "landscape")
        self.assertEqual(page.paper_size, "Letter")
        self.assertIsNone(page.scale)
        self.assertEqual((page._fit_width, page._fit_height), (1, None))
        page.scale = 90
        self.assertEqual(page.scale, 90)
        self.assertEqual((page._fit_width, page._fit_height), (None, None))
        with self.assertRaises(ValueError):
            page.fit(width=None, height=None)
        with self.assertRaises(ValueError):
            page.scale = 5
        with self.assertRaises(ValueError):
            page.scale = None

    def test_print_area_titles_margins_headers_and_flags(self):
        """功能：验证打印区域、0-based重复标题、厘米边距和页眉页脚。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言页面打印属性均保存规范值。
        """
        page = Workbook().active.page
        page.print_area = "a1:f100"
        page.repeat_rows = (0, 1)
        page.repeat_columns = (0, 0)
        page.margins = PageMargins(1.5, 1.5, 2, 2, 0.8, 0.8)
        page.header = HeaderFooter(left="ExcelKit", center="报表", right="&D")
        page.footer = HeaderFooter(center="第 &P 页，共 &N 页")
        page.center_horizontal = True
        page.print_gridlines = True
        page.print_headings = True
        page.first_page_number = 3
        self.assertEqual(page.print_area, "A1:F100")
        self.assertEqual(page.repeat_rows, (0, 1))
        self.assertEqual(page.repeat_columns, (0, 0))
        self.assertEqual(page.margins.left, 1.5)
        self.assertEqual(page.footer.center, "第 &P 页，共 &N 页")
        self.assertTrue(page.center_horizontal)
        self.assertTrue(page.print_gridlines)
        self.assertTrue(page.print_headings)
        self.assertEqual(page.first_page_number, 3)


if __name__ == "__main__":
    unittest.main()
