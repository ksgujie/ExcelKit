"""工作簿、工作表、单元格、公式、区域和追加功能的回归测试。"""

import unittest
from datetime import date, datetime

from excelkit import Alignment, Fill, Font, Style, Workbook
from excelkit.errors import InvalidWorksheetNameError
from excelkit.storage import ValueStore


class WorkbookTests(unittest.TestCase):
    """验证工作簿和 0-based 工作表索引。"""

    def test_active_is_lazy_and_stable(self):
        """功能：验证 active 延迟创建 Sheet1 且重复访问对象稳定。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = Workbook()
        self.assertEqual(workbook.sheets, ())
        active = workbook.active
        self.assertEqual(active.label, "Sheet1")
        self.assertIs(workbook.active, active)

    def test_sheet_supports_name_and_zero_based_index(self):
        """功能：验证 sheet 的名称查询和 0-based 正负索引查询。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = Workbook()
        first = workbook.add_sheet("第一张")
        second = workbook.add_sheet("第二张")
        self.assertIs(workbook.sheet("第一张"), first)
        self.assertIs(workbook.sheet(0), first)
        self.assertIs(workbook.sheet(1), second)
        self.assertIs(workbook.sheet(-1), second)
        with self.assertRaises(KeyError):
            workbook.sheet("不存在")
        with self.assertRaises(IndexError):
            workbook.sheet(2)
        with self.assertRaises(TypeError):
            workbook.sheet(True)

    def test_sheet_names_are_validated(self):
        """功能：验证工作表名称长度、字符和重复规则。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = Workbook()
        workbook.add_sheet("Data")
        with self.assertRaises(ValueError):
            workbook.add_sheet("data")
        for name in ("", "x" * 32, "bad/name", "bad\x00name", 123):
            with self.subTest(name=name), self.assertRaises(InvalidWorksheetNameError):
                workbook.add_sheet(name)

    def test_worksheet_label_setter_updates_workbook_index_atomically(self):
        """功能：验证工作表重命名同步名称索引并在失败时保持原名称。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = Workbook()
        first = workbook.add_sheet("原名称")
        second = workbook.add_sheet("其他")
        first.label = "新名称"
        self.assertEqual(first.label, "新名称")
        self.assertIs(workbook.sheet("新名称"), first)
        self.assertEqual(workbook.sheets, (first, second))
        with self.assertRaises(KeyError):
            workbook.sheet("原名称")
        with self.assertRaises(ValueError):
            first.label = "其他"
        self.assertEqual(first.label, "新名称")
        with self.assertRaises(InvalidWorksheetNameError):
            first.label = ""
        self.assertEqual(first.label, "新名称")
        first.label = "NewName"
        first.label = "newname"
        self.assertEqual(first.label, "newname")
        self.assertIs(workbook.sheet("newname"), first)
        self.assertFalse(hasattr(first, "name"))
        self.assertFalse(hasattr(first, "tab_name"))


class WorksheetTests(unittest.TestCase):
    """验证工作表单元格、公式、区域和追加 API。"""

    def setUp(self):
        """功能：为每项测试创建独立空工作表。

        使用方法：由 unittest 在每个测试方法前自动调用。
        参数：无。
        返回：无；工作表保存到 ``self.worksheet``。
        """
        self.worksheet = Workbook().active

    def test_cell_has_one_method_and_zero_based_coordinates(self):
        """功能：验证 cell 同时处理 A1 和数字索引，且 cell_at 已彻底删除。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        from_address = self.worksheet.cell("C8")
        from_indexes = self.worksheet.cell(7, 2)
        self.assertEqual((from_address.row, from_address.column), (7, 2))
        self.assertEqual(from_indexes.address, "C8")
        self.assertFalse(hasattr(self.worksheet, "cell_at"))
        with self.assertRaises(TypeError):
            self.worksheet.cell("C8", 2)
        with self.assertRaises(TypeError):
            self.worksheet.cell(7)

    def test_max_indexes_track_touched_positions(self):
        """功能：验证空表最大索引为 -1，写入后按 0-based 位置增长。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.assertEqual((self.worksheet.max_row, self.worksheet.max_column), (-1, -1))
        self.worksheet["F10"] = None
        self.assertEqual((self.worksheet.max_row, self.worksheet.max_column), (9, 5))
        self.worksheet["A1"] = "值"
        self.assertEqual((self.worksheet.max_row, self.worksheet.max_column), (9, 5))

    def test_formula_and_value_are_mutually_exclusive(self):
        """功能：验证公式标准化以及普通值与公式相互覆盖。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        cell = self.worksheet["C2"]
        cell.value = 10
        cell.formula = "SUM(A2:B2)"
        self.assertEqual(cell.formula, "=SUM(A2:B2)")
        self.assertIsNone(cell.value)
        cell.value = 100
        self.assertEqual(cell.value, 100)
        self.assertIsNone(cell.formula)
        for invalid in ("", "=", "   ", None, 123):
            with self.subTest(invalid=invalid), self.assertRaises(TypeError):
                cell.formula = invalid

    def test_append_and_append_rows(self):
        """功能：验证空表追加、连续多行追加和空行不推进最大索引。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.assertIs(self.worksheet.append([]), self.worksheet)
        self.assertEqual(self.worksheet.max_row, -1)
        self.worksheet.append(["姓名", "成绩"])
        self.worksheet.append_rows((("张三", 95), ("李四", 92)))
        self.assertEqual(self.worksheet.range("A1:B3").values, [
            ["姓名", "成绩"],
            ["张三", 95],
            ["李四", 92],
        ])
        self.assertEqual((self.worksheet.max_row, self.worksheet.max_column), (2, 1))
        self.assertFalse(hasattr(self.worksheet, "append_many"))

    def test_range_indexes_values_and_formula_overwrite(self):
        """功能：验证 Range 边界为 0-based，普通值读取不返回公式。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        area = self.worksheet.range("B3:D4")
        self.assertEqual(
            (area.min_row, area.min_column, area.max_row, area.max_column),
            (2, 1, 3, 3),
        )
        self.worksheet["B3"].formula = "=1+1"
        self.assertEqual(area.values[0][0], None)
        result = area.set_values([[1, 2, 3], [4, 5, 6]])
        self.assertIs(result, area)
        self.assertEqual(area.values, [[1, 2, 3], [4, 5, 6]])
        self.assertIsNone(self.worksheet["B3"].formula)

    def test_range_shape_validation_is_atomic(self):
        """功能：验证区域尺寸错误不会造成普通值部分覆盖。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        area = self.worksheet.range("A1:B2")
        area.set_values([[1, 2], [3, 4]])
        with self.assertRaises(ValueError):
            area.set_values([[9, 8], [7]])
        self.assertEqual(area.values, [[1, 2], [3, 4]])

    def test_date_and_datetime_literals_use_one_normalizer(self):
        """功能：验证日期及日期时间字面量通过全部普通值入口统一转换。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.worksheet["A1"] = "#2026-8-1"
        self.worksheet["A2"] = "#2026/8/1 12:33"
        self.worksheet.range("A3:A4").set_values(
            [["#2026-8-1 1:02:03"], ["普通文本"]]
        )
        self.assertEqual(self.worksheet["A1"].value, date(2026, 8, 1))
        self.assertEqual(
            self.worksheet["A2"].value, datetime(2026, 8, 1, 12, 33)
        )
        self.assertEqual(
            self.worksheet["A3"].value, datetime(2026, 8, 1, 1, 2, 3)
        )
        self.assertEqual(self.worksheet["A4"].value, "普通文本")
        self.worksheet["B1"] = "保留"
        with self.assertRaises(ValueError):
            self.worksheet["B1"] = "#2026-2-30 12:00"
        self.assertEqual(self.worksheet["B1"].value, "保留")

    def test_worksheet_values_and_cell_style(self):
        """功能：验证工作表全部值属性和不可变单元格样式。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.assertEqual(self.worksheet.values, [])
        self.worksheet["B2"] = 2
        self.worksheet["C1"].formula = "=1+1"
        self.assertEqual(
            self.worksheet.values, [[None, None, None], [None, 2, None]]
        )
        style = Style(
            font=Font(bold=True, color="FF0000"),
            fill=Fill("FFFF00"),
            alignment=Alignment(horizontal="center", wrap_text=True),
            number_format="0.00",
        )
        self.worksheet["D4"].style = style
        self.assertEqual(self.worksheet["D4"].style, style)
        self.assertEqual((self.worksheet.max_row, self.worksheet.max_column), (3, 3))
        with self.assertRaises(TypeError):
            self.worksheet["A1"].style = {"bold": True}

    def test_worksheet_label_color_normalizes_and_clears(self):
        """功能：验证工作表标签颜色使用统一 RGB/ARGB 规则并支持清除。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.assertIsNone(self.worksheet.label_color)
        self.worksheet.label_color = "4472c4"
        self.assertEqual(self.worksheet.label_color, "FF4472C4")
        self.worksheet.label_color = "804472C4"
        self.assertEqual(self.worksheet.label_color, "804472C4")
        with self.assertRaises(ValueError):
            self.worksheet.label_color = "blue"
        self.assertEqual(self.worksheet.label_color, "804472C4")
        self.worksheet.label_color = None
        self.assertIsNone(self.worksheet.label_color)
        self.assertFalse(hasattr(self.worksheet, "tab_color"))


class ValueStoreTests(unittest.TestCase):
    """验证 storage 模块的 0-based 稀疏普通值存储。"""

    def test_value_store_get_set_and_items(self):
        """功能：验证普通值设置、删除、读取和行优先迭代。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        store = ValueStore()
        store.set(1, 2, "后")
        store.set(0, 1, "前")
        self.assertEqual(store.get(0, 1), "前")
        self.assertEqual(list(store.items()), [((0, 1), "前"), ((1, 2), "后")])
        store.set(0, 1, None)
        self.assertIsNone(store.get(0, 1))


if __name__ == "__main__":
    unittest.main()
