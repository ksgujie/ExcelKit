"""工作簿、工作表、单元格、公式、区域和追加功能的回归测试。"""

import unittest
from datetime import date, datetime

from excelkit import Workbook
from excelkit.errors import InvalidWorksheetNameError
from excelkit.page_setup import PageSettings
from excelkit.storage import ValueStore
from excelkit.style import Alignment, Border, BorderSide, Fill, Font, Style


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
        self.assertEqual(active.name, "Sheet1")
        self.assertIs(workbook.active, active)

    def test_public_modules_are_categorized_and_constants_are_available(self):
        """功能：验证顶层只暴露核心对象，分类模块暴露样式和页面类型。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言顶层边界、样式线型常量和页面常量。
        """
        import excelkit

        self.assertFalse(hasattr(excelkit, "Border"))
        self.assertFalse(hasattr(excelkit, "PageSettings"))
        self.assertFalse(hasattr(excelkit, "HeaderFooter"))
        self.assertEqual(BorderSide(Border.DASH_DOT).style, "dashDot")
        self.assertEqual(Alignment.HORIZONTAL_CENTER, "center")
        self.assertEqual(PageSettings.A4, "A4")

    def test_sheet_supports_name_and_zero_based_index(self):
        """功能：验证 sheet 的标签查询和非负 0-based 索引查询。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = Workbook()
        first = workbook.add_sheet("第一张")
        second = workbook.add_sheet("第二张")
        self.assertIs(workbook.sheet("第一张"), first)
        self.assertIs(workbook.sheet("第一张".lower()), first)
        self.assertIs(workbook.sheet(0), first)
        self.assertIs(workbook.sheet(1), second)
        with self.assertRaises(IndexError):
            workbook.sheet(-1)
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
        data = workbook.add_sheet("Data")
        self.assertIs(workbook.sheet("data"), data)
        self.assertIs(workbook.sheet("DATA"), data)
        with self.assertRaises(ValueError):
            workbook.add_sheet("data")
        for name in ("", "x" * 32, "bad/name", "bad\x00name", 123):
            with self.subTest(name=name), self.assertRaises(InvalidWorksheetNameError):
                workbook.add_sheet(name)

    def test_worksheet_name_setter_updates_workbook_index_atomically(self):
        """功能：验证工作表重命名同步名称索引并在失败时保持原名称。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = Workbook()
        first = workbook.add_sheet("原名称")
        second = workbook.add_sheet("其他")
        first.name = "新名称"
        self.assertEqual(first.name, "新名称")
        self.assertIs(workbook.sheet("新名称"), first)
        self.assertEqual(workbook.sheets, (first, second))
        with self.assertRaises(KeyError):
            workbook.sheet("原名称")
        with self.assertRaises(ValueError):
            first.name = "其他"
        self.assertEqual(first.name, "新名称")
        with self.assertRaises(InvalidWorksheetNameError):
            first.name = ""
        self.assertEqual(first.name, "新名称")
        first.name = "NewName"
        self.assertIs(workbook.sheet("newname"), first)
        first.name = "newname"
        self.assertEqual(first.name, "newname")
        self.assertIs(workbook.sheet("newname"), first)
        self.assertTrue(hasattr(first, "name"))
        self.assertFalse(hasattr(first, "label"))
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
        self.assertEqual(from_address.index, (7, 2))
        self.assertEqual(self.worksheet["D3"].index, (2, 3))
        self.assertEqual(from_indexes.address, "C8")
        with self.assertRaises(AttributeError):
            from_address.index = (0, 0)
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
        cell.formula = "=1+1"
        cell.formula = None
        self.assertIsNone(cell.formula)
        for invalid in ("", "=", "   ", 123):
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

    def test_append_and_append_rows_validate_before_writing(self):
        """功能：验证单行和多行追加失败时不会留下部分写入数据。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言转换错误发生前后工作表数据完全一致。
        """
        self.worksheet.append(["原值"])
        before = self.worksheet.values
        with self.assertRaises(ValueError):
            self.worksheet.append([1, "#2026-99-1"])
        self.assertEqual(self.worksheet.values, before)
        with self.assertRaises(ValueError):
            self.worksheet.append_rows([[1], ["#2026-99-1"]])
        self.assertEqual(self.worksheet.values, before)

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

    def test_cell_set_value_and_as_date_chain(self):
        """功能：验证链式写值、显式日期转换、返回值及失败原子性。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言 ``set_value().as_date()`` 使用同一 Cell 并正确转换。
        """
        cell = self.worksheet.cell("a1")
        self.assertIs(cell.set_value("2026-8-1"), cell)
        self.assertEqual(cell.as_date(), date(2026, 8, 1))
        self.assertEqual(cell.value, date(2026, 8, 1))

        self.assertEqual(
            cell.set_value("2026/8/2 12:33").as_date(), date(2026, 8, 2)
        )
        self.assertEqual(cell.value, date(2026, 8, 2))
        self.assertEqual(
            cell.set_value(datetime(2026, 8, 3, 8, 30)).as_date(),
            date(2026, 8, 3),
        )
        self.assertEqual(cell.value, date(2026, 8, 3))

        cell.set_value("保留")
        with self.assertRaises(ValueError):
            cell.as_date()
        self.assertEqual(cell.value, "保留")
        cell.set_value(20260801)
        with self.assertRaises(TypeError):
            cell.as_date()
        self.assertEqual(cell.value, 20260801)

    def test_cell_write_back_type_converters(self):
        """功能：验证六种 ``Cell.as_*`` 转换均写回并直接返回目标类型。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言链式转换结果、单元格值和严格转换规则。
        """
        cell = self.worksheet["A1"]
        self.assertEqual(cell.set_value(123).as_string(), "123")
        self.assertEqual(cell.value, "123")
        self.assertEqual(cell.set_value("-42").as_int(), -42)
        self.assertIs(type(cell.value), int)
        self.assertEqual(cell.set_value("1.25e2").as_float(), 125.0)
        self.assertIs(type(cell.value), float)
        self.assertIs(cell.set_value("是").as_bool(), True)
        self.assertIs(cell.value, True)
        self.assertEqual(
            cell.set_value("2026-8-1 12:33:04").as_datetime(),
            datetime(2026, 8, 1, 12, 33, 4),
        )
        self.assertEqual(cell.set_value(None).as_string(), "")
        self.assertFalse(hasattr(cell, "as_str"))

        invalid_cases = [
            ("12.5", "as_int", ValueError),
            (True, "as_float", TypeError),
            (2, "as_bool", ValueError),
            ("not-a-date", "as_datetime", ValueError),
        ]
        for original, method_name, error_type in invalid_cases:
            with self.subTest(method=method_name):
                cell.set_value(original)
                with self.assertRaises(error_type):
                    getattr(cell, method_name)()
                self.assertEqual(cell.value, original)

    def test_cell_read_converters_do_not_write_back(self):
        """功能：验证 ``Cell.read().as_*`` 只转换快照且不影响工作表保存值。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言所有只读转换返回目标类型并保持原单元格值不变。
        """
        cell = self.worksheet["A1"]
        cases = [
            (123, "as_string", "123"),
            ("42", "as_int", 42),
            ("12.5", "as_float", 12.5),
            ("yes", "as_bool", True),
            ("2026/8/1", "as_date", date(2026, 8, 1)),
            ("2026-8-1 12:33", "as_datetime", datetime(2026, 8, 1, 12, 33)),
        ]
        for original, method_name, expected in cases:
            with self.subTest(method=method_name):
                cell.value = original
                result = getattr(cell.read(), method_name)()
                self.assertEqual(result, expected)
                self.assertEqual(cell.value, original)

        cell.value = "7"
        snapshot = cell.read()
        cell.value = "8"
        self.assertEqual(snapshot.value, "7")
        self.assertEqual(snapshot.as_int(), 7)
        self.assertEqual(cell.value, "8")
        with self.assertRaises(AttributeError):
            snapshot.value = "9"

    def test_formula_cached_value_is_read_only_and_invalidated_on_change(self):
        """功能：验证公式缓存结果可只读转换且修改公式会使旧缓存失效。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言缓存读取、转换和公式修改后的失效规则。
        """
        cell = self.worksheet["C3"]
        cell.formula = "=SUM(B3:B3)"
        self.worksheet._set_cached_value(2, 2, 95)
        self.assertEqual(cell.cached_value, 95)
        self.assertEqual(cell.read().as_string(), "95")
        self.assertIsNone(cell.value)
        with self.assertRaises(ValueError):
            cell.as_string()
        self.assertEqual(cell.formula, "=SUM(B3:B3)")

        cell.formula = "=SUM(B3:B3)*2"
        self.assertIsNone(cell.cached_value)
        self.assertEqual(cell.formula, "=SUM(B3:B3)*2")

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

    def test_worksheet_color_normalizes_and_clears(self):
        """功能：验证工作表标签颜色使用统一 RGB/ARGB 规则并支持清除。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.assertIsNone(self.worksheet.color)
        self.worksheet.color = "4472c4"
        self.assertEqual(self.worksheet.color, "FF4472C4")
        self.worksheet.color = "804472C4"
        self.assertEqual(self.worksheet.color, "804472C4")
        with self.assertRaises(ValueError):
            self.worksheet.color = "blue"
        self.assertEqual(self.worksheet.color, "804472C4")
        self.worksheet.color = None
        self.assertIsNone(self.worksheet.color)
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
