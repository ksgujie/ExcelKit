"""0-based 地址转换和统一异常行为的回归测试。"""

import unittest

from excelkit.address import (
    MAX_COLUMN,
    MAX_ROW,
    column_to_index,
    index_to_column,
    cell_address,
    cell_index,
    range_index,
    range_address,
)
from excelkit.errors import ExcelKitError, InvalidAddressError


class AddressTests(unittest.TestCase):
    """验证 A1 地址与 0-based 行列索引之间的唯一转换 API。"""

    def test_column_conversions_are_zero_based(self):
        """功能：验证列转换从 0 开始且能够双向还原。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        cases = {"A": 0, "Z": 25, "AA": 26, "XFD": 16383}
        for letters, index in cases.items():
            with self.subTest(letters=letters):
                self.assertEqual(column_to_index(letters.lower()), index)
                self.assertEqual(index_to_column(index), letters)

    def test_cell_and_range_parsing_are_zero_based_and_row_first(self):
        """功能：验证解析结果全部 0-based 且元组顺序固定为先行后列。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.assertEqual(cell_index("A1"), (0, 0))
        self.assertEqual(cell_index("C8"), (7, 2))
        self.assertEqual(cell_index("XFD1048576"), (MAX_ROW - 1, MAX_COLUMN - 1))
        self.assertEqual(range_index("B3:D8"), (2, 1, 7, 3))
        self.assertEqual(range_address(2, 1, 7, 3), "B3:D8")

    def test_cell_address_uses_row_then_column(self):
        """功能：验证地址生成参数使用先行后列的 0-based 顺序。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        self.assertEqual(cell_address(0, 0), "A1")
        self.assertEqual(cell_address(7, 2), "C8")
        self.assertEqual(cell_address(MAX_ROW - 1, MAX_COLUMN - 1), "XFD1048576")

    def test_invalid_inputs_raise_one_address_exception(self):
        """功能：验证所有地址错误统一抛出 InvalidAddressError。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        for address in ("", "A0", "A1048577", "XFE1", "A", "1A", "A1:B2"):
            with self.subTest(address=address), self.assertRaises(InvalidAddressError):
                cell_index(address)
        for address in ("A1", "B2:A1", "A2:B1", "A1:"):
            with self.subTest(address=address), self.assertRaises(InvalidAddressError):
                range_index(address)
        for index in (-1, MAX_COLUMN, True, 1.0):
            with self.subTest(index=index), self.assertRaises(InvalidAddressError):
                index_to_column(index)
        self.assertTrue(issubclass(InvalidAddressError, ExcelKitError))
        self.assertTrue(issubclass(InvalidAddressError, ValueError))

    def test_duplicate_address_names_are_not_exported(self):
        """功能：验证地址模块不再暴露同功能的旧别名。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        import excelkit.address as address

        for name in (
            "column_letter_to_index",
            "column_index_to_letter",
            "parse_cell_address",
            "format_cell_address",
            "parse_cell",
            "make_cell_address",
            "parse_range",
        ):
            self.assertFalse(hasattr(address, name))


if __name__ == "__main__":
    unittest.main()
