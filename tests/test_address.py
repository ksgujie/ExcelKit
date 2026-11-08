"""地址解析工具测试。"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import (
    XLS_MAX_COLUMNS,
    XLS_MAX_ROWS,
    XLSX_MAX_COLUMNS,
    XLSX_MAX_ROWS,
    cell_address,
    cell_index,
    column_to_index,
    index_to_column,
    range_address,
    range_index,
)
from excelkit.errors import InvalidAddressError


class TestAddress(unittest.TestCase):
    def test_cell_index(self):
        self.assertEqual(cell_index("A1"), (0, 0))
        self.assertEqual(cell_index("C8"), (7, 2))
        self.assertEqual(cell_index("AA1"), (0, 26))

    def test_cell_address(self):
        self.assertEqual(cell_address(0, 0), "A1")
        self.assertEqual(cell_address(7, 2), "C8")

    def test_range_index(self):
        self.assertEqual(range_index("B3:D8"), (2, 1, 7, 3))
        self.assertEqual(range_index("A1"), (0, 0, 0, 0))
        self.assertEqual(range_index("$B$3:$D$8"), (2, 1, 7, 3))

    def test_range_address(self):
        self.assertEqual(range_address(2, 1, 7, 3), "B3:D8")
        self.assertEqual(range_address(0, 0, 0, 0), "A1:A1")

    def test_dollar_and_case(self):
        self.assertEqual(cell_index("$A$1"), (0, 0))
        self.assertEqual(cell_index("$c$8"), (7, 2))
        self.assertEqual(column_to_index("aa"), 26)
        self.assertEqual(index_to_column(26), "AA")

    def test_top_level_exports(self):
        self.assertEqual(XLSX_MAX_ROWS, 1048576)
        self.assertEqual(XLSX_MAX_COLUMNS, 16384)
        self.assertEqual(XLS_MAX_ROWS, 65536)
        self.assertEqual(XLS_MAX_COLUMNS, 256)

    def test_invalid(self):
        with self.assertRaises(InvalidAddressError):
            cell_index("")
        with self.assertRaises(InvalidAddressError):
            cell_index("1A")
        with self.assertRaises(InvalidAddressError):
            cell_index("A0")
        with self.assertRaises(InvalidAddressError):
            column_to_index("A1")


if __name__ == "__main__":
    unittest.main()
