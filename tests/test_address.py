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
    index_to_cell,
    cell_to_index,
    column_to_index,
    index_to_column,
    index_to_range,
    range_to_index,
)
from excelkit.errors import InvalidAddressError


class TestAddress(unittest.TestCase):
    def test_cell_to_index(self):
        self.assertEqual(cell_to_index("A1"), (0, 0))
        self.assertEqual(cell_to_index("C8"), (7, 2))
        self.assertEqual(cell_to_index("AA1"), (0, 26))

    def test_index_to_cell(self):
        self.assertEqual(index_to_cell(0, 0), "A1")
        self.assertEqual(index_to_cell(7, 2), "C8")

    def test_range_to_index(self):
        self.assertEqual(range_to_index("B3:D8"), (2, 1, 7, 3))
        self.assertEqual(range_to_index("A1"), (0, 0, 0, 0))
        self.assertEqual(range_to_index("$B$3:$D$8"), (2, 1, 7, 3))

    def test_index_to_range(self):
        self.assertEqual(index_to_range(2, 1, 7, 3), "B3:D8")
        self.assertEqual(index_to_range(0, 0, 0, 0), "A1:A1")

    def test_dollar_and_case(self):
        self.assertEqual(cell_to_index("$A$1"), (0, 0))
        self.assertEqual(cell_to_index("$c$8"), (7, 2))
        self.assertEqual(column_to_index("aa"), 26)
        self.assertEqual(index_to_column(26), "AA")

    def test_top_level_exports(self):
        self.assertEqual(XLSX_MAX_ROWS, 1048576)
        self.assertEqual(XLSX_MAX_COLUMNS, 16384)
        self.assertEqual(XLS_MAX_ROWS, 65536)
        self.assertEqual(XLS_MAX_COLUMNS, 256)

    def test_invalid(self):
        with self.assertRaises(InvalidAddressError):
            cell_to_index("")
        with self.assertRaises(InvalidAddressError):
            cell_to_index("1A")
        with self.assertRaises(InvalidAddressError):
            cell_to_index("A0")
        with self.assertRaises(InvalidAddressError):
            column_to_index("A1")


if __name__ == "__main__":
    unittest.main()
