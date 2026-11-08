"""地址解析工具测试。"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit.address import cell_index, cell_address, range_index, range_address
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

    def test_range_address(self):
        self.assertEqual(range_address(2, 1, 7, 3), "B3:D8")

    def test_invalid(self):
        with self.assertRaises(InvalidAddressError):
            cell_index("")
        with self.assertRaises(InvalidAddressError):
            cell_index("1A")


if __name__ == "__main__":
    unittest.main()
