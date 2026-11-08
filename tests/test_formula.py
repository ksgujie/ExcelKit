"""公式计算器测试。"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import Workbook
from excelkit.errors import FormulaCalculationError


class TestFormulaCalculation(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
        self.ws = self.wb.active

    def test_basic_arithmetic(self):
        self.ws["A1"] = 10
        self.ws["A2"] = 3
        self.ws["A3"].formula = "=A1+A2"
        self.ws["A4"].formula = "=A1-A2"
        self.ws["A5"].formula = "=A1*A2"
        self.ws["A6"].formula = "=A1/A2"
        self.wb.calculate()
        self.assertEqual(self.ws["A3"].value, 13)
        self.assertEqual(self.ws["A4"].value, 7)
        self.assertEqual(self.ws["A5"].value, 30)
        self.assertEqual(self.ws["A6"].value, 10 / 3)

    def test_functions(self):
        self.ws["A1"] = 1
        self.ws["A2"] = 2
        self.ws["A3"] = 3
        self.ws["B1"].formula = "=SUM(A1:A3)"
        self.ws["B2"].formula = "=AVERAGE(A1:A3)"
        self.ws["B3"].formula = "=MAX(A1:A3)"
        self.ws["B4"].formula = "=MIN(A1:A3)"
        self.ws["B5"].formula = "=COUNT(A1:A3)"
        self.ws["B6"].formula = '=IF(A1>A2,"大","小")'
        self.wb.calculate()
        self.assertEqual(self.ws["B1"].value, 6)
        self.assertEqual(self.ws["B2"].value, 2.0)
        self.assertEqual(self.ws["B3"].value, 3)
        self.assertEqual(self.ws["B4"].value, 1)
        self.assertEqual(self.ws["B5"].value, 3)
        self.assertEqual(self.ws["B6"].value, "小")

    def test_cross_sheet(self):
        ws2 = self.wb.add_sheet("S2")
        self.ws["A1"] = 10
        ws2["B1"] = 100
        self.ws["C1"].formula = '=Sheet1!A1 + S2!B1'
        self.wb.calculate(strict=True)
        self.assertEqual(self.ws["C1"].value, 110)

    def test_strict_error(self):
        self.ws["A1"].formula = "=1/0"
        with self.assertRaises(FormulaCalculationError):
            self.wb.calculate(strict=True)

    def test_string_functions(self):
        self.ws["A1"].formula = '=CONCAT("a","b")'
        self.ws["A2"].formula = '=LEN("hello")'
        self.ws["A3"].formula = '=LEFT("hello",2)'
        self.ws["A4"].formula = '=RIGHT("hello",2)'
        self.ws["A5"].formula = '=MID("hello",2,3)'
        self.wb.calculate()
        self.assertEqual(self.ws["A1"].value, "ab")
        self.assertEqual(self.ws["A2"].value, 5)
        self.assertEqual(self.ws["A3"].value, "he")
        self.assertEqual(self.ws["A4"].value, "lo")
        self.assertEqual(self.ws["A5"].value, "ell")


if __name__ == "__main__":
    unittest.main()
