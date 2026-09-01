"""0.9.0 使用范围、校验、转置和新增公式函数回归测试。"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from excelkit import Workbook


class Release090Tests(unittest.TestCase):
    """验证 0.9.0 新增公共 API。"""

    def test_used_range_and_empty_range(self) -> None:
        """功能：验证实际使用范围与空区域判断；参数无；返回无。"""
        workbook = Workbook()
        worksheet = workbook.active
        self.assertIsNone(worksheet.used_range)
        self.assertTrue(worksheet.range("A1:C3").is_empty)
        worksheet["C5"].formula = "=1"
        self.assertEqual(worksheet.used_range.address, "C5:C5")
        worksheet["A1"].value = "数据"
        self.assertEqual(worksheet.used_range.address, "A1:C5")
        self.assertFalse(worksheet.range("A1:C3").is_empty)

    def test_transpose(self) -> None:
        """功能：验证区域转置写入；参数无；返回无。"""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.range("A1:B2").set_values([[1, 2], [3, 4]])
        result = worksheet.range("A1:B2").transpose_to(worksheet.range("D1:E2"))
        self.assertEqual(result.address, "D1:E2")
        self.assertEqual(worksheet.range("D1:E2").values, [[1, 3], [2, 4]])

    def test_validation_and_save_switch(self) -> None:
        """功能：验证结构检查及保存开关；参数无；返回无。"""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet["A1"].formula = "=MissingSheet!A1"
        self.assertTrue(workbook.validate())
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError):
                workbook.save(Path(directory) / "invalid.xlsx", validate=True)

    def test_new_formula_functions(self) -> None:
        """功能：验证 0.9.0 新增公式函数；参数无；返回无。"""
        workbook = Workbook()
        worksheet = workbook.active
        rows = [("张三", "销售", 100), ("李四", "销售", 200), ("王五", "技术", 300)]
        for row, values in enumerate(rows):
            worksheet.cell(row, 0).value = values[0]
            worksheet.cell(row, 1).value = values[1]
            worksheet.cell(row, 2).value = values[2]
        worksheet["E1"].formula = '=SUMIFS(C1:C3,B1:B3,"销售")'
        worksheet["E2"].formula = '=COUNTIFS(B1:B3,"销售")'
        worksheet["E3"].formula = '=INDEX(C1:C3,MATCH("李四",A1:A3,0))'
        worksheet["E4"].formula = '=IFNA(MATCH("不存在",A1:A3,0),"未找到")'
        worksheet["E5"].formula = '=TEXT(1234.5,"#,##0.00")'
        workbook.calculate(strict=True)
        self.assertEqual(worksheet["E1"].value, 300)
        self.assertEqual(worksheet["E2"].value, 2)
        self.assertEqual(worksheet["E3"].value, 200)
        self.assertEqual(worksheet["E4"].value, "未找到")
        self.assertEqual(worksheet["E5"].value, "1,234.50")


if __name__ == "__main__":
    unittest.main()
