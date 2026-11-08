"""模板渲染测试。"""

import os
import sys
import unittest

sys.path.insert(0, r'W:\eyes\Temp\opencode')

from excelkit import Workbook


class TestTemplate(unittest.TestCase):
    def test_render_scalar_and_loop(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "{title}"
        ws.append(["产品", "数量", "单价"])
        ws["A3"] = "{loop items}"
        ws["A4"] = "{items.name}"
        ws["B4"] = "{items.quantity}"
        ws["A5"] = "{/loop}"
        wb.render({"title": "报表"}, sheet_data={
            0: {"items": [{"name": "A", "quantity": 2}, {"name": "B", "quantity": 3}]},
        })
        self.assertEqual(ws["A1"].value, "报表")
        self.assertEqual(ws.range("A3:B3").values, [["A", 2]])
        self.assertEqual(ws.range("A4:B4").values, [["B", 3]])


if __name__ == "__main__":
    unittest.main()
