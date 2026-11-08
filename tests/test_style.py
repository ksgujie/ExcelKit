"""样式系统测试。"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import Workbook
from excelkit.style import Font, Fill, BorderSide, Border, Alignment, Style


class TestStyle(unittest.TestCase):
    def test_style_immutable(self):
        s = Style(font=Font(bold=True, size=12))
        self.assertTrue(s.font.bold)
        self.assertEqual(s.font.size, 12)

    def test_font_fill_border_alignment(self):
        s = Style(
            font=Font(name="微软雅黑", size=14, bold=True, color="FFFFFF"),
            fill=Fill("4472C4"),
            border=Border(bottom=BorderSide(Border.THIN, "000000")),
            alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
            number_format="0.00",
        )
        self.assertEqual(s.font.name, "微软雅黑")
        self.assertEqual(s.fill.fg_color, "4472C4")
        self.assertEqual(s.border.bottom.style, "thin")
        self.assertEqual(s.alignment.horizontal, "center")
        self.assertEqual(s.number_format, "0.00")


if __name__ == "__main__":
    unittest.main()
