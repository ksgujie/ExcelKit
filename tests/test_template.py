"""工作簿模板标量标签、循环块、样式和公式渲染测试。"""

import tempfile
import unittest
from datetime import date
from pathlib import Path

from excelkit import Fill, Font, Style, Workbook
from excelkit.errors import TemplateError


class TemplateTests(unittest.TestCase):
    """验证 Workbook.render 的显式循环作用域和原子行为。"""

    def _template(self) -> Workbook:
        """功能：创建包含一个循环行块的内存模板。

        使用方法：各测试方法调用以获得互不共享的模板工作簿。
        参数：无。
        返回：包含标量、循环、公式、样式和尾行的 Workbook。
        """
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "报表：{title}"
        sheet["A2"] = "{loop items}"
        sheet["A3"] = "{items.@index}"
        sheet["B3"] = "{items.name}"
        sheet["C3"] = "{items.quantity}"
        sheet["D3"].formula = '=LOG10(C3)+C3*{items.price}+"C3"'
        sheet["E3"] = "负责人：{name}"
        sheet["F3"] = "{items.date}"
        sheet["B3"].style = Style(font=Font(bold=True), fill=Fill("FFFF00"))
        sheet["A4"] = "{/loop}"
        sheet["A5"] = "尾部"
        sheet["D5"].formula = "=C5"
        return workbook

    def test_scalar_and_loop_render_preserve_types_styles_and_formulas(self):
        """功能：验证标量替换、显式循环字段、类型、样式和公式行偏移。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = self._template()
        original_style = workbook.active["B3"].style
        result = workbook.render({
            "title": "八月",
            "name": "管理员",
            "items": [
                {"name": "苹果", "quantity": 2, "price": 5, "date": "#2026-8-1"},
                {"name": "梨", "quantity": 3, "price": 4, "date": "#2026/8/2"},
            ],
        })

        self.assertIs(result, workbook)
        sheet = workbook.active
        self.assertEqual(sheet["A1"].value, "报表：八月")
        self.assertEqual(sheet["A2"].value, 0)
        self.assertEqual(sheet["A3"].value, 1)
        self.assertEqual(sheet["B2"].value, "苹果")
        self.assertEqual(sheet["B3"].value, "梨")
        self.assertEqual(sheet["C2"].value, 2)
        self.assertEqual(sheet["C3"].value, 3)
        self.assertEqual(sheet["E2"].value, "负责人：管理员")
        self.assertEqual(sheet["F2"].value, date(2026, 8, 1))
        self.assertEqual(sheet["B2"].style, original_style)
        self.assertEqual(sheet["B3"].style, original_style)
        self.assertEqual(sheet["D2"].formula, '=LOG10(C2)+C2*5+"C3"')
        self.assertEqual(sheet["D3"].formula, '=LOG10(C3)+C3*4+"C3"')
        self.assertEqual(sheet["A4"].value, "尾部")
        self.assertEqual(sheet["D4"].formula, "=C4")
        self.assertEqual((sheet.max_row, sheet.max_column), (3, 5))

    def test_empty_loop_removes_whole_template_block(self):
        """功能：验证空集合会删除开始行、模板行和结束行。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = self._template()
        workbook.render({"title": "空报表", "name": "管理员", "items": []})
        self.assertEqual(workbook.active["A1"].value, "报表：空报表")
        self.assertEqual(workbook.active["A2"].value, "尾部")
        self.assertEqual(workbook.active["D2"].formula, "=C2")
        self.assertEqual(workbook.active.max_row, 1)

    def test_missing_data_is_atomic_and_non_strict_keeps_tag(self):
        """功能：验证严格模式失败不修改工作簿，非严格模式保留普通缺失标签。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = self._template()
        original = workbook.active.values
        with self.assertRaises(TemplateError):
            workbook.render({"items": []})
        self.assertEqual(workbook.active.values, original)

        plain = Workbook()
        plain.active["A1"] = "值：{missing}"
        plain.render({}, strict=False)
        self.assertEqual(plain.active["A1"].value, "值：{missing}")

    def test_invalid_loop_structure_and_data_raise_template_error(self):
        """功能：验证未配对、嵌套和非列表循环数据会得到模板异常。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        missing_end = Workbook()
        missing_end.active["A1"] = "{loop items}"
        missing_end.active["A2"] = "{items.name}"
        with self.assertRaises(TemplateError):
            missing_end.render({"items": []})

        invalid_data = self._template()
        with self.assertRaises(TemplateError):
            invalid_data.render({"title": "错误", "name": "甲", "items": {"name": "乙"}})

    def test_loaded_template_can_render_and_save(self):
        """功能：验证真实 XLSX 模板加载、渲染、保存和重新读取闭环。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        with tempfile.TemporaryDirectory() as directory:
            template_file = Path(directory) / "template.xlsx"
            output_file = Path(directory) / "output.xlsx"
            self._template().save(template_file)
            Workbook.load(template_file).render({
                "title": "闭环",
                "name": "甲",
                "items": [{"name": "项目", "quantity": 1, "price": 8, "date": "#2026-8-1"}],
            }).save(output_file)
            loaded = Workbook.load(output_file)
            self.assertEqual(loaded.active["B2"].value, "项目")
            self.assertEqual(
                loaded.active["D2"].formula, '=LOG10(C2)+C2*8+"C3"'
            )

    def test_multiple_separate_loops_render_from_explicit_prefixes(self):
        """功能：验证同一工作表的多个非嵌套循环可独立展开。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "{loop first}"
        sheet["A2"] = "{first.name}"
        sheet["A3"] = "{/loop}"
        sheet["A4"] = "中间"
        sheet["A5"] = "{loop second}"
        sheet["A6"] = "{second.name}"
        sheet["A7"] = "{/loop}"
        sheet["A8"] = "结束"
        workbook.render({
            "first": [{"name": "甲"}, {"name": "乙"}],
            "second": [{"name": "丙"}],
        })
        self.assertEqual(
            workbook.active.values,
            [["甲"], ["乙"], ["中间"], ["丙"], ["结束"]],
        )


if __name__ == "__main__":
    unittest.main()
