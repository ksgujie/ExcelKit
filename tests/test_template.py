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
        sheet["A3"] = "{items.@index + 1}"
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
        self.assertEqual(sheet["A2"].value, 1)
        self.assertEqual(sheet["A3"].value, 2)
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

    def test_missing_data_is_atomic_and_default_mode_clears_tag(self):
        """功能：验证严格模式失败不修改工作簿，默认模式把缺失标签清为空值。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言整格、混合文本和公式中的缺失标签均不会残留。
        """
        workbook = self._template()
        original = workbook.active.values
        with self.assertRaises(TemplateError):
            workbook.render({"items": []}, strict=True)
        self.assertEqual(workbook.active.values, original)

        plain = Workbook()
        plain.active["A1"] = "值：{missing}"
        plain.active["A2"] = "{missing}"
        plain.active["A3"].formula = "=1+{missing}"
        plain.render()
        self.assertEqual(plain.active["A1"].value, "值：")
        self.assertIsNone(plain.active["A2"].value)
        self.assertIsNone(plain.active["A3"].formula)

    def test_missing_loop_collection_is_empty_by_default(self):
        """功能：验证默认模式把缺失循环集合视为空数组并删除循环块。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言循环标记和模板行消失，循环后的内容正确上移。
        """
        workbook = self._template()
        workbook.render({"title": "无明细", "name": "管理员"})
        self.assertEqual(workbook.active["A1"].value, "报表：无明细")
        self.assertEqual(workbook.active["A2"].value, "尾部")
        self.assertEqual(workbook.active["D2"].formula, "=C2")

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

    def test_numeric_expressions_and_format_filter(self):
        """功能：验证索引偏移、四则运算、括号及显示格式过滤器。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言计算结果保留数值类型、格式过滤结果为字符串。
        """
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "{loop items}"
        sheet["A2"] = "{items.@index + 1}"
        sheet["B2"] = "{items.quantity * items.price}"
        sheet["C2"] = "{(items.price - items.discount) * items.quantity}"
        sheet["D2"] = '{items.price | format:",.2f"}'
        sheet["E2"] = '金额：{items.quantity * items.price | format:",.2f"}'
        sheet["F2"] = "{items.quantity // divisor}"
        sheet["G2"] = "{items.quantity % divisor}"
        sheet["A3"] = "{/loop}"
        workbook.render({
            "divisor": 2,
            "items": [
                {"quantity": 3, "price": 1234.5, "discount": 34.5},
                {"quantity": 4, "price": 8, "discount": 1},
            ],
        })

        self.assertEqual(sheet["A1"].value, 1)
        self.assertEqual(sheet["A2"].value, 2)
        self.assertEqual(sheet["B1"].value, 3703.5)
        self.assertEqual(sheet["B2"].value, 32)
        self.assertEqual(sheet["C1"].value, 3600.0)
        self.assertEqual(sheet["C2"].value, 28)
        self.assertEqual(sheet["D1"].value, "1,234.50")
        self.assertEqual(sheet["D2"].value, "8.00")
        self.assertEqual(sheet["E1"].value, "金额：3,703.50")
        self.assertEqual(sheet["F1"].value, 1)
        self.assertEqual(sheet["G1"].value, 1)

    def test_expression_errors_are_safe_and_atomic(self):
        """功能：验证非法运算、除零和函数调用均报错且不会执行或修改工作簿。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言失败时由测试框架报告。
        """
        executed = []
        cases = [
            ("{value ** 2}", {"value": 3}),
            ("{value / zero}", {"value": 3, "zero": 0}),
            ('{value | format:"invalid"}', {"value": 3}),
            ("{call(value)}", {"call": lambda _value: executed.append(True), "value": 3}),
        ]
        for template, data in cases:
            with self.subTest(template=template):
                workbook = Workbook()
                workbook.active["A1"] = template
                original = workbook.active.values
                with self.assertRaises(TemplateError):
                    workbook.render(data)
                self.assertEqual(workbook.active.values, original)
        self.assertEqual(executed, [])

    def test_missing_expression_path_respects_strict_mode(self):
        """功能：验证计算表达式缺失路径在严格模式报错、默认模式清空。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；严格模式报错，默认模式删除完整原标签。
        """
        strict_workbook = Workbook()
        strict_workbook.active["A1"] = "{price * quantity}"
        with self.assertRaises(TemplateError):
            strict_workbook.render({"price": 8}, strict=True)

        relaxed_workbook = Workbook()
        relaxed_workbook.active["A1"] = "{price * quantity}"
        relaxed_workbook.render({"price": 8})
        self.assertIsNone(relaxed_workbook.active["A1"].value)

    def test_by_sheet_combines_shared_and_independent_data(self):
        """功能：验证一次调用可按名称或0-based索引为多张表提供独立根数据。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言公共字段共享、独立字段覆盖且未指定工作表不改变。
        """
        workbook = Workbook()
        cover = workbook.add_sheet("封面")
        detail = workbook.add_sheet("明细")
        untouched = workbook.add_sheet("保留")
        cover["A1"] = "{company}"
        cover["B1"] = "{title}"
        detail["A1"] = "{company}"
        detail["B1"] = "{title}"
        detail["A2"] = "{loop items}"
        detail["A3"] = "{items.name}"
        detail["A4"] = "{/loop}"
        untouched["A1"] = "{remain}"

        result = workbook.render(
            {"company": "示例公司", "title": "公共标题"},
            by_sheet={
                "封面": {"title": "封面标题"},
                1: {"title": "明细标题", "items": [{"name": "产品A"}]},
            },
        )

        self.assertIs(result, workbook)
        self.assertEqual(cover.values, [["示例公司", "封面标题"]])
        self.assertEqual(detail.values, [["示例公司", "明细标题"], ["产品A", None]])
        self.assertEqual(untouched["A1"].value, "{remain}")

    def test_by_sheet_validation_and_rendering_are_atomic(self):
        """功能：验证分工作表标识、独立数据及严格渲染失败均不会部分提交。

        使用方法：由 unittest 自动发现执行。
        参数：无。
        返回：无；断言重复工作表和目标表缺失字段得到异常且全部内容保持不变。
        """
        workbook = Workbook()
        first = workbook.add_sheet("一")
        second = workbook.add_sheet("二")
        first["A1"] = "{value}"
        second["A1"] = "{required}"
        original = [sheet.values for sheet in workbook.sheets]

        with self.assertRaises(ValueError):
            workbook.render(by_sheet={"一": {"value": 1}, 0: {"value": 2}})
        self.assertEqual([sheet.values for sheet in workbook.sheets], original)

        with self.assertRaises(TemplateError):
            workbook.render(
                by_sheet={"一": {"value": 1}, "二": {}}, strict=True
            )
        self.assertEqual([sheet.values for sheet in workbook.sheets], original)

        with self.assertRaises(TypeError):
            workbook.render(by_sheet={"一": [1, 2]})


if __name__ == "__main__":
    unittest.main()
