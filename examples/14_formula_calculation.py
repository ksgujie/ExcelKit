"""公式计算、缓存状态和只读结果转换完整示例。"""

from pathlib import Path

from excelkit import Workbook


def create_formula_workbook() -> Workbook:
    """功能：创建包含跨表引用、依赖公式和常用函数的工作簿。

    使用方法：``workbook = create_formula_workbook()``。
    参数：无。
    返回：已经完成公式计算并带缓存结果的 :class:`Workbook`。
    """
    workbook = Workbook()
    source = workbook.add_sheet("销售明细")
    source.append_rows([
        ["产品", "数量", "单价"],
        ["产品 A", 2, 19.5],
        ["产品 B", 3, 8],
    ])

    summary = workbook.add_sheet("汇总")
    summary["A1"] = "销售额"
    summary["B1"].formula = (
        "='销售明细'!B2*'销售明细'!C2+"
        "'销售明细'!B3*'销售明细'!C3"
    )
    summary["A2"] = "说明"
    summary["B2"].formula = '=CONCAT("合计：",ROUND(B1,2))'

    workbook.calculate()
    assert summary["B1"].value == 63
    assert summary["B1"].formula_status == "calculated"
    assert summary["B1"].read().as_float() == 63.0
    return workbook


def main() -> None:
    """功能：保存带公式表达式和缓存结果的 XLSX 示例文件。

    使用方法：在项目根目录执行 ``python -m examples.14_formula_calculation``。
    参数：无。
    返回：``None``；生成 ``14_formula_calculation.xlsx``。
    """
    output = Path("14_formula_calculation.xlsx")
    create_formula_workbook().save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
