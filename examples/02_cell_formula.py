"""0-based 单元格索引与公式读写示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：演示 cell 的两种调用形式以及普通值和公式相互覆盖。

    使用方法：在项目根目录执行 ``python -m examples.02_cell_formula``。
    参数：无。
    返回：``None``；在当前目录生成 ``02_cell_formula.xlsx``。
    """
    workbook = Workbook()
    worksheet = workbook.active

    worksheet.cell(0, 0).value = 10  # 0-based (0, 0) 对应 A1。
    worksheet.cell(1, 0).value = 20  # 0-based (1, 0) 对应 A2。
    worksheet.cell("A3").formula = "SUM(A1:A2)"

    assert worksheet["A3"].formula == "=SUM(A1:A2)"
    assert worksheet["A3"].value is None
    assert (worksheet["A3"].row, worksheet["A3"].column) == (2, 0)

    output = Path("02_cell_formula.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
