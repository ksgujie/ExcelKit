"""追加单行和多行数据示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：使用 append 和 append_rows 连续生成表格数据。

    使用方法：在项目根目录执行 ``python -m examples.04_append``。
    参数：无。
    返回：``None``；在当前目录生成 ``04_append.xlsx``。
    """
    workbook = Workbook()
    worksheet = workbook.active

    worksheet.append(["姓名", "成绩"])
    worksheet.append_rows([
        ["张三", 95],
        ["李四", 92],
        ["王五", 88],
    ])

    assert worksheet.max_row == 3
    assert worksheet.max_column == 1

    output = Path("04_append.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
