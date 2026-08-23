"""连续矩形区域读取与批量写入示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：演示 0-based 区域边界、二维数据写入和普通值读取。

    使用方法：在项目根目录执行 ``python -m examples.03_range``。
    参数：无。
    返回：``None``；在当前目录生成 ``03_range.xlsx``。
    """
    workbook = Workbook()
    worksheet = workbook.active
    area = worksheet.range("B2:D4")

    assert (area.min_row, area.min_column) == (1, 1)
    assert (area.max_row, area.max_column) == (3, 3)

    area.set_values([
        ["姓名", "语文", "数学"],
        ["张三", 90, 95],
        ["李四", 88, 92],
    ])
    print(area.values)

    output = Path("03_range.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
