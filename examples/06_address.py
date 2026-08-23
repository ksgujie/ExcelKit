"""稳定地址工具与 0-based 行列索引示例。"""

from excelkit.address import (
    MAX_COLUMN,
    MAX_ROW,
    column_to_index,
    index_to_column,
    cell_address,
    cell_index,
    range_index,
    range_address,
)


def main() -> None:
    """功能：调用全部稳定地址工具并展示先行后列的转换结果。

    使用方法：在项目根目录执行 ``python -m examples.06_address``。
    参数：无。
    返回：``None``；转换结果输出到终端，不生成文件。
    """
    print("Excel 可用行数：", MAX_ROW)
    print("Excel 可用列数：", MAX_COLUMN)
    print("AA 的 0-based 列索引：", column_to_index("AA"))
    print("列索引 26 的列字母：", index_to_column(26))
    print("C8 的行列索引：", cell_index("C8"))
    print("B3:D8 的边界：", range_index("B3:D8"))
    print("边界转地址：", range_address(2, 1, 7, 3))
    print("行索引 7、列索引 2 的地址：", cell_address(7, 2))


if __name__ == "__main__":
    main()
