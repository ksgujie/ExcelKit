"""A1 地址与 0-based 行列索引的互相转换工具。

所有数字索引统一为 0-based，行列参数统一为先行、后列：
    cell_index("C8") == (7, 2)        # 先行后列
    range_index("B3:D8") == (2, 1, 7, 3)  # (min_row, min_col, max_row, max_col)

A1 字符串地址遵循 Excel 原生表示（1-based 列字母 + 1-based 行号）。
"""

from excelkit.errors import InvalidAddressError

MAX_ROW = 1048576
MAX_COLUMN = 16384


def column_to_index(column: str) -> int:
    """把 Excel 列字母转为 0-based 列索引。

    例如 "A" -> 0, "Z" -> 25, "AA" -> 26, "AB" -> 27。
    """
    if not column or not column.isalpha() or not column.isupper():
        raise InvalidAddressError(f"无效的列字母：{column!r}")
    index = 0
    for ch in column:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def index_to_column(index: int) -> str:
    """把 0-based 列索引转为 Excel 列字母。

    例如 0 -> "A", 25 -> "Z", 26 -> "AA"。
    """
    if index < 0 or index >= MAX_COLUMN:
        raise InvalidAddressError(f"列索引超出范围：{index}")
    letters = []
    n = index + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def cell_index(address: str) -> tuple[int, int]:
    """把 A1 地址解析为 (row, column) 0-based 元组，先行后列。

    例如 "C8" -> (7, 2)。
    """
    import re

    m = re.fullmatch(r"([A-Za-z]+)(\d+)", address.strip())
    if not m:
        raise InvalidAddressError(f"无效的单元格地址：{address!r}")
    col_str, row_str = m.group(1).upper(), m.group(2)
    column = column_to_index(col_str)
    row = int(row_str) - 1
    if row < 0 or row >= MAX_ROW:
        raise InvalidAddressError(f"行号超出范围：{row + 1}")
    return (row, column)


def cell_address(row: int, column: int) -> str:
    """把 (row, column) 0-based 索引转为 A1 地址字符串。

    例如 (7, 2) -> "C8"。
    """
    if row < 0 or row >= MAX_ROW:
        raise InvalidAddressError(f"行索引超出范围：{row}")
    if column < 0 or column >= MAX_COLUMN:
        raise InvalidAddressError(f"列索引超出范围：{column}")
    return f"{index_to_column(column)}{row + 1}"


def range_index(address: str) -> tuple[int, int, int, int]:
    """把 A1 区域地址解析为 (min_row, min_col, max_row, max_col) 0-based 元组。

    例如 "B3:D8" -> (2, 1, 7, 3)。
    支持单单元格地址，例如 "A1" -> (0, 0, 0, 0)。
    """
    import re

    address = address.strip()
    m = re.fullmatch(
        r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)", address
    ) or re.fullmatch(r"([A-Za-z]+)(\d+)", address)
    if not m:
        raise InvalidAddressError(f"无效的区域地址：{address!r}")
    if m.lastindex == 4:
        c1, r1, c2, r2 = m.group(1).upper(), int(m.group(2)), m.group(3).upper(), int(m.group(4))
    else:
        c1, r1, c2, r2 = m.group(1).upper(), int(m.group(2)), m.group(1).upper(), int(m.group(2))
    min_col = column_to_index(c1)
    min_row = r1 - 1
    max_col = column_to_index(c2)
    max_row = r2 - 1
    if min_row > max_row or min_col > max_col:
        raise InvalidAddressError(f"区域起止顺序无效：{address!r}")
    if min_row < 0 or max_row >= MAX_ROW or min_col < 0 or max_col >= MAX_COLUMN:
        raise InvalidAddressError(f"区域超出范围：{address!r}")
    return (min_row, min_col, max_row, max_col)


def range_address(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    """把 (min_row, min_col, max_row, max_col) 0-based 转为 A1 区域地址。

    例如 (2, 1, 7, 3) -> "B3:D8"。
    """
    if min_row > max_row or min_col > max_col:
        raise InvalidAddressError("区域起止索引顺序无效")
    return f"{cell_address(min_row, min_col)}:{cell_address(max_row, max_col)}"


__all__ = [
    "MAX_ROW",
    "MAX_COLUMN",
    "column_to_index",
    "index_to_column",
    "cell_index",
    "cell_address",
    "range_index",
    "range_address",
]
