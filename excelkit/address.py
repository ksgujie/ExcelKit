"""Excel A1 地址与 0-based 行列索引之间的转换工具。"""

from __future__ import annotations

import re
from typing import Tuple

from .errors import InvalidAddressError

# MAX_ROW 和 MAX_COLUMN 表示 Excel 可用行、列的总数量，不是最大索引。
MAX_ROW = 1_048_576
MAX_COLUMN = 16_384

_CELL_PATTERN = re.compile(r"^([A-Za-z]+)([1-9][0-9]*)$")


def validate_row_index(row: int) -> int:
    """功能：验证 0-based 行索引。

    使用方法：内部需要检查数字行坐标时调用 ``validate_row_index(row)``。
    参数：``row`` 必须是整数，范围为 0～1048575；布尔值不作为整数索引。
    返回：验证后的原行索引。
    异常：类型或范围无效时抛出 :class:`InvalidAddressError`。
    """
    if isinstance(row, bool) or not isinstance(row, int) or not 0 <= row < MAX_ROW:
        raise InvalidAddressError(
            f"行索引必须是 0 到 {MAX_ROW - 1} 之间的整数，实际为 {row!r}"
        )
    return row


def validate_column_index(column: int) -> int:
    """功能：验证 0-based 列索引。

    使用方法：内部需要检查数字列坐标时调用 ``validate_column_index(column)``。
    参数：``column`` 必须是整数，范围为 0～16383；布尔值不作为整数索引。
    返回：验证后的原列索引。
    异常：类型或范围无效时抛出 :class:`InvalidAddressError`。
    """
    if (
        isinstance(column, bool)
        or not isinstance(column, int)
        or not 0 <= column < MAX_COLUMN
    ):
        raise InvalidAddressError(
            f"列索引必须是 0 到 {MAX_COLUMN - 1} 之间的整数，实际为 {column!r}"
        )
    return column


def column_to_index(column: str) -> int:
    """功能：把 Excel 列字母转换为 0-based 列索引。

    使用方法：``column_to_index("A")`` 返回 ``0``，``column_to_index("AA")``
    返回 ``26``。
    参数：``column`` 必须是仅含 ASCII 英文字母的字符串，不区分大小写，最大为
    ``XFD``。
    返回：0～16383 范围内的整数列索引。
    异常：类型、字符或范围无效时抛出 :class:`InvalidAddressError`。
    """
    if (
        not isinstance(column, str)
        or not column
        or not column.isascii()
        or not column.isalpha()
    ):
        raise InvalidAddressError(f"无效的列名称：{column!r}")

    one_based = 0
    for character in column.upper():
        one_based = one_based * 26 + ord(character) - ord("A") + 1
    index = one_based - 1
    validate_column_index(index)
    return index


def index_to_column(index: int) -> str:
    """功能：把 0-based 列索引转换为 Excel 列字母。

    使用方法：``index_to_column(0)`` 返回 ``"A"``，``index_to_column(26)``
    返回 ``"AA"``。
    参数：``index`` 必须是 0～16383 的整数；布尔值不作为整数索引。
    返回：``A``～``XFD`` 范围内的大写列字母字符串。
    异常：类型或范围无效时抛出 :class:`InvalidAddressError`。
    """
    validate_column_index(index)
    one_based = index + 1
    letters = []
    while one_based:
        one_based, remainder = divmod(one_based - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def cell_index(address: str) -> Tuple[int, int]:
    """功能：把 A1 单元格地址解析为 0-based 行列索引。

    使用方法：``cell_index("C8")`` 返回 ``(7, 2)``。
    参数：``address`` 必须是“列字母+Excel 行号”形式的字符串。
    返回：``(行索引, 列索引)`` 元组，顺序固定为先行后列，两项均从 0 开始。
    异常：格式、类型或 Excel 边界无效时抛出 :class:`InvalidAddressError`。
    """
    if not isinstance(address, str):
        raise InvalidAddressError(f"单元格地址必须是字符串，实际为 {type(address).__name__}")
    match = _CELL_PATTERN.fullmatch(address)
    if match is None:
        raise InvalidAddressError(f"无效的单元格地址：{address!r}")
    column = column_to_index(match.group(1))
    row = int(match.group(2)) - 1
    validate_row_index(row)
    return row, column


def parse_range(address: str) -> Tuple[int, int, int, int]:
    """功能：把 A1 矩形区域解析为 0-based 边界索引。

    使用方法：``parse_range("B3:D8")`` 返回 ``(2, 1, 7, 3)``。
    参数：``address`` 必须是 ``起始单元格:结束单元格`` 字符串，起点不能位于
    终点的下方或右侧。
    返回：``(最小行, 最小列, 最大行, 最大列)`` 元组，顺序始终先行后列。
    异常：格式、边界或方向无效时抛出 :class:`InvalidAddressError`。
    """
    if not isinstance(address, str):
        raise InvalidAddressError(f"区域地址必须是字符串，实际为 {type(address).__name__}")
    parts = address.split(":")
    if len(parts) != 2 or not all(parts):
        raise InvalidAddressError(f"无效的区域地址：{address!r}")
    try:
        min_row, min_column = cell_index(parts[0])
        max_row, max_column = cell_index(parts[1])
    except InvalidAddressError as error:
        raise InvalidAddressError(f"无效的区域地址：{address!r}") from error
    if min_row > max_row or min_column > max_column:
        raise InvalidAddressError(f"区域起点必须位于终点的左上方：{address!r}")
    return min_row, min_column, max_row, max_column


def cell_address(row: int, column: int) -> str:
    """功能：把 0-based 行列索引转换为 A1 单元格地址。

    使用方法：``cell_address(7, 2)`` 返回 ``"C8"``。
    参数：``row``、``column`` 均为 0-based 整数索引，参数顺序固定为先行后列。
    返回：列字母大写的 A1 地址字符串。
    异常：任一索引的类型或范围无效时抛出 :class:`InvalidAddressError`。
    """
    validate_row_index(row)
    return f"{index_to_column(column)}{row + 1}"


__all__ = [
    "MAX_ROW",
    "MAX_COLUMN",
    "column_to_index",
    "index_to_column",
    "cell_index",
    "parse_range",
    "cell_address",
]
