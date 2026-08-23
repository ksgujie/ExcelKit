"""普通单元格值的内部稀疏存储。"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any, Dict, Tuple

from .address import validate_column_index, validate_row_index


class ValueStore:
    """以 ``(0-based 行索引, 0-based 列索引)`` 为键保存非空普通值。"""

    __slots__ = ("_values",)

    def __init__(self) -> None:
        """功能：创建空的普通值存储。

        使用方法：由 :class:`Worksheet` 内部创建，普通用户不需要直接实例化。
        参数：无。
        返回：无。
        """
        self._values: Dict[Tuple[int, int], Any] = {}

    def get(self, row: int, column: int) -> Any:
        """功能：读取指定 0-based 行列索引处的普通值。

        使用方法：内部调用 ``store.get(row, column)``。
        参数：``row``、``column`` 均为 0-based 整数，顺序固定为先行后列。
        返回：已保存的 Python 值；没有普通值时返回 ``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        return self._values.get((row, column))

    def set(self, row: int, column: int, value: Any) -> None:
        """功能：设置或删除指定 0-based 行列索引处的普通值。

        使用方法：内部调用 ``store.set(row, column, value)``。
        参数：``row``、``column`` 为 0-based 整数，顺序为先行后列；``value``
        为任意 Python 对象，``None`` 表示删除。
        返回：``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        key = (row, column)
        if value is None:
            self._values.pop(key, None)
        else:
            self._values[key] = value

    def items(self) -> Iterator[Tuple[Tuple[int, int], Any]]:
        """功能：按行优先顺序迭代全部非空普通值。

        使用方法：``for (row, column), value in store.items(): ...``。
        参数：无。
        返回：依次产生 ``((行索引, 列索引), 值)`` 的迭代器。
        """
        for key in sorted(self._values):
            yield key, self._values[key]


__all__ = ["ValueStore"]
