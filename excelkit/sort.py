"""工作表区域排序的公开数据对象。"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SortKey:
    """表示区域排序时的一项 0-based 相对列排序键。"""

    column: int
    descending: bool = False

    def __post_init__(self) -> None:
        """功能：验证排序列索引和方向开关。

        使用方法：``SortKey(1, descending=True)``，表示按区域内第2列降序。
        参数：``column`` 为区域内 0-based 非负列偏移；``descending`` 为是否降序。
        返回：无；对象不可变，便于安全传给 ``Worksheet.sort()``。
        异常：列索引或方向类型无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if isinstance(self.column, bool) or not isinstance(self.column, int):
            raise TypeError("column 必须是 0-based 整数")
        if self.column < 0:
            raise ValueError("column 不能为负数")
        if not isinstance(self.descending, bool):
            raise TypeError("descending 必须是 bool")


__all__ = ["SortKey"]
