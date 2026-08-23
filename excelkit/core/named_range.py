"""工作簿级命名区域对象。"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .range import Range
    from .workbook import Workbook
    from .worksheet import Worksheet


class NamedRange:
    """表示工作簿中由唯一名称引用的一块工作表区域。"""

    __slots__ = ("_workbook", "_name", "_worksheet", "_bounds")

    def __init__(self, workbook: "Workbook", name: str, area: "Range") -> None:
        """功能：保存命名区域的工作簿、名称、工作表和0-based边界。

        使用方法：由 ``Workbook.add_named_range()`` 创建，不直接调用。
        参数：``workbook`` 为所属工作簿；``name`` 为已验证唯一名称；``area``
        为同一工作簿中的 :class:`Range`。
        返回：无。
        """
        self._workbook = workbook
        self._name = name
        self._worksheet = area.worksheet
        self._bounds = (
            area.min_row,
            area.min_column,
            area.max_row,
            area.max_column,
        )

    @property
    def name(self) -> str:
        """功能：取得命名区域的唯一名称。

        使用方法：``name = named_range.name``。
        参数：无，只读属性。
        返回：创建时登记的名称字符串。
        """
        return self._name

    @property
    def worksheet(self) -> "Worksheet":
        """功能：取得命名区域当前所属工作表。

        使用方法：``worksheet = named_range.worksheet``。
        参数：无，只读属性。
        返回：对应 :class:`Worksheet`；工作表重命名不影响对象关系。
        """
        return self._worksheet

    @property
    def range(self) -> "Range":
        """功能：根据保存的边界创建当前命名区域的 Range。

        使用方法：``area = workbook.named_range("SalesAmount").range``。
        参数：无，只读属性。
        返回：对应 :class:`Range`。
        """
        from .range import Range

        return Range(self._worksheet, *self._bounds)

    def __repr__(self) -> str:
        """功能：生成包含名称、工作表和A1地址的调试表示。

        使用方法：``repr(named_range)``。
        参数：无。
        返回：可读字符串。
        """
        return f"<NamedRange {self._name!r} {self._worksheet.name}!{self.range.address}>"


__all__ = ["NamedRange"]
