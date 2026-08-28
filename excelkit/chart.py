"""基础 XLSX 图表对象与类型常量。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from .core.worksheet import Worksheet


class ChartType:
    """提供可由 IDE 补全的基础图表类型常量。"""

    COLUMN = "column"
    BAR = "bar"
    LINE = "line"
    PIE = "pie"


class ChartLegend:
    """表示图表图例的位置设置。"""

    BOTTOM = "bottom"
    TOP = "top"
    LEFT = "left"
    RIGHT = "right"
    NONE = "none"

    __slots__ = ("_position",)

    def __init__(self, position: str = BOTTOM) -> None:
        """功能：创建图例位置对象。

        使用方法：由 ``Chart`` 自动创建，通常通过 ``chart.legend.position`` 修改。
        参数：``position`` 为本类位置常量之一，默认底部。
        返回：无。
        异常：位置类型或取值错误时抛出 ``TypeError`` 或 ``ValueError``。
        """
        self.position = position

    @property
    def position(self) -> str:
        """功能：取得图例位置。

        使用方法：``position = chart.legend.position``。
        参数：无。
        返回：``BOTTOM``、``TOP``、``LEFT``、``RIGHT`` 或 ``NONE`` 常量之一。
        """
        return self._position

    @position.setter
    def position(self, value: str) -> None:
        """功能：设置图例位置。

        使用方法：``chart.legend.position = ChartLegend.BOTTOM``。
        参数：``value`` 为 ``BOTTOM``、``TOP``、``LEFT``、``RIGHT`` 或 ``NONE``。
        返回：``None``。
        异常：类型或取值无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(value, str):
            raise TypeError("图例位置必须是字符串常量")
        if value not in {self.BOTTOM, self.TOP, self.LEFT, self.RIGHT, self.NONE}:
            raise ValueError("图例位置无效")
        self._position = value


@dataclass(frozen=True)
class ChartSeries:
    """表示一组同类图表数据引用。"""

    values: str
    categories: Optional[str] = None
    name: Optional[str] = None


class Chart:
    """表示锚定于一个工作表上的基础 XLSX 图表。"""

    __slots__ = ("_worksheet", "type", "anchor", "title", "width", "height", "legend", "_series")

    def __init__(self, worksheet: "Worksheet", chart_type: str, anchor: str) -> None:
        """功能：创建已验证类型和锚点的图表对象。

        使用方法：由 ``Worksheet.add_chart()`` 创建，不直接调用。
        参数：``worksheet`` 为所属表；``chart_type`` 为 ``ChartType`` 常量；``anchor``
        为图表左上角单元格 A1 地址。
        返回：无。
        """
        self._worksheet = worksheet
        self.type = chart_type
        self.anchor = anchor
        self.title = ""
        self.width = 16.0
        self.height = 9.0
        self.legend = ChartLegend()
        self._series: list[ChartSeries] = []

    @property
    def worksheet(self) -> "Worksheet":
        """功能：取得图表所属工作表。

        使用方法：``worksheet = chart.worksheet``。
        参数：无。
        返回：创建该图表的 ``Worksheet``。
        """
        return self._worksheet

    @property
    def series(self) -> tuple[ChartSeries, ...]:
        """功能：取得图表全部数据系列的只读快照。

        使用方法：``series = chart.series``。
        参数：无。
        返回：按添加顺序排列的 ``tuple[ChartSeries, ...]``，不可直接修改。
        """
        return tuple(self._series)

    def add_series(self, *, values: str, categories: Optional[str] = None, name: Optional[str] = None) -> "Chart":
        """功能：添加一组图表数据系列。

        使用方法：``chart.add_series(values='B2:B10', categories='A2:A10', name='成绩')``。
        参数：``values`` 为非空 A1 连续区域；``categories`` 为可选分类区域；``name``
        为可选显示名称。引用区域必须属于图表所在工作表。
        返回：当前 ``Chart``，支持链式调用。
        异常：地址、类型或区域形状无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        value_range = self._worksheet.range(values)
        if categories is not None:
            category_range = self._worksheet.range(categories)
            category_count = (
                (category_range.max_row - category_range.min_row + 1)
                * (category_range.max_column - category_range.min_column + 1)
            )
            value_count = (
                (value_range.max_row - value_range.min_row + 1)
                * (value_range.max_column - value_range.min_column + 1)
            )
            if category_count != value_count:
                raise ValueError("分类区域与数据区域的元素数量必须一致")
        if name is not None and (not isinstance(name, str) or not name):
            raise ValueError("系列名称必须是非空字符串或 None")
        self._series.append(ChartSeries(value_range.address, None if categories is None else category_range.address, name))
        return self

    def remove(self) -> "Worksheet":
        """功能：从所属工作表移除当前图表。使用方法：``chart.remove()``。

        参数：无。
        返回：所属 ``Worksheet``。
        异常：图表已不属于所属表时抛出 ``ValueError``。
        """
        self._worksheet._remove_chart(self)
        return self._worksheet


__all__ = ["Chart", "ChartLegend", "ChartSeries", "ChartType"]
