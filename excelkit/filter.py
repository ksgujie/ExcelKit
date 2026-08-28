"""工作表自动筛选设置。"""

from __future__ import annotations


class AutoFilter:
    """绑定工作表的自动筛选区域和列条件。"""

    __slots__ = ("_worksheet",)

    def __init__(self, worksheet) -> None:
        """功能：创建绑定工作表的筛选代理；参数为 Worksheet；返回无。"""
        self._worksheet = worksheet

    @property
    def range(self):
        """功能：读取或设置筛选区域；参数为 A1 区域或 None；返回规范化地址。"""
        return self._worksheet.auto_filter_range

    @range.setter
    def range(self, value):
        """功能：设置或清除自动筛选区域。

        使用方法：``worksheet.auto_filter.range = 'A1:D100'``，赋值 ``None`` 清除
        筛选区域但不会删除已保存的列条件。
        参数：``value`` 为合法 A1 矩形区域字符串或 ``None``。
        返回：``None``。
        异常：地址类型、格式或边界无效时透传 ``Worksheet.auto_filter_range`` 的异常。
        """
        self._worksheet.auto_filter_range = value

    @property
    def filters(self):
        """功能：取得列筛选条件字典；键为区域内0-based列偏移，值为字符串元组。"""
        return self._worksheet._filter_conditions

    def add(self, column: int, values) -> "AutoFilter":
        """功能：为区域内指定0-based列偏移设置值筛选。

        使用方法：``ws.auto_filter.add(1, ['已完成'])``。返回代理本身。
        """
        if isinstance(column, bool) or not isinstance(column, int) or column < 0:
            raise TypeError("column 必须是非负0-based整数")
        self._worksheet._filter_conditions[column] = tuple(str(value) for value in values)
        return self

    def clear(self) -> "AutoFilter":
        """功能：清除区域和全部筛选条件；返回代理本身。"""
        self._worksheet.auto_filter_range = None
        self._worksheet._filter_conditions.clear()
        return self


__all__ = ["AutoFilter"]
