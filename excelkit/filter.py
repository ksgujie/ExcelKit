"""工作表自动筛选设置。"""

from __future__ import annotations


class AutoFilter:
    """绑定工作表的自动筛选区域和列条件。"""

    __slots__ = ("_worksheet",)

    def __init__(self, worksheet) -> None:
        """功能：创建绑定工作表的筛选代理。

        使用方法：由 ``Worksheet.auto_filter`` 属性内部创建，不直接调用。
        参数：``worksheet`` 为需要管理筛选状态的工作表。
        返回：无。
        """
        self._worksheet = worksheet

    @property
    def range(self):
        """功能：读取自动筛选区域。

        使用方法：``address = worksheet.auto_filter.range``。
        参数：无。
        返回：规范化的 A1 区域字符串或 ``None``。
        """
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
        """功能：取得列筛选条件字典。

        使用方法：``conditions = worksheet.auto_filter.filters``。
        参数：无。
        返回：键为区域内 0-based 列偏移、值为允许文本值元组的内部字典；调用者
        应使用 :meth:`add` 和 :meth:`clear` 修改筛选条件。
        """
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
        """功能：清除区域、筛选条件及本筛选产生的行隐藏状态。

        使用方法：``worksheet.auto_filter.clear()``。
        参数：无。
        返回：当前 ``AutoFilter``，支持链式调用。
        """
        self._worksheet.auto_filter_range = None
        self._worksheet._filter_conditions.clear()
        for row in range(self._worksheet.max_row + 1):
            if row in self._worksheet._rows:
                self._worksheet._rows[row].hidden = False
        return self

    def apply(self) -> "AutoFilter":
        """功能：按已设置的值条件在内存中隐藏不匹配的数据行。

        使用方法：先设置 ``range`` 并通过 :meth:`add` 添加条件，再调用
        ``worksheet.auto_filter.apply()``。
        参数：无；筛选列索引来自已经设置的 ``filters``，并且相对于筛选区域。
        返回：当前 ``AutoFilter``，支持链式调用；没有条件时取消区域内数据行隐藏。
        异常：未设置筛选区域或条件列越出区域时抛出 ``ValueError``。
        """
        if self._worksheet.auto_filter_range is None:
            raise ValueError("应用筛选前必须先设置 auto_filter.range")
        area = self._worksheet.range(self._worksheet.auto_filter_range)
        width = area.max_column - area.min_column + 1
        if any(column >= width for column in self._worksheet._filter_conditions):
            raise ValueError("筛选列索引超出筛选区域")
        for row in range(area.min_row + 1, area.max_row + 1):
            matches = all(
                str(self._worksheet.cell(row, area.min_column + column).value) in candidates
                for column, candidates in self._worksheet._filter_conditions.items()
            )
            self._worksheet.row(row).hidden = not matches if self._worksheet._filter_conditions else False
        return self


__all__ = ["AutoFilter"]
