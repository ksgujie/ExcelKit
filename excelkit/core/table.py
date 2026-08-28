"""Excel 工作表数据表对象。"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .range import Range
    from .worksheet import Worksheet


class Table:
    """表示带名称、表头和样式选项的连续 Excel 数据表。"""

    __slots__ = (
        "_worksheet", "_name", "_bounds", "_style", "_has_header",
        "_show_row_stripes", "_show_column_stripes", "_show_totals", "_totals",
    )

    def __init__(
        self,
        worksheet: "Worksheet",
        name: str,
        area: "Range",
        *,
        style: str = "TableStyleMedium2",
        has_header: bool = True,
        show_row_stripes: bool = True,
        show_column_stripes: bool = False,
    ) -> None:
        """功能：创建已经验证名称和区域的工作表数据表。

        使用方法：由 ``Worksheet.add_table()`` 创建，不直接调用。
        参数：``worksheet`` 为所属表；``name`` 为工作簿内唯一名称；``area`` 为
        表格范围；其余参数控制样式、表头和条纹显示。
        返回：无。
        """
        self._worksheet = worksheet
        self._name = name
        self._bounds = (area.min_row, area.min_column, area.max_row, area.max_column)
        self.style = style
        self.has_header = has_header
        self.show_row_stripes = show_row_stripes
        self.show_column_stripes = show_column_stripes
        self._show_totals = False
        self._totals: dict[str, str] = {}

    @property
    def name(self) -> str:
        """功能：取得数据表的工作簿级唯一名称。

        使用方法：``name = table.name``。
        参数：无，只读属性。
        返回：名称字符串。
        """
        return self._name

    @property
    def worksheet(self) -> "Worksheet":
        """功能：取得数据表所属工作表。

        使用方法：``worksheet = table.worksheet``。
        参数：无，只读属性。
        返回：对应 :class:`Worksheet`。
        """
        return self._worksheet

    @property
    def range(self) -> "Range":
        """功能：取得数据表当前连续区域。

        使用方法：``area = table.range``。
        参数：无，只读属性。
        返回：对应 :class:`Range`。
        """
        from .range import Range

        return Range(self._worksheet, *self._bounds)

    @property
    def style(self) -> str:
        """功能：取得 Excel 数据表样式名称。

        使用方法：``style = table.style``。
        参数：无。
        返回：非空样式名称字符串。
        """
        return self._style

    @style.setter
    def style(self, value: str) -> None:
        """功能：设置 Excel 数据表样式名称。

        使用方法：``table.style = "TableStyleMedium9"``。
        参数：``value`` 为非空字符串。
        返回：``None``；类型或空值无效时抛出 ``ValueError``。
        """
        if not isinstance(value, str) or not value:
            raise ValueError("数据表样式名称必须是非空字符串")
        self._style = value

    @property
    def has_header(self) -> bool:
        """功能：读取数据表首行是否作为表头。

        使用方法：``enabled = table.has_header``。
        参数：无。
        返回：布尔值。
        """
        return self._has_header

    @has_header.setter
    def has_header(self, value: bool) -> None:
        """功能：设置数据表是否包含表头行。

        使用方法：``table.has_header = True``。
        参数：``value`` 必须是布尔值。
        返回：``None``；类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("has_header 必须是布尔值")
        self._has_header = value

    @property
    def show_row_stripes(self) -> bool:
        """功能：读取数据表是否显示隔行条纹。

        使用方法：``enabled = table.show_row_stripes``。
        参数：无。
        返回：布尔值。
        """
        return self._show_row_stripes

    @show_row_stripes.setter
    def show_row_stripes(self, value: bool) -> None:
        """功能：设置数据表隔行条纹显示。

        使用方法：``table.show_row_stripes = False``。
        参数：``value`` 必须是布尔值。
        返回：``None``；类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("show_row_stripes 必须是布尔值")
        self._show_row_stripes = value

    @property
    def show_column_stripes(self) -> bool:
        """功能：读取数据表是否显示隔列条纹。

        使用方法：``enabled = table.show_column_stripes``。
        参数：无。
        返回：布尔值。
        """
        return self._show_column_stripes

    @show_column_stripes.setter
    def show_column_stripes(self, value: bool) -> None:
        """功能：设置数据表隔列条纹显示。

        使用方法：``table.show_column_stripes = True``。
        参数：``value`` 必须是布尔值。
        返回：``None``；类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("show_column_stripes 必须是布尔值")
        self._show_column_stripes = value

    @property
    def columns(self) -> tuple[str, ...]:
        """功能：取得数据表各列的表头名称。

        使用方法：``names = table.columns``。
        参数：无。
        返回：首行单元格文本组成的元组；空表头或重复表头会自动使用
        ``Column1``、``Column2`` 等唯一名称。
        """
        names: list[str] = []
        used: set[str] = set()
        row = self._bounds[0]
        for column in range(self._bounds[1], self._bounds[3] + 1):
            value = self._worksheet.cell(row, column).value if self.has_header else None
            name = str(value) if value not in (None, "") else f"Column{len(names) + 1}"
            base, number = name, 1
            while name.casefold() in used:
                number += 1
                name = f"{base}_{number}"
            used.add(name.casefold())
            names.append(name)
        return tuple(names)

    @property
    def show_totals(self) -> bool:
        """功能：读取是否显示数据表汇总行。使用方法：``table.show_totals``。"""
        return self._show_totals

    @show_totals.setter
    def show_totals(self, value: bool) -> None:
        """功能：设置是否显示数据表汇总行。参数 ``value`` 必须为布尔值。"""
        if not isinstance(value, bool):
            raise TypeError("show_totals 必须是布尔值")
        self._show_totals = value

    @property
    def totals(self) -> dict[str, str]:
        """功能：取得或修改列汇总函数映射。使用方法：``table.totals['金额'] = 'sum'``。

        参数：字典键为列名，值为 Excel 支持的汇总函数名，如 ``sum``、``average``、
        ``count``、``min``、``max``。返回可直接修改的字典。
        """
        return self._totals

    def resize(self, address: str) -> "Table":
        """功能：把数据表调整为同一工作表上的新矩形区域。

        使用方法：``table.resize('A1:F200')``。参数 ``address`` 为标准 A1 区域；
        返回当前表格以支持链式调用。新区域与其他表重叠时抛出 ``ValueError``。
        """
        area = self._worksheet.range(address)
        for other in self._worksheet.tables:
            if other is self:
                continue
            if not (area.max_row < other.range.min_row or area.min_row > other.range.max_row
                    or area.max_column < other.range.min_column or area.min_column > other.range.max_column):
                raise ValueError("数据表区域不能与其他数据表重叠")
        self._bounds = (area.min_row, area.min_column, area.max_row, area.max_column)
        return self

    def append(self, values: list[Any] | tuple[Any, ...]) -> "Table":
        """功能：向数据表末尾追加一行并自动扩展区域。

        使用方法：``table.append(['张三', 95])``。参数 ``values`` 的长度必须等于
        表格列数；返回当前表格。表格仅含表头时追加位置为表头下一行。
        """
        if not isinstance(values, (list, tuple)):
            raise TypeError("values 必须是 list 或 tuple")
        width = self._bounds[3] - self._bounds[1] + 1
        if len(values) != width:
            raise ValueError(f"追加数据需要 {width} 个值")
        row = self._bounds[2] + 1
        for offset, value in enumerate(values):
            self._worksheet.cell(row, self._bounds[1] + offset).value = value
        self._bounds = (self._bounds[0], self._bounds[1], row, self._bounds[3])
        return self

    def append_rows(self, rows: Any) -> "Table":
        """功能：批量追加多行数据。参数 ``rows`` 为可迭代的等长行序列；验证完毕后
        写入并返回当前表格，空序列不改变区域。"""
        prepared = list(rows)
        width = self._bounds[3] - self._bounds[1] + 1
        if any(not isinstance(row, (list, tuple)) or len(row) != width for row in prepared):
            raise ValueError(f"每一行都必须包含 {width} 个值")
        for row in prepared:
            self.append(row)
        return self

    def clear_data(self) -> "Table":
        """功能：清除数据表中除表头外的所有单元格，并保持表格区域。返回当前表格。"""
        first = self._bounds[0] + (1 if self.has_header else 0)
        for row in range(first, self._bounds[2] + 1):
            for column in range(self._bounds[1], self._bounds[3] + 1):
                self._worksheet.cell(row, column).value = None
        return self

    def __repr__(self) -> str:
        """功能：生成包含名称、工作表和区域的调试文本。

        使用方法：``repr(table)``。
        参数：无。
        返回：可读字符串。
        """
        return f"<Table {self._name!r} {self._worksheet.name}!{self.range.address}>"


__all__ = ["Table"]
