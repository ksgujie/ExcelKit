"""Excel 工作表数据表对象。"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .range import Range
    from .worksheet import Worksheet


class Table:
    """表示带名称、表头和样式选项的连续 Excel 数据表。"""

    __slots__ = (
        "_worksheet", "_name", "_bounds", "_style", "_has_header",
        "_show_row_stripes", "_show_column_stripes",
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

    def __repr__(self) -> str:
        """功能：生成包含名称、工作表和区域的调试文本。

        使用方法：``repr(table)``。
        参数：无。
        返回：可读字符串。
        """
        return f"<Table {self._name!r} {self._worksheet.name}!{self.range.address}>"


__all__ = ["Table"]
