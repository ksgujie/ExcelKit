"""Excel 工作表数据表对象。"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import TYPE_CHECKING, Any


class TotalFunction:
    """Excel 数据表汇总行函数常量。

    使用方法：``table.set_total('金额', TotalFunction.SUM)``。
    参数：本类只提供固定字符串常量，不需要实例化。
    返回：无。
    """

    SUM = "sum"
    AVERAGE = "average"
    COUNT = "count"
    COUNT_NUMS = "countNums"
    MIN = "min"
    MAX = "max"


_TOTAL_FUNCTIONS = {
    TotalFunction.SUM, TotalFunction.AVERAGE, TotalFunction.COUNT,
    TotalFunction.COUNT_NUMS, TotalFunction.MIN, TotalFunction.MAX,
}

if TYPE_CHECKING:
    from .range import Range
    from .worksheet import Worksheet


class Table:
    """表示带名称、表头和样式选项的连续 Excel 数据表。"""

    __slots__ = (
        "_worksheet", "_name", "_bounds", "_style", "_has_header",
        "_show_row_stripes", "_show_column_stripes", "_show_totals", "_totals",
        "_data_max_row",
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
        self._data_max_row = area.max_row
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
        """功能：设置是否显示数据表汇总行并维护 Table 的物理区域边界。

        使用方法：``table.show_totals = True``。
        参数：``value`` 必须是布尔值；启用时会占用数据末行下一行作为汇总行。
        返回：``None``。
        异常：汇总行位置已经含有值或公式时抛出 ``ValueError``。
        """
        if not isinstance(value, bool):
            raise TypeError("show_totals 必须是布尔值")
        if value == self._show_totals:
            return
        min_row, min_column, max_row, max_column = self._bounds
        if value:
            total_row = self._data_max_row + 1
            for column in range(min_column, max_column + 1):
                if (
                    self._worksheet._values.get(total_row, column) is not None
                    or (total_row, column) in self._worksheet._formulas
                ):
                    raise ValueError("数据表下一行已有内容，不能作为汇总行")
            self._bounds = (min_row, min_column, total_row, max_column)
            self._worksheet._touch(total_row, max_column)
        else:
            self._bounds = (min_row, min_column, self._data_max_row, max_column)
        self._show_totals = value

    @property
    def totals(self) -> dict[str, str]:
        """功能：取得或修改列汇总函数映射。使用方法：``table.totals['金额'] = 'sum'``。

        参数：字典键为列名，值为 Excel 支持的汇总函数名，如 ``sum``、``average``、
        ``count``、``min``、``max``。返回可直接修改的字典。
        """
        return self._totals

    @property
    def records(self) -> list[dict[str, Any]]:
        """功能：把表头以下的数据读取为按列名组织的字典列表。

        使用方法：``records = table.records``。
        参数：无；当表格没有表头时，自动使用 ``Column1``、``Column2`` 等字段名。
        返回：按当前行顺序排列的 ``list[dict]``，空单元格对应 ``None``。
        """
        names = self.columns
        first = self._bounds[0] + int(self.has_header)
        return [
            {
                name: self._worksheet._values.get(row, column)
                for name, column in zip(names, range(self._bounds[1], self._bounds[3] + 1))
            }
            for row in range(first, self._data_max_row + 1)
        ]

    def set_total(self, column: str, function: str) -> "Table":
        """功能：为指定表头列设置 Excel 汇总行函数并显示汇总行。

        使用方法：``table.set_total('金额', TotalFunction.SUM)``。
        参数：``column`` 为表格中的精确列名；``function`` 为 ``TotalFunction``
        常量之一。
        返回：当前表格，支持链式调用。
        异常：列名或函数不受支持时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(column, str):
            raise TypeError("column 必须是字符串表头")
        if column not in self.columns:
            raise ValueError(f"数据表不存在列：{column!r}")
        if function not in _TOTAL_FUNCTIONS:
            raise ValueError("function 必须是 TotalFunction 的固定值")
        self._totals[column] = function
        self.show_totals = True
        return self

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
        if self._show_totals:
            if area.max_row <= area.min_row:
                raise ValueError("显示汇总行的数据表必须至少包含表头和汇总行")
            self._data_max_row = area.max_row - 1
        else:
            self._data_max_row = area.max_row
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
        row = self._data_max_row + 1
        if self._show_totals:
            next_total_row = row + 1
            for column in range(self._bounds[1], self._bounds[3] + 1):
                if (
                    self._worksheet._values.get(next_total_row, column) is not None
                    or (next_total_row, column) in self._worksheet._formulas
                ):
                    raise ValueError("数据表汇总行下一行已有内容，不能追加记录")
        for offset, value in enumerate(values):
            self._worksheet.cell(row, self._bounds[1] + offset).value = value
        self._data_max_row = row
        final_row = row + int(self._show_totals)
        self._bounds = (self._bounds[0], self._bounds[1], final_row, self._bounds[3])
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

    def append_records(self, records: Iterable[Mapping[str, Any]]) -> "Table":
        """功能：按表头名称批量追加字典记录。

        使用方法：``table.append_records([{'姓名': '张三', '成绩': 95}])``。
        参数：``records`` 为映射对象的可迭代序列；键必须属于数据表表头，缺少的列
        自动写入 ``None``。
        返回：当前表格。
        异常：记录不是映射、包含未知列或不可迭代时抛出 ``TypeError`` 或
        ``ValueError``。
        """
        try:
            prepared = list(records)
        except TypeError as error:
            raise TypeError("records 必须是映射记录的可迭代对象") from error
        names = self.columns
        permitted = set(names)
        for record in prepared:
            if not isinstance(record, Mapping):
                raise TypeError("records 中的每条记录必须是映射对象")
            unknown = set(record) - permitted
            if unknown:
                raise ValueError(f"记录包含数据表中不存在的列：{sorted(unknown)!r}")
        for record in prepared:
            self.append([record.get(name) for name in names])
        return self

    def clear_data(self) -> "Table":
        """功能：清除数据表中除表头外的所有单元格，并保持表格区域。返回当前表格。"""
        first = self._bounds[0] + (1 if self.has_header else 0)
        for row in range(first, self._data_max_row + 1):
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


__all__ = ["Table", "TotalFunction"]
