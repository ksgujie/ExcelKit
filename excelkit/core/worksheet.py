"""工作表对象及单元格、区域访问接口。"""

from __future__ import annotations

import re
from collections.abc import Iterable
from datetime import date, datetime
from typing import TYPE_CHECKING, Any, Dict, Optional, Tuple, Union

from ..address import (
    cell_address,
    cell_index,
    parse_range,
    validate_column_index,
    validate_row_index,
)
from ..storage import ValueStore
from ..style import DEFAULT_STYLE, Style, _color
from .cell import Cell
from .range import Range

if TYPE_CHECKING:
    from .workbook import Workbook

_DATE_LITERAL_PATTERN = re.compile(
    r"^#(\d{4})([-/])(\d{1,2})\2(\d{1,2})"
    r"(?: (\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$"
)


def _normalize_value(value: Any) -> Any:
    """功能：集中规范化写入工作表的普通值。

    使用方法：由 :meth:`Worksheet._set_value` 内部调用，所有单元格、区域和追加
    写入都经过此函数。
    参数：``value`` 为任意 Python 对象；严格匹配 ``#YYYY-M-D``、
    ``#YYYY/M/D`` 或在其后追加 ``H:M``、``H:M:S`` 的字符串会转换。
    返回：无时间部分时返回 :class:`datetime.date`，有时间部分时返回
    :class:`datetime.datetime`；未匹配字面量时返回原值。
    异常：格式匹配但日期或时间取值无效时抛出 ``ValueError``。
    """
    if not isinstance(value, str):
        return value
    match = _DATE_LITERAL_PATTERN.fullmatch(value)
    if match is None:
        return value
    try:
        year, month, day = (int(match.group(index)) for index in (1, 3, 4))
        if match.group(5) is None:
            return date(year, month, day)
        hour = int(match.group(5))
        minute = int(match.group(6))
        second = int(match.group(7) or 0)
        return datetime(year, month, day, hour, minute, second)
    except ValueError as error:
        raise ValueError(f"无效的日期字面量：{value!r}") from error


class Worksheet:
    """表示隶属于某个 :class:`Workbook` 的命名工作表。"""

    __slots__ = (
        "_workbook",
        "_label",
        "_label_color",
        "_values",
        "_formulas",
        "_styles",
        "_max_row",
        "_max_column",
    )

    def __init__(self, workbook: "Workbook", name: str) -> None:
        """功能：初始化工作表、普通值存储、公式存储和最大索引。

        使用方法：通常由 ``workbook.add_sheet(name)`` 创建，不直接调用。
        参数：``workbook`` 为所属工作簿；``name`` 为已经验证的工作表名称。
        返回：无；空表的最大行、列索引均初始化为 ``-1``。
        """
        self._workbook = workbook
        self._label = name
        self._label_color: Optional[str] = None
        self._values = ValueStore()
        self._formulas: Dict[Tuple[int, int], str] = {}
        self._styles: Dict[Tuple[int, int], Style] = {}
        self._max_row = -1
        self._max_column = -1

    @property
    def label(self) -> str:
        """功能：取得工作表标签名称。

        使用方法：``label = worksheet.label``。
        参数：无。
        返回：工作表底部标签显示的名称字符串。
        """
        return self._label

    @label.setter
    def label(self, label: str) -> None:
        """功能：重命名工作表并同步所属工作簿的名称索引。

        使用方法：``worksheet.label = '新名称'``。
        参数：``label`` 为符合 Excel 规则的新标签名称字符串，长度为 1～31。
        返回：``None``；工作表对象、顺序、数据和样式均保持不变。
        异常：名称无效时抛出 ``InvalidWorksheetNameError``；与其他工作表名称
        大小写不敏感重复时抛出 ``ValueError``。
        """
        self._workbook._rename_sheet(self, label)

    @property
    def label_color(self) -> Optional[str]:
        """功能：取得工作表标签颜色。

        使用方法：``color = worksheet.label_color``。
        参数：无，只读时不需要参数；设置颜色使用同名属性设置器。
        返回：8 位大写 ARGB 字符串；没有设置颜色时返回 ``None``。
        """
        return self._label_color

    @label_color.setter
    def label_color(self, color: Optional[str]) -> None:
        """功能：设置或清除工作表标签颜色。

        使用方法：``worksheet.label_color = '4472C4'``；赋值 ``None`` 清除颜色。
        参数：``color`` 为 6 位 ``RRGGBB``、8 位 ``AARRGGBB`` 字符串或
        ``None``；6 位颜色自动补为完全不透明 ARGB。
        返回：``None``。
        异常：颜色类型、长度或十六进制字符无效时抛出 ``ValueError``。
        """
        self._label_color = _color(color)

    def cell(self, row: Union[str, int], column: Optional[int] = None) -> Cell:
        """功能：按 A1 地址或 0-based 行列索引取得单元格。

        使用方法：固定地址使用 ``worksheet.cell("B3")``；动态坐标使用
        ``worksheet.cell(2, 1)``，两者都指向 ``B3``。
        参数：``row`` 为字符串时表示 A1 地址且必须省略 ``column``；``row`` 为
        整数时表示 0-based 行索引，``column`` 必须是 0-based 列索引。顺序始终
        先行后列，布尔值不作为整数索引。
        返回：指向指定位置的 :class:`Cell`。
        异常：参数组合错误时抛出 ``TypeError``；地址或索引无效时抛出
        ``InvalidAddressError``。
        """
        if isinstance(row, str):
            if column is not None:
                raise TypeError("使用 A1 地址访问单元格时不能再传入 column")
            parsed_row, parsed_column = cell_index(row)
            return Cell(self, parsed_row, parsed_column)
        if isinstance(row, int) and not isinstance(row, bool) and column is not None:
            validate_row_index(row)
            validate_column_index(column)
            return Cell(self, row, column)
        raise TypeError("cell() 需要一个 A1 地址，或 row、column 两个 0-based 整数索引")

    def range(self, address: str) -> Range:
        """功能：按 A1 区域地址取得连续矩形区域。

        使用方法：``area = worksheet.range("A1:C10")``。
        参数：``address`` 为包含冒号的标准矩形 A1 区域字符串。
        返回：对应的 :class:`Range`。
        异常：地址、边界或方向无效时抛出 ``InvalidAddressError``。
        """
        return Range(self, *parse_range(address))

    @property
    def max_row(self) -> int:
        """功能：取得工作表已经触及的最大 0-based 行索引。

        使用方法：A10 被写入后 ``worksheet.max_row`` 返回 ``9``。
        参数：无。
        返回：最大 0-based 行索引；空工作表返回 ``-1``。
        """
        return self._max_row

    @property
    def max_column(self) -> int:
        """功能：取得工作表已经触及的最大 0-based 列索引。

        使用方法：F1 被写入后 ``worksheet.max_column`` 返回 ``5``。
        参数：无。
        返回：最大 0-based 列索引；空工作表返回 ``-1``。
        """
        return self._max_column

    @property
    def values(self) -> list[list[Any]]:
        """功能：读取工作表当前已经触及范围内的全部普通值。

        使用方法：``data = worksheet.values``。
        参数：无，只读属性；需要写入数据时使用单元格、``append``、
        ``append_rows`` 或 ``Range.set_values``。
        返回：从 A1 到 ``max_row``、``max_column`` 的二维列表；空工作表返回
        ``[]``。公式单元格没有普通值，因此对应位置返回 ``None``。
        """
        if self._max_row < 0:
            return []
        end_address = cell_address(self._max_row, self._max_column)
        return self.range(f"A1:{end_address}").values

    def append(self, values: Iterable[Any]) -> "Worksheet":
        """功能：在当前最大行索引之后追加一行普通值。

        使用方法：``worksheet.append(["姓名", "成绩"])``；空表从 A1 开始。
        参数：``values`` 为一维可迭代对象，元素按 0-based 列索引从左到右写入；
        字符串和字节对象不能作为整行数据，空可迭代对象不会触及新行。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：参数不可迭代或是字符串、字节对象时抛出 ``TypeError``；数据超过
        Excel 行列上限时抛出 ``InvalidAddressError``。
        """
        if isinstance(values, (str, bytes)):
            raise TypeError("append() 需要一维行数据，不能直接传入字符串或字节对象")
        try:
            row_values = list(values)
        except TypeError as error:
            raise TypeError("append() 需要一维可迭代对象") from error
        if not row_values:
            return self

        target_row = self._max_row + 1
        validate_row_index(target_row)
        validate_column_index(len(row_values) - 1)
        for column, value in enumerate(row_values):
            self._set_value(target_row, column, value)
        return self

    def append_rows(self, rows: Iterable[Iterable[Any]]) -> "Worksheet":
        """功能：按给定顺序连续追加多行普通值。

        使用方法：``worksheet.append_rows([["张三", 90], ["李四", 88]])``。
        参数：``rows`` 为二维可迭代对象；每个元素会作为一行传给 :meth:`append`。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：外层或任一行不可迭代时抛出 ``TypeError``；超过 Excel 上限时抛出
        ``InvalidAddressError``。已经成功追加的前置行不会回滚。
        """
        if isinstance(rows, (str, bytes)):
            raise TypeError("append_rows() 需要二维数据，不能直接传入字符串或字节对象")
        try:
            iterator = iter(rows)
        except TypeError as error:
            raise TypeError("append_rows() 需要二维可迭代对象") from error
        for values in iterator:
            self.append(values)
        return self

    def _touch(self, row: int, column: int) -> None:
        """功能：更新工作表已经触及的最大 0-based 行列索引。

        使用方法：由普通值和公式写入方法内部调用，业务代码不应直接依赖。
        参数：``row``、``column`` 为有效 0-based 整数索引，顺序为先行后列。
        返回：``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        self._max_row = max(self._max_row, row)
        self._max_column = max(self._max_column, column)

    def _set_value(self, row: int, column: int, value: Any) -> None:
        """功能：设置普通值并清除同一位置的公式。

        使用方法：由 :attr:`Cell.value`、区域写入和追加方法内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列；
        ``value`` 为任意 Python 对象，``None`` 表示删除普通值。
        返回：``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        normalized_value = _normalize_value(value)
        self._touch(row, column)
        self._formulas.pop((row, column), None)
        self._values.set(row, column, normalized_value)

    def _get_formula(self, row: int, column: int) -> Optional[str]:
        """功能：读取指定位置的标准化公式。

        使用方法：由 :attr:`Cell.formula` 读取器内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列。
        返回：带前导 ``=`` 的公式字符串；没有公式时返回 ``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        return self._formulas.get((row, column))

    def _set_formula(self, row: int, column: int, formula: str) -> None:
        """功能：校验并设置公式，同时清除同一位置的普通值。

        使用方法：通过 ``worksheet["A1"].formula = "=SUM(B1:B5)"`` 间接调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列；
        ``formula`` 为非空字符串，可以包含或省略开头的 ``=``。
        返回：``None``；内部统一保存为带前导 ``=`` 的字符串。
        异常：公式不是字符串、为空或只有 ``=`` 时抛出 ``TypeError``；索引无效时
        抛出 ``InvalidAddressError``。
        """
        if not isinstance(formula, str):
            raise TypeError("公式必须是非空字符串")
        expression = formula.strip()
        if expression.startswith("="):
            expression = expression[1:].strip()
        if not expression:
            raise TypeError("公式必须包含表达式")
        self._touch(row, column)
        self._values.set(row, column, None)
        self._formulas[(row, column)] = f"={expression}"

    def _get_style(self, row: int, column: int) -> Style:
        """功能：读取指定位置的单元格样式。

        使用方法：由 :attr:`Cell.style` 读取器内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列。
        返回：已设置的 :class:`Style`；没有自定义样式时返回 ``DEFAULT_STYLE``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        return self._styles.get((row, column), DEFAULT_STYLE)

    def _set_style(self, row: int, column: int, style: Style) -> None:
        """功能：设置或恢复指定位置的单元格样式。

        使用方法：由 ``cell.style = style`` 内部调用。
        参数：``row``、``column`` 为 0-based 整数索引；``style`` 必须是
        :class:`Style`。赋值 ``Style()`` 会移除显式样式记录。
        返回：``None``；即使单元格没有值，设置样式也会更新最大索引。
        异常：样式类型错误时抛出 ``TypeError``；索引无效时抛出
        ``InvalidAddressError``。
        """
        if not isinstance(style, Style):
            raise TypeError("cell.style 必须是 Style 对象")
        self._touch(row, column)
        key = (row, column)
        if style == DEFAULT_STYLE:
            self._styles.pop(key, None)
        else:
            self._styles[key] = style

    def __getitem__(self, address: str) -> Cell:
        """功能：支持 ``worksheet["A1"]`` 形式的单元格读取。

        使用方法：``cell = worksheet["A1"]``。
        参数：``address`` 为合法 A1 单元格地址字符串。
        返回：对应的 :class:`Cell`。
        异常：地址无效时抛出 ``InvalidAddressError``。
        """
        return self.cell(address)

    def __setitem__(self, address: str, value: Any) -> None:
        """功能：支持 ``worksheet["A1"] = value`` 形式的普通值赋值。

        使用方法：``worksheet["A1"] = "标题"``。
        参数：``address`` 为 A1 地址；``value`` 为普通值，``None`` 表示清除。
        返回：``None``；写入普通值会清除同一位置的公式。
        异常：地址无效时抛出 ``InvalidAddressError``。
        """
        self.cell(address).value = value

    def __repr__(self) -> str:
        """功能：生成用于调试的工作表文本表示。

        使用方法：``repr(worksheet)``。
        参数：无。
        返回：包含工作表名称的字符串。
        """
        return f"<Worksheet {self._label!r}>"
