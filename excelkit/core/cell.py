"""单元格的普通值和公式访问对象。"""

from __future__ import annotations

from datetime import date, datetime
from typing import TYPE_CHECKING, Any, Optional

from ..address import cell_address
from ..style import Style
from .conversion import (
    as_bool as _convert_bool,
    as_date as _convert_date,
    as_datetime as _convert_datetime,
    as_float as _convert_float,
    as_int as _convert_int,
    as_string as _convert_string,
)

if TYPE_CHECKING:
    from .worksheet import Worksheet


class CellValue:
    """表示一次 ``Cell.read()`` 获得的不可变普通值快照。"""

    __slots__ = ("_value",)

    def __init__(self, value: Any) -> None:
        """功能：创建不会写回工作表的单元格值快照。

        使用方法：由 ``cell.read()`` 创建，通常不直接实例化。
        参数：``value`` 为读取瞬间的单元格普通值。
        返回：无。
        """
        self._value = value

    def as_string(self) -> str:
        """功能：只读转换快照为字符串，不修改原单元格。

        使用方法：``text = cell.read().as_string()``。
        参数：无；``None`` 转为空字符串，其余值使用 ``str()``。
        返回：转换后的 ``str``。
        异常：对象自身字符串转换失败时透传异常。
        """
        return _convert_string(self._value)

    def as_int(self) -> int:
        """功能：只读转换快照为整数，不修改原单元格。

        使用方法：``number = cell.read().as_int()``。
        参数：无；接受整数、整数字符串和无小数部分的有限浮点数。
        返回：转换后的 ``int``。
        异常：无法无损转换时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return _convert_int(self._value)

    def as_float(self) -> float:
        """功能：只读转换快照为有限浮点数，不修改原单元格。

        使用方法：``number = cell.read().as_float()``。
        参数：无；接受字符串、整数或浮点数，不接受布尔值。
        返回：转换后的有限 ``float``。
        异常：类型、格式无效或结果非有限时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return _convert_float(self._value)

    def as_bool(self) -> bool:
        """功能：只读转换快照为布尔值，不修改原单元格。

        使用方法：``flag = cell.read().as_bool()``。
        参数：无；仅接受统一真假值白名单中的文本、布尔值或 0/1。
        返回：转换后的 ``bool``。
        异常：类型或取值无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return _convert_bool(self._value)

    def as_date(self) -> date:
        """功能：只读转换快照为日期，不修改原单元格。

        使用方法：``day = cell.read().as_date()``。
        参数：无；接受支持格式的字符串、``date`` 或 ``datetime``。
        返回：转换后的 :class:`date`。
        异常：类型、格式或日期取值无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return _convert_date(self._value)

    def as_datetime(self) -> datetime:
        """功能：只读转换快照为日期时间，不修改原单元格。

        使用方法：``moment = cell.read().as_datetime()``。
        参数：无；纯日期的时间部分补为 ``00:00:00``。
        返回：转换后的 :class:`datetime`。
        异常：类型、格式或日期时间无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return _convert_datetime(self._value)

    def __repr__(self) -> str:
        """功能：生成用于调试的只读值快照表示。

        使用方法：``repr(cell.read())``。
        参数：无。
        返回：包含原始快照值的字符串。
        """
        return f"<CellValue value={self._value!r}>"


class Cell:
    """表示工作表中一个由 0-based 行列索引定位的单元格。"""

    __slots__ = ("_worksheet", "_row", "_column")

    def __init__(self, worksheet: "Worksheet", row: int, column: int) -> None:
        """功能：创建指向指定工作表位置的单元格对象。

        使用方法：由 ``worksheet["A1"]`` 或 ``worksheet.cell(...)`` 创建。
        参数：``worksheet`` 为所属工作表；``row``、``column`` 为已验证的
        0-based 整数索引，顺序为先行后列。
        返回：无。
        """
        self._worksheet = worksheet
        self._row = row
        self._column = column

    @property
    def row(self) -> int:
        """功能：取得单元格的 0-based 行索引。

        使用方法：``worksheet["C8"].row`` 返回 ``7``。
        参数：无。
        返回：0～1048575 范围内的整数。
        """
        return self._row

    @property
    def column(self) -> int:
        """功能：取得单元格的 0-based 列索引。

        使用方法：``worksheet["C8"].column`` 返回 ``2``。
        参数：无。
        返回：0～16383 范围内的整数。
        """
        return self._column

    @property
    def index(self) -> tuple[int, int]:
        """功能：取得单元格的 0-based 行列组合索引。

        使用方法：``worksheet["D3"].index`` 返回 ``(2, 3)``。
        参数：无；本属性只读，元组元素顺序固定为先行后列。
        返回：``(row, column)`` 形式的二元整数元组。
        """
        return self._row, self._column

    @property
    def address(self) -> str:
        """功能：取得单元格对应的规范化 A1 地址。

        使用方法：``worksheet.cell(7, 2).address`` 返回 ``"C8"``。
        参数：无。
        返回：列字母大写的 A1 地址字符串。
        """
        return cell_address(self._row, self._column)

    @property
    def value(self) -> Any:
        """功能：读取单元格的普通值。

        使用方法：``value = worksheet["A1"].value``。
        参数：无。
        返回：已保存的普通值；空单元格或公式单元格返回 ``None``。
        """
        return self._worksheet._values.get(self._row, self._column)

    @value.setter
    def value(self, value: Any) -> None:
        """功能：写入或清除普通值，并清除同一位置的公式。

        使用方法：``worksheet["A1"].value = 100``；赋值 ``None`` 表示清除。
        参数：``value`` 为任意 Python 对象或 ``None``。
        返回：``None``。
        """
        self._worksheet._set_value(self._row, self._column, value)

    def set_value(self, value: Any) -> "Cell":
        """功能：以链式调用方式写入普通值，并清除同一位置的公式。

        使用方法：``worksheet.cell("A1").set_value("2026-8-1").as_date()``；
        不需要链式转换时仍可直接使用 ``cell.value = value``。
        参数：``value`` 为任意 Python 对象或 ``None``，处理规则与
        :attr:`value` 属性赋值完全一致。
        返回：当前 :class:`Cell`，可继续调用 ``as_date()`` 等链式方法。
        异常：值规范化失败时透传相应异常，并保持原单元格内容不变。
        """
        self.value = value
        return self

    def read(self) -> CellValue:
        """功能：读取当前普通值的不可变快照，用于不写回的链式类型转换。

        使用方法：``number = cell.read().as_int()``；与直接 ``cell.as_int()``
        不同，转换结果不会修改工作表，也不会影响之后的文件保存内容。
        参数：无。
        返回：包含读取瞬间原始值的 :class:`CellValue`。
        """
        return CellValue(self.value)

    def as_string(self) -> str:
        """功能：转换当前值为字符串、写回单元格并返回转换结果。

        使用方法：``text = cell.set_value(123).as_string()``。
        参数：无；``None`` 转为空字符串，其余值使用 ``str()``。
        返回：写回后的 ``str``。
        异常：转换失败时透传异常并保持原值不变。
        """
        converted = _convert_string(self.value)
        self.value = converted
        return converted

    def as_int(self) -> int:
        """功能：无损转换当前值为整数、写回单元格并返回转换结果。

        使用方法：``number = cell.set_value("123").as_int()``。
        参数：无；接受整数、整数字符串和无小数部分的有限浮点数。
        返回：写回后的 ``int``。
        异常：无法无损转换时抛出 ``TypeError`` 或 ``ValueError``，原值不变。
        """
        converted = _convert_int(self.value)
        self.value = converted
        return converted

    def as_float(self) -> float:
        """功能：转换当前值为有限浮点数、写回单元格并返回转换结果。

        使用方法：``number = cell.set_value("12.5").as_float()``。
        参数：无；接受字符串、整数或浮点数，不接受布尔值。
        返回：写回后的有限 ``float``。
        异常：类型、格式无效或结果非有限时抛出异常，原值不变。
        """
        converted = _convert_float(self.value)
        self.value = converted
        return converted

    def as_bool(self) -> bool:
        """功能：转换当前值为布尔值、写回单元格并返回转换结果。

        使用方法：``flag = cell.set_value("是").as_bool()``。
        参数：无；仅接受统一真假值白名单中的文本、布尔值或 0/1。
        返回：写回后的 ``bool``。
        异常：类型或取值无效时抛出异常，原值不变。
        """
        converted = _convert_bool(self.value)
        self.value = converted
        return converted

    def as_date(self) -> date:
        """功能：转换当前值为日期、写回单元格并返回转换结果。

        使用方法：``day = cell.set_value("2026/8/1").as_date()``。
        参数：无；接受支持格式的字符串、``date`` 或 ``datetime``。
        返回：写回后的 :class:`date`；``datetime`` 会舍弃时间部分。
        异常：类型、格式或日期取值无效时抛出异常，原值不变。
        """
        converted = _convert_date(self.value)
        self.value = converted
        return converted

    def as_datetime(self) -> datetime:
        """功能：转换当前值为日期时间、写回单元格并返回转换结果。

        使用方法：``moment = cell.set_value("2026-8-1 12:33").as_datetime()``。
        参数：无；接受支持格式的字符串、``date`` 或 ``datetime``。
        返回：写回后的 :class:`datetime`；纯日期补为午夜时间。
        异常：类型、格式或日期时间无效时抛出异常，原值不变。
        """
        converted = _convert_datetime(self.value)
        self.value = converted
        return converted

    @property
    def formula(self) -> Optional[str]:
        """功能：读取单元格公式。

        使用方法：``formula = worksheet["D2"].formula``。
        参数：无。
        返回：带前导 ``=`` 的标准化公式字符串；没有公式时返回 ``None``。
        """
        return self._worksheet._get_formula(self._row, self._column)

    @formula.setter
    def formula(self, formula: str) -> None:
        """功能：设置单元格公式并清除同一位置的普通值。

        使用方法：``worksheet["D2"].formula = "=SUM(B2:C2)"``；开头的 ``=``
        可以省略。
        参数：``formula`` 必须是包含表达式的非空字符串。
        返回：``None``。
        异常：公式类型或内容无效时抛出 ``TypeError``。
        """
        self._worksheet._set_formula(self._row, self._column, formula)

    @property
    def style(self) -> Style:
        """功能：读取单元格完整样式。

        使用方法：``style = worksheet["A1"].style``。
        参数：无。
        返回：不可变 :class:`Style`；未设置时返回默认 ``Style()``。
        """
        return self._worksheet._get_style(self._row, self._column)

    @style.setter
    def style(self, style: Style) -> None:
        """功能：设置字体、填充、边框、对齐和数字格式组合样式。

        使用方法：``worksheet["A1"].style = Style(...)``。
        参数：``style`` 必须是 :class:`Style`；不接受含义不明确的字典。
        返回：``None``。
        异常：类型错误时抛出 ``TypeError``。
        """
        self._worksheet._set_style(self._row, self._column, style)

    def __repr__(self) -> str:
        """功能：生成用于调试的单元格文本表示。

        使用方法：``repr(cell)``。
        参数：无。
        返回：包含工作表名、A1 地址、普通值和公式的字符串。
        """
        return (
            f"<Cell {self._worksheet.label}!{self.address} "
            f"value={self.value!r} formula={self.formula!r}>"
        )
