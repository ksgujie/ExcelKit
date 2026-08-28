"""单元格的普通值和公式访问对象。"""

from __future__ import annotations

from datetime import date, datetime
from typing import TYPE_CHECKING, Any, Optional

from ..address import cell_address
from ..hyperlink import Hyperlink
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
    """表示一次 ``Cell.read()`` 获得的不可变值快照。

    普通单元格快照保存 ``Cell.value``；公式单元格快照优先保存 Excel/WPS
    最近一次写入文件的缓存计算结果。快照转换始终只影响 Python 返回值，不会写回工作簿。
    """

    __slots__ = ("_value",)

    def __init__(self, value: Any) -> None:
        """功能：创建不会写回工作表的单元格值快照。

        使用方法：由 ``cell.read()`` 创建，通常不直接实例化。
        参数：``value`` 为读取瞬间的普通值或公式缓存结果。
        返回：无。
        """
        self._value = value

    @property
    def value(self) -> Any:
        """功能：取得创建快照时读取到的原始普通值。

        使用方法：``raw_value = cell.read().value``。
        参数：无；本属性只读，不会把任何内容写回原单元格。
        返回：快照保存的原始 Python 值；空单元格或没有缓存结果的公式单元格返回 ``None``。
        """
        return self._value

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

    @property
    def hyperlink(self) -> Optional[Hyperlink]:
        """功能：读取当前单元格的超链接定义。

        使用方法：``link = worksheet["A1"].hyperlink``。
        参数：无，只读时不需要参数；设置链接使用同名属性设置器。
        返回：``Hyperlink``；当前单元格没有超链接时返回 ``None``。
        """
        return self._worksheet._get_hyperlink(self._row, self._column)

    @hyperlink.setter
    def hyperlink(self, value: Optional[Hyperlink | str]) -> None:
        """功能：设置、替换或清除当前单元格的超链接。

        使用方法：``cell.hyperlink = "https://example.com"``；也可赋值
        ``Hyperlink(...)``，赋值 ``None`` 清除链接。
        参数：``value`` 为网址/文件路径字符串、``Hyperlink`` 对象或 ``None``。
        返回：``None``；链接不改变单元格普通值、公式和样式。
        异常：类型或超链接字段无效时抛出 ``TypeError`` 或 ``ValueError``；合并区域
        的非左上角单元格禁止设置。
        """
        self._worksheet._set_hyperlink(self._row, self._column, value)

    @property
    def cached_value(self) -> Any:
        """功能：读取公式最近一次由表格软件或工作簿计算器生成的缓存结果。

        使用方法：``result = worksheet["C3"].cached_value``。
        参数：无；本属性只读，不会触发公式计算。
        返回：公式缓存的 Python 值；普通单元格、未计算公式或已失效缓存返回
        ``None``。缓存值可能因源数据变化而过期，不能替代表格软件重新计算。
        """
        return self._worksheet._get_cached_value(self._row, self._column)

    @property
    def formula_status(self) -> str:
        """功能：读取公式单元格当前的计算状态。

        使用方法：``status = worksheet["C3"].formula_status``。
        参数：无，只读属性。
        返回：``empty``、``pending``、``calculated`` 或 ``error``；分别表示没有
        公式、等待计算、已有缓存结果或最近一次计算失败。
        """
        coordinate = (self._row, self._column)
        if self.formula is None:
            return "empty"
        if coordinate in self._worksheet._formula_errors:
            return "error"
        if coordinate in self._worksheet._formula_values:
            return "calculated"
        return "pending"

    @property
    def dependencies(self) -> tuple["Cell", ...]:
        """功能：取得当前公式直接依赖的单元格快照对象。

        使用方法：``for source in worksheet["C3"].dependencies: print(source.address)``。
        参数：无。
        返回：按公式出现顺序去重的 ``Cell`` 元组；普通值单元格返回空元组。
        该属性只解析引用，不触发公式计算，也不会写回文件。
        异常：公式中的工作表名称或地址无效时透传对应异常。
        """
        if self.formula is None:
            return ()
        from .calculation import formula_dependencies
        return formula_dependencies(self._worksheet._workbook, self._worksheet, self.formula)

    @property
    def dependents(self) -> tuple["Cell", ...]:
        """功能：查找当前单元格被哪些公式单元格直接引用。

        使用方法：``for item in worksheet["A1"].dependents: print(item.address)``。
        参数：无。
        返回：工作簿中按工作表顺序、地址顺序排列的公式 ``Cell`` 元组；没有引用
        时返回空元组。只解析依赖关系，不触发公式计算。
        异常：依赖公式含有无效工作表或地址时透传对应异常。
        """
        result = []
        for worksheet in self._worksheet._workbook.sheets:
            for row, column in sorted(worksheet._formulas):
                candidate = worksheet.cell(row, column)
                if any(source._worksheet is self._worksheet and source.index == self.index
                       for source in candidate.dependencies):
                    result.append(candidate)
        return tuple(result)

    @property
    def calculation_error(self) -> Optional[str]:
        """功能：读取公式最近一次 Python 计算失败的错误说明。

        使用方法：``message = worksheet["C3"].calculation_error``。
        参数：无，只读属性。
        返回：错误说明字符串；没有公式或最近计算未失败时返回 ``None``。
        """
        return self._worksheet._formula_errors.get((self._row, self._column))

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
        """功能：读取当前值的不可变快照，用于不写回的链式类型转换。

        使用方法：``number = cell.read().as_int()``；与直接 ``cell.as_int()``
        不同，转换结果不会修改工作表，也不会影响之后的文件保存内容。公式单元格
        优先读取 Excel/WPS 保存的 ``cached_value``，没有缓存时读取到 ``None``。
        参数：无。
        返回：包含读取瞬间普通值或公式缓存结果的 :class:`CellValue`。
        """
        value = self.cached_value if self.formula is not None else self.value
        return CellValue(value)

    def _ensure_no_formula(self) -> None:
        """功能：阻止写回型类型转换覆盖公式单元格。

        使用方法：由 ``Cell.as_*`` 方法在转换前内部调用。
        参数：无；检查当前单元格是否存在公式。
        返回：``None``；没有公式时允许继续转换。
        异常：公式单元格抛出 ``ValueError``，并保持公式与缓存结果不变。
        """
        if self.formula is not None:
            raise ValueError("公式单元格不能使用写回型 as_*()；请使用 cell.read().as_*()")

    def as_string(self) -> str:
        """功能：转换当前值为字符串、写回单元格并返回转换结果。

        使用方法：``text = cell.set_value(123).as_string()``。
        参数：无；``None`` 转为空字符串，其余值使用 ``str()``。
        返回：写回后的 ``str``。
        异常：公式单元格抛出 ``ValueError``；其他转换失败时透传异常并保持原值不变。
        """
        self._ensure_no_formula()
        converted = _convert_string(self.value)
        self.value = converted
        return converted

    def as_int(self) -> int:
        """功能：无损转换当前值为整数、写回单元格并返回转换结果。

        使用方法：``number = cell.set_value("123").as_int()``。
        参数：无；接受整数、整数字符串和无小数部分的有限浮点数。
        返回：写回后的 ``int``。
        异常：公式单元格抛出 ``ValueError``；无法无损转换时抛出 ``TypeError`` 或
        ``ValueError``，原值不变。
        """
        self._ensure_no_formula()
        converted = _convert_int(self.value)
        self.value = converted
        return converted

    def as_float(self) -> float:
        """功能：转换当前值为有限浮点数、写回单元格并返回转换结果。

        使用方法：``number = cell.set_value("12.5").as_float()``。
        参数：无；接受字符串、整数或浮点数，不接受布尔值。
        返回：写回后的有限 ``float``。
        异常：公式单元格抛出 ``ValueError``；类型、格式无效或结果非有限时抛出
        异常，原值不变。
        """
        self._ensure_no_formula()
        converted = _convert_float(self.value)
        self.value = converted
        return converted

    def as_bool(self) -> bool:
        """功能：转换当前值为布尔值、写回单元格并返回转换结果。

        使用方法：``flag = cell.set_value("是").as_bool()``。
        参数：无；仅接受统一真假值白名单中的文本、布尔值或 0/1。
        返回：写回后的 ``bool``。
        异常：公式单元格抛出 ``ValueError``；类型或取值无效时抛出异常，原值不变。
        """
        self._ensure_no_formula()
        converted = _convert_bool(self.value)
        self.value = converted
        return converted

    def as_date(self) -> date:
        """功能：转换当前值为日期、写回单元格并返回转换结果。

        使用方法：``day = cell.set_value("2026/8/1").as_date()``。
        参数：无；接受支持格式的字符串、``date`` 或 ``datetime``。
        返回：写回后的 :class:`date`；``datetime`` 会舍弃时间部分。
        异常：公式单元格抛出 ``ValueError``；类型、格式或日期取值无效时抛出异常，
        原值不变。
        """
        self._ensure_no_formula()
        converted = _convert_date(self.value)
        self.value = converted
        return converted

    def as_datetime(self) -> datetime:
        """功能：转换当前值为日期时间、写回单元格并返回转换结果。

        使用方法：``moment = cell.set_value("2026-8-1 12:33").as_datetime()``。
        参数：无；接受支持格式的字符串、``date`` 或 ``datetime``。
        返回：写回后的 :class:`datetime`；纯日期补为午夜时间。
        异常：公式单元格抛出 ``ValueError``；类型、格式或日期时间无效时抛出异常，
        原值不变。
        """
        self._ensure_no_formula()
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
    def formula(self, formula: Optional[str]) -> None:
        """功能：设置或清除单元格公式；设置公式会清除同一位置的普通值。

        使用方法：``worksheet["D2"].formula = "=SUM(B2:C2)"``；开头的 ``=``
        可以省略；``worksheet["D2"].formula = None`` 清除公式。
        参数：``formula`` 为包含表达式的非空字符串或 ``None``。
        返回：``None``。
        异常：非空公式的类型或内容无效时抛出 ``TypeError``。
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

    def copy_style(self, source: "Cell") -> "Cell":
        """功能：从另一个单元格复制完整样式到当前目标单元格。

        使用方法：``target.copy_style(source)``，例如
        ``worksheet["B1"].copy_style(worksheet["A1"])``。
        参数：``source`` 必须是 :class:`Cell`；可以来自其他工作表或工作簿。
        返回：当前目标 :class:`Cell`，支持继续设置 ``value`` 等属性。
        异常：参数不是 ``Cell`` 时抛出 ``TypeError``。只复制样式，不复制值、公式
        或公式缓存结果。
        """
        if not isinstance(source, Cell):
            raise TypeError("source 必须是 Cell")
        self.style = source.style
        return self

    def __repr__(self) -> str:
        """功能：生成用于调试的单元格文本表示。

        使用方法：``repr(cell)``。
        参数：无。
        返回：包含工作表名、A1 地址、普通值、缓存结果和公式的字符串。
        """
        return (
            f"<Cell {self._worksheet.name}!{self.address} "
            f"value={self.value!r} cached_value={self.cached_value!r} "
            f"formula={self.formula!r}>"
        )
