"""单元格的普通值和公式访问对象。"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Optional

from ..address import cell_address
from ..style import Style

if TYPE_CHECKING:
    from .worksheet import Worksheet


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
