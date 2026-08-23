"""工作表对象及单元格、区域访问接口。"""

from __future__ import annotations

from collections.abc import Iterable
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
from .conversion import normalize_value
from .dimension import ColumnDimension, RowDimension
from .page import PageSettings
from .range import Range

if TYPE_CHECKING:
    from .workbook import Workbook


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
    return normalize_value(value)


class Worksheet:
    """表示隶属于某个 :class:`Workbook` 的命名工作表。"""

    __slots__ = (
        "_workbook",
        "_label",
        "_label_color",
        "_values",
        "_formulas",
        "_styles",
        "_merged_ranges",
        "_rows",
        "_columns",
        "_freeze",
        "_filter_range",
        "_show_gridlines",
        "_page",
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
        self._merged_ranges: list[Tuple[int, int, int, int]] = []
        self._rows: Dict[int, RowDimension] = {}
        self._columns: Dict[int, ColumnDimension] = {}
        self._freeze: Optional[str] = None
        self._filter_range: Optional[str] = None
        self._show_gridlines = True
        self._page = PageSettings()
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

    def row(self, index: int) -> RowDimension:
        """功能：按0-based索引取得可设置行高和隐藏状态的行对象。

        使用方法：``worksheet.row(0).height = 28``。
        参数：``index`` 为0～1048575的整数，布尔值不作为索引。
        返回：当前行唯一的 :class:`RowDimension` 对象，重复读取返回同一实例。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(index)
        if index not in self._rows:
            self._rows[index] = RowDimension(index)
        return self._rows[index]

    def column(self, index: int) -> ColumnDimension:
        """功能：按0-based索引取得可设置列宽和隐藏状态的列对象。

        使用方法：``worksheet.column(0).width = 20``。
        参数：``index`` 为0～16383的整数，布尔值不作为索引。
        返回：当前列唯一的 :class:`ColumnDimension` 对象，重复读取返回同一实例。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_column_index(index)
        if index not in self._columns:
            self._columns[index] = ColumnDimension(index)
        return self._columns[index]

    @property
    def merged_ranges(self) -> tuple[Range, ...]:
        """功能：取得全部合并区域的只读顺序快照。

        使用方法：``for area in worksheet.merged_ranges: print(area.address)``。
        参数：无；修改合并状态使用 ``Range.merge()`` 或 ``Range.unmerge()``。
        返回：按左上角位置排序的 ``tuple[Range, ...]``。
        """
        return tuple(Range(self, *bounds) for bounds in self._merged_ranges)

    @property
    def freeze(self) -> Optional[str]:
        """功能：读取冻结窗格后的第一个可滚动单元格地址。

        使用方法：``address = worksheet.freeze``。
        参数：无。
        返回：大写A1地址或没有冻结窗格时的 ``None``。
        """
        return self._freeze

    @freeze.setter
    def freeze(self, address: Optional[str]) -> None:
        """功能：设置或清除冻结行列。

        使用方法：``worksheet.freeze = "B2"`` 冻结第一行和第一列；赋值
        ``None`` 或 ``"A1"`` 清除冻结。
        参数：``address`` 为第一个可滚动单元格的A1地址或 ``None``。
        返回：``None``。
        异常：地址类型、格式或边界无效时抛出 ``InvalidAddressError``。
        """
        if address is None:
            self._freeze = None
            return
        row, column = cell_index(address)
        self._freeze = None if (row, column) == (0, 0) else cell_address(row, column)

    @property
    def filter_range(self) -> Optional[str]:
        """功能：读取工作表自动筛选区域。

        使用方法：``address = worksheet.filter_range``。
        参数：无。
        返回：大写A1矩形区域或未启用筛选时的 ``None``。
        """
        return self._filter_range

    @filter_range.setter
    def filter_range(self, address: Optional[str]) -> None:
        """功能：设置或清除一块连续区域的自动筛选按钮。

        使用方法：``worksheet.filter_range = "A1:F100"``；赋值 ``None`` 清除。
        参数：``address`` 为合法A1矩形区域字符串或 ``None``。
        返回：``None``。
        异常：地址无效时抛出 ``InvalidAddressError``。
        """
        if address is None:
            self._filter_range = None
            return
        self._filter_range = Range(self, *parse_range(address)).address

    @property
    def show_gridlines(self) -> bool:
        """功能：读取Excel屏幕是否显示工作表网格线。

        使用方法：``visible = worksheet.show_gridlines``。
        参数：无。
        返回：布尔值；此属性不控制打印网格线。
        """
        return self._show_gridlines

    @show_gridlines.setter
    def show_gridlines(self, value: bool) -> None:
        """功能：设置Excel屏幕中的工作表网格线可见性。

        使用方法：``worksheet.show_gridlines = False``。
        参数：``value`` 必须是布尔值。
        返回：``None``。
        异常：类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("show_gridlines 必须是布尔值")
        self._show_gridlines = value

    @property
    def page(self) -> PageSettings:
        """功能：取得当前工作表唯一的页面和打印设置对象。

        使用方法：``worksheet.page.orientation = "landscape"``。
        参数：无；属性本身只读，不允许整体替换。
        返回：当前 :class:`PageSettings`。
        """
        return self._page

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

    @staticmethod
    def _prepare_append_row(values: Iterable[Any], method_name: str) -> list[Any]:
        """功能：完整读取并规范化一行待追加数据。

        使用方法：由 ``append()`` 和 ``append_rows()`` 在写入前共同调用。
        参数：``values`` 为一维可迭代数据；``method_name`` 用于生成明确错误消息。
        返回：已经完成日期字面量转换的普通值列表。
        异常：整行是字符串、字节或不可迭代对象时抛出 ``TypeError``；值转换失败
        时透传相应异常，工作表尚未发生改变。
        """
        if isinstance(values, (str, bytes)):
            raise TypeError(f"{method_name} 需要一维行数据，不能直接传入字符串或字节对象")
        try:
            row_values = list(values)
        except TypeError as error:
            raise TypeError(f"{method_name} 需要一维可迭代对象") from error
        return [_normalize_value(value) for value in row_values]

    def append(self, values: Iterable[Any]) -> "Worksheet":
        """功能：在当前最大行索引之后追加一行普通值。

        使用方法：``worksheet.append(["姓名", "成绩"])``；空表从 A1 开始。
        参数：``values`` 为一维可迭代对象，元素按 0-based 列索引从左到右写入；
        字符串和字节对象不能作为整行数据，空可迭代对象不会触及新行。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：参数不可迭代或是字符串、字节对象时抛出 ``TypeError``；数据超过
        Excel 行列上限时抛出 ``InvalidAddressError``；值转换失败或目标位于合并
        区域非左上角时保持整行写入前状态。
        """
        row_values = self._prepare_append_row(values, "append()")
        if not row_values:
            return self

        target_row = self._max_row + 1
        validate_row_index(target_row)
        validate_column_index(len(row_values) - 1)
        for column in range(len(row_values)):
            anchor = self._merged_anchor(target_row, column)
            if anchor is not None and anchor != (target_row, column):
                raise ValueError("不能向合并区域的非左上角单元格追加值")
        for column, value in enumerate(row_values):
            self._set_value(target_row, column, value)
        return self

    def append_rows(self, rows: Iterable[Iterable[Any]]) -> "Worksheet":
        """功能：按给定顺序连续追加多行普通值。

        使用方法：``worksheet.append_rows([["张三", 90], ["李四", 88]])``。
        参数：``rows`` 为二维可迭代对象；每个非空元素会作为一行传给
        :meth:`append`，空行与单行 ``append([])`` 一样不会推进位置。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：外层或任一行不可迭代时抛出 ``TypeError``；超过 Excel 上限时抛出
        ``InvalidAddressError``。全部行和值会在首次写入前完成验证，任何失败都不会
        留下前置行或部分单元格。
        """
        if isinstance(rows, (str, bytes)):
            raise TypeError("append_rows() 需要二维数据，不能直接传入字符串或字节对象")
        try:
            input_rows = list(rows)
        except TypeError as error:
            raise TypeError("append_rows() 需要二维可迭代对象") from error
        prepared_rows = [
            self._prepare_append_row(values, "append_rows()")
            for values in input_rows
        ]
        prepared_rows = [values for values in prepared_rows if values]
        if not prepared_rows:
            return self

        first_row = self._max_row + 1
        validate_row_index(first_row + len(prepared_rows) - 1)
        for row_offset, row_values in enumerate(prepared_rows):
            validate_column_index(len(row_values) - 1)
            target_row = first_row + row_offset
            for column in range(len(row_values)):
                anchor = self._merged_anchor(target_row, column)
                if anchor is not None and anchor != (target_row, column):
                    raise ValueError("不能向合并区域的非左上角单元格追加值")
        for row_values in prepared_rows:
            self.append(row_values)
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

    def _merged_anchor(self, row: int, column: int) -> Optional[Tuple[int, int]]:
        """功能：查询坐标所属合并区域的左上角锚点。

        使用方法：普通值和公式写入前内部调用。
        参数：``row``、``column`` 为0-based行列索引，顺序先行后列。
        返回：属于合并区域时返回锚点元组，否则返回 ``None``。
        """
        for min_row, min_column, max_row, max_column in self._merged_ranges:
            if min_row <= row <= max_row and min_column <= column <= max_column:
                return min_row, min_column
        return None

    def _merge_range(
        self, min_row: int, min_column: int, max_row: int, max_column: int
    ) -> None:
        """功能：原子登记一个不重叠的多单元格合并区域。

        使用方法：仅由 ``Range.merge()`` 调用。
        参数：四项为0-based最小行、最小列、最大行、最大列，顺序先行后列。
        返回：``None``；完全相同的区域重复合并视为幂等操作。
        异常：单格区域、重叠区域或非锚点单元格存在值或公式时抛出 ``ValueError``。
        """
        bounds = (min_row, min_column, max_row, max_column)
        if min_row == max_row and min_column == max_column:
            raise ValueError("合并区域必须至少包含两个单元格")
        for existing in self._merged_ranges:
            if existing == bounds:
                return
            a, b, c, d = existing
            overlaps = not (
                max_row < a or min_row > c or max_column < b or min_column > d
            )
            if overlaps:
                raise ValueError("合并区域不能与已有合并区域重叠")
        for row in range(min_row, max_row + 1):
            for column in range(min_column, max_column + 1):
                if (row, column) == (min_row, min_column):
                    continue
                if (
                    self._values.get(row, column) is not None
                    or (row, column) in self._formulas
                ):
                    raise ValueError("合并前除左上角外的单元格必须为空")
        self._merged_ranges.append(bounds)
        self._merged_ranges.sort()
        self._touch(max_row, max_column)

    def _unmerge_range(
        self, min_row: int, min_column: int, max_row: int, max_column: int
    ) -> None:
        """功能：删除与给定边界完全相同的合并区域记录。

        使用方法：仅由 ``Range.unmerge()`` 调用。
        参数：四项为0-based区域边界，顺序先行后列。
        返回：``None``；单元格内容和样式不改变。
        异常：区域没有被完整合并时抛出 ``ValueError``。
        """
        bounds = (min_row, min_column, max_row, max_column)
        if bounds not in self._merged_ranges:
            raise ValueError("当前区域不是一个完整的合并区域")
        self._merged_ranges.remove(bounds)

    def _set_value(self, row: int, column: int, value: Any) -> None:
        """功能：设置普通值并清除同一位置的公式。

        使用方法：由 :attr:`Cell.value`、区域写入和追加方法内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列；
        ``value`` 为任意 Python 对象，``None`` 表示删除普通值。
        返回：``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        anchor = self._merged_anchor(row, column)
        if anchor is not None and anchor != (row, column):
            raise ValueError("只能向合并区域的左上角单元格写入值")
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

    def _set_formula(
        self, row: int, column: int, formula: Optional[str]
    ) -> None:
        """功能：校验并设置公式，或使用 ``None`` 清除现有公式。

        使用方法：通过 ``worksheet["A1"].formula = "=SUM(B1:B5)"`` 间接调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列；
        ``formula`` 为非空字符串或 ``None``，字符串可以包含或省略开头的 ``=``。
        返回：``None``；字符串统一保存为带前导 ``=`` 的形式，``None`` 仅删除公式。
        异常：非空公式不是字符串、为空或只有 ``=`` 时抛出 ``TypeError``；索引
        无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        if formula is None:
            self._formulas.pop((row, column), None)
            return
        anchor = self._merged_anchor(row, column)
        if anchor is not None and anchor != (row, column):
            raise ValueError("只能向合并区域的左上角单元格写入公式")
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
