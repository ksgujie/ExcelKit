"""连续矩形区域的普通值批量访问对象。"""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from copy import deepcopy
from dataclasses import replace
from datetime import date, datetime, timedelta
from typing import TYPE_CHECKING, Any, List

from ..address import cell_address, validate_row_index
from ..autofill import AutoFillMode
from ..style import DEFAULT_STYLE, Style

if TYPE_CHECKING:
    from .worksheet import Worksheet


class Range:
    """表示由 0-based 起止索引确定的连续矩形区域。"""

    __slots__ = ("_worksheet", "_min_row", "_min_column", "_max_row", "_max_column")

    def __init__(
        self,
        worksheet: "Worksheet",
        min_row: int,
        min_column: int,
        max_row: int,
        max_column: int,
    ) -> None:
        """功能：创建指定矩形边界的区域对象。

        使用方法：由 ``worksheet.range("A1:C10")`` 创建，不直接调用。
        参数：``worksheet`` 为所属工作表；其余参数依次为最小行、最小列、最大行、
        最大列的 0-based 索引，整体顺序为先行后列。
        返回：无。
        """
        self._worksheet = worksheet
        self._min_row = min_row
        self._min_column = min_column
        self._max_row = max_row
        self._max_column = max_column

    @property
    def min_row(self) -> int:
        """功能：取得区域最小 0-based 行索引。

        使用方法：``worksheet.range("B3:D8").min_row`` 返回 ``2``。
        参数：无。
        返回：0-based 整数行索引。
        """
        return self._min_row

    @property
    def worksheet(self) -> "Worksheet":
        """功能：取得当前区域所属工作表。

        使用方法：``worksheet = area.worksheet``。
        参数：无，只读属性。
        返回：创建当前区域的 :class:`Worksheet`。
        """
        return self._worksheet

    @property
    def min_column(self) -> int:
        """功能：取得区域最小 0-based 列索引。

        使用方法：``worksheet.range("B3:D8").min_column`` 返回 ``1``。
        参数：无。
        返回：0-based 整数列索引。
        """
        return self._min_column

    @property
    def max_row(self) -> int:
        """功能：取得区域最大 0-based 行索引。

        使用方法：``worksheet.range("B3:D8").max_row`` 返回 ``7``。
        参数：无。
        返回：0-based 整数行索引。
        """
        return self._max_row

    @property
    def max_column(self) -> int:
        """功能：取得区域最大 0-based 列索引。

        使用方法：``worksheet.range("B3:D8").max_column`` 返回 ``3``。
        参数：无。
        返回：0-based 整数列索引。
        """
        return self._max_column

    @property
    def address(self) -> str:
        """功能：取得区域规范化的大写A1地址。

        使用方法：``worksheet.range("a1:c3").address`` 返回 ``"A1:C3"``。
        参数：无。
        返回：包含起止单元格的A1区域字符串。
        """
        start = cell_address(self._min_row, self._min_column)
        end = cell_address(self._max_row, self._max_column)
        return f"{start}:{end}"

    @property
    def values(self) -> List[List[Any]]:
        """功能：读取区域内的全部有效值。

        使用方法：``values = worksheet.range("A1:C2").values``。
        参数：无。
        返回：按行组织的二维 ``list``；公式单元格返回当前公式结果，空单元格或没有
        有效结果的公式返回 ``None``。
        """
        return [
            [
                self._worksheet.cell(row, column).value
                for column in range(self._min_column, self._max_column + 1)
            ]
            for row in range(self._min_row, self._max_row + 1)
        ]

    @property
    def is_empty(self) -> bool:
        """功能：判断区域是否没有任何有效内容。

        使用方法：``worksheet.range("A1:C3").is_empty``。
        参数：无，只读属性；普通值、公式（即使结果为空）、样式、链接和批注中，
        仅检查普通值与公式是否存在。
        返回：区域内没有普通值和公式时为 ``True``，否则为 ``False``。
        """
        coordinates = set(self._worksheet._values._values) | set(self._worksheet._formulas)
        return not any(
            self._min_row <= row <= self._max_row
            and self._min_column <= column <= self._max_column
            for row, column in coordinates
        )

    def transpose_to(self, target: "Range") -> "Range":
        """功能：将当前区域的有效值转置写入目标区域。

        使用方法：``source.transpose_to(worksheet.range("E1:G2"))``。
        参数：``target`` 为同一工作簿中、大小应等于源区域转置尺寸的目标区域。
        返回：目标 :class:`Range`，便于继续链式调用；公式会写入为当前计算结果。
        异常：目标类型、工作簿归属或尺寸不匹配时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(target, Range):
            raise TypeError("target 必须是 Range")
        if target.worksheet._workbook is not self._worksheet._workbook:
            raise ValueError("目标区域必须属于同一工作簿")
        source_rows = self._max_row - self._min_row + 1
        source_columns = self._max_column - self._min_column + 1
        target_rows = target._max_row - target._min_row + 1
        target_columns = target._max_column - target._min_column + 1
        if (target_rows, target_columns) != (source_columns, source_rows):
            raise ValueError("目标区域尺寸必须等于源区域转置后的尺寸")
        matrix = self.values
        transposed = [list(row) for row in zip(*matrix)]
        return target.set_values(transposed)

    def to_records(
        self,
        *,
        headers: bool | Sequence[str] = True,
        header_row: int | None = None,
    ) -> list[dict[str, Any]]:
        """功能：把区域按行转换为字典记录列表。

        使用方法：``records = worksheet.range("A1:C100").to_records()``；如果区域
        没有表头，使用 ``to_records(headers=False)`` 自动生成 ``Column1``、
        ``Column2`` 等字段名；字段在区域外的 0-based 第1行时使用
        ``to_records(header_row=1)``；也可向 ``headers`` 传入显式字段名序列。
        参数：``headers`` 为布尔值或字段名序列；``True`` 时默认使用区域首行，
        ``False`` 时生成字段名，序列则直接作为字段名且区域全部行都是数据；
        ``header_row`` 为工作表绝对 0-based 字段行索引，并使用当前区域相同的列。
        返回：按区域原始行顺序排列的 ``list[dict[str, Any]]``；只有表头而没有数据
        时返回空列表。
        异常：参数组合、字段数量、类型、空值或重复字段无效时抛出 ``TypeError`` 或
        ``ValueError``；``header_row`` 位于区域中间时也会报错。
        """
        values = self.values
        width = self._max_column - self._min_column + 1
        if isinstance(headers, bool):
            if not headers:
                if header_row is not None:
                    raise ValueError("headers=False 时不能同时设置 header_row")
                names = tuple(f"Column{index}" for index in range(1, width + 1))
            else:
                if header_row is None:
                    header_values = values[0]
                    values = values[1:]
                else:
                    validate_row_index(header_row)
                    if self._min_row < header_row <= self._max_row:
                        raise ValueError("header_row 不能位于数据区域中间")
                    header_values = [
                        self._worksheet.cell(header_row, column).value
                        for column in range(self._min_column, self._max_column + 1)
                    ]
                    if header_row == self._min_row:
                        values = values[1:]
                names = tuple(header_values)
        else:
            if header_row is not None:
                raise ValueError("显式 headers 不能与 header_row 同时使用")
            if isinstance(headers, (str, bytes)):
                raise TypeError("headers 必须是 bool 或字段名序列")
            try:
                names = tuple(headers)
            except TypeError as error:
                raise TypeError("headers 必须是 bool 或字段名序列") from error
        if len(names) != width:
            raise ValueError(f"headers 必须包含 {width} 个字段名")
        if any(not isinstance(value, str) or not value for value in names):
            raise ValueError("表头必须全部是非空字符串")
        if len(set(names)) != len(names):
            raise ValueError("表头不能重复")
        return [dict(zip(names, row_values)) for row_values in values]

    @property
    def format(self) -> "RangeFormat":
        """功能：取得当前区域的批量格式代理。

        使用方法：``worksheet.range('C2:C20').format.number = '0.00%'``。
        参数：无；返回的代理只作用于当前 ``Range``。
        返回：可设置 ``number`` 属性的 :class:`RangeFormat`。
        """
        return RangeFormat(self)

    def apply_style(self, style: Style) -> "Range":
        """功能：把同一个完整单元格样式应用到当前区域。

        使用方法：``worksheet.range('A1:D1').apply_style(header_style)``。
        参数：``style`` 必须是 :class:`Style`；会覆盖区域内原有完整样式。
        返回：当前区域，支持继续链式调用。
        异常：``style`` 类型无效时抛出 ``TypeError``。
        """
        if not isinstance(style, Style):
            raise TypeError("style 必须是 Style")
        for row in range(self._min_row, self._max_row + 1):
            for column in range(self._min_column, self._max_column + 1):
                self._worksheet._styles[(row, column)] = style
                self._worksheet._touch(row, column)
        return self

    def set_values(self, values: Iterable[Iterable[Any]]) -> "Range":
        """功能：校验并批量写入与区域形状完全一致的二维普通值。

        使用方法：``worksheet.range("A1:B2").set_values([[1, 2], [3, 4]])``。
        参数：``values`` 为二维可迭代对象；行数和每行列数必须与区域完全一致，
        字符串和字节对象不能充当外层数据或单独一行。
        返回：当前 :class:`Range`，支持链式调用。
        异常：数据不可迭代、不是二维结构或形状不匹配时抛出 ``ValueError``。
        """
        if isinstance(values, (str, bytes)):
            raise ValueError("区域数据必须是二维可迭代对象")
        try:
            input_rows = list(values)
        except TypeError as error:
            raise ValueError("区域数据必须是二维可迭代对象") from error

        expected_rows = self._max_row - self._min_row + 1
        expected_columns = self._max_column - self._min_column + 1
        if len(input_rows) != expected_rows:
            raise ValueError(f"区域需要 {expected_rows} 行，实际收到 {len(input_rows)} 行")

        materialized: List[List[Any]] = []
        for row_offset, input_row in enumerate(input_rows):
            if isinstance(input_row, (str, bytes)):
                raise ValueError(f"第 {row_offset} 行必须是一维可迭代对象")
            try:
                row_values = list(input_row)
            except TypeError as error:
                raise ValueError(f"第 {row_offset} 行必须是一维可迭代对象") from error
            if len(row_values) != expected_columns:
                raise ValueError(
                    f"第 {row_offset} 行需要 {expected_columns} 个值，"
                    f"实际收到 {len(row_values)} 个"
                )
            materialized.append(row_values)

        # 完整形状验证通过后才开始写入，避免尺寸错误造成部分覆盖。
        for row_offset, row_values in enumerate(materialized):
            for column_offset, _value in enumerate(row_values):
                row = self._min_row + row_offset
                column = self._min_column + column_offset
                anchor = self._worksheet._merged_anchor(row, column)
                if anchor is not None and anchor != (row, column):
                    raise ValueError("不能向合并区域的非左上角单元格批量写入值")
        for row_offset, row_values in enumerate(materialized):
            for column_offset, value in enumerate(row_values):
                self._worksheet._set_value(
                    self._min_row + row_offset,
                    self._min_column + column_offset,
                    value,
                )
        return self

    def clear(
        self,
        *,
        values: bool = True,
        styles: bool = True,
        hyperlinks: bool = False,
        notes: bool = False,
    ) -> "Range":
        """功能：按开关清除区域内内容、样式、超链接和批注。

        使用方法：``worksheet.range("A1:C10").clear()`` 清除值、公式和样式；
        ``clear(values=True, styles=False)`` 只清内容；``clear(values=False,
        styles=True)`` 只清样式；附属信息使用对应开关。
        参数：``values`` 控制普通值、公式与公式缓存；``styles`` 控制单元格样式；
        ``hyperlinks``、``notes`` 默认关闭，以保持旧版本行为。所有参数必须为布尔值。
        返回：当前 :class:`Range`，支持链式调用。
        异常：任一开关不是布尔值时抛出 ``TypeError``。
        """
        for name, enabled in (
            ("values", values), ("styles", styles),
            ("hyperlinks", hyperlinks), ("notes", notes),
        ):
            if not isinstance(enabled, bool):
                raise TypeError(f"{name} 必须是 bool")
        if values:
            self._worksheet._workbook._invalidate_formula_caches()
        for row in range(self._min_row, self._max_row + 1):
            for column in range(self._min_column, self._max_column + 1):
                coordinate = (row, column)
                if values:
                    self._worksheet._values.set(row, column, None)
                    self._worksheet._formulas.pop(coordinate, None)
                    self._worksheet._formula_values.pop(coordinate, None)
                    self._worksheet._formula_errors.pop(coordinate, None)
                if styles:
                    self._worksheet._styles.pop(coordinate, None)
                if hyperlinks:
                    self._worksheet._hyperlinks.pop(coordinate, None)
                if notes:
                    self._worksheet._notes.pop(coordinate, None)
        return self

    def copy_to(
        self,
        target: "Range",
        *,
        values: bool = True,
        formulas: bool = True,
        styles: bool = True,
    ) -> "Range":
        """功能：把当前区域的值、公式和样式复制到同尺寸目标区域。

        使用方法：``source.copy_to(target)``；仅复制样式可传入
        ``values=False, formulas=False, styles=True``。
        参数：``target`` 为同尺寸 :class:`Range`；三个布尔开关分别控制普通值、
        公式和样式。公式按源目标行列偏移调整相对 A1 引用；当 ``formulas=False``
        且 ``values=True`` 时，公式单元格复制现有缓存结果作为普通值。
        返回：目标 :class:`Range`，支持继续操作目标区域。
        异常：参数类型、区域尺寸或开关类型无效，以及目标合并区域禁止写入时抛出
        ``TypeError`` 或 ``ValueError``；失败时目标区域保持不变。
        """
        if not isinstance(target, Range):
            raise TypeError("target 必须是 Range")
        for name, enabled in (
            ("values", values), ("formulas", formulas), ("styles", styles)
        ):
            if not isinstance(enabled, bool):
                raise TypeError(f"{name} 必须是布尔值")
        source_shape = (
            self._max_row - self._min_row + 1,
            self._max_column - self._min_column + 1,
        )
        target_shape = (
            target._max_row - target._min_row + 1,
            target._max_column - target._min_column + 1,
        )
        if source_shape != target_shape:
            raise ValueError("源区域和目标区域尺寸必须完全一致")

        snapshots = []
        for row_offset in range(source_shape[0]):
            for column_offset in range(source_shape[1]):
                source_coordinate = (
                    self._min_row + row_offset,
                    self._min_column + column_offset,
                )
                target_coordinate = (
                    target._min_row + row_offset,
                    target._min_column + column_offset,
                )
                if values or formulas:
                    anchor = target._worksheet._merged_anchor(*target_coordinate)
                    if anchor is not None and anchor != target_coordinate:
                        raise ValueError("不能复制到合并区域的非左上角单元格")
                formula = self._worksheet._formulas.get(source_coordinate)
                ordinary = self._worksheet._values.get(*source_coordinate)
                cached = self._worksheet._formula_values.get(source_coordinate)
                style = self._worksheet._styles.get(source_coordinate)
                snapshots.append(
                    (
                        row_offset,
                        column_offset,
                        deepcopy(ordinary),
                        formula,
                        deepcopy(cached),
                        style,
                    )
                )

        row_delta = target._min_row - self._min_row
        column_delta = target._min_column - self._min_column
        if formulas:
            from ..template import _translate_formula

            snapshots = [
                (
                    row_offset,
                    column_offset,
                    ordinary,
                    _translate_formula(formula, row_delta, column_delta)
                    if formula is not None else None,
                    cached,
                    style,
                )
                for row_offset, column_offset, ordinary, formula, cached, style
                in snapshots
            ]
        if values or formulas:
            target._worksheet._workbook._invalidate_formula_caches()

        for row_offset, column_offset, ordinary, formula, cached, style in snapshots:
            row = target._min_row + row_offset
            column = target._min_column + column_offset
            coordinate = (row, column)
            if formulas and formula is not None:
                target._worksheet._values.set(row, column, None)
                target._worksheet._formulas[coordinate] = formula
                target._worksheet._formula_values.pop(coordinate, None)
                target._worksheet._formula_errors.pop(coordinate, None)
                target._worksheet._touch(row, column)
            elif values:
                copied_value = cached if formula is not None else ordinary
                target._worksheet._values.set(row, column, copied_value)
                target._worksheet._formulas.pop(coordinate, None)
                target._worksheet._formula_values.pop(coordinate, None)
                target._worksheet._formula_errors.pop(coordinate, None)
                target._worksheet._touch(row, column)
            if styles:
                if style is None:
                    target._worksheet._styles.pop(coordinate, None)
                else:
                    target._worksheet._styles[coordinate] = style
                    target._worksheet._touch(row, column)
        return target

    def auto_fill(
        self, target: str | "Range", *, mode: str = AutoFillMode.AUTO
    ) -> "Range":
        """功能：按 Excel 填充柄语义向包含源区域的目标区域扩展内容和格式。

        使用方法：``ws.range('A1:A2').auto_fill('A1:A20')``；公式、样式、超链接
        和批注会随填充复制。``mode=AutoFillMode.SERIES`` 强制按数值或日期序列
        扩展；``COPY`` 重复源模式；``FORMATS`` 仅复制样式。
        参数：``target`` 为同表 A1 区域或 ``Range``，必须完整包含当前源区域且边界
        从源区域左上角开始；``mode`` 为 ``AutoFillMode`` 固定值。
        返回：目标区域。
        异常：目标不合法、方向不一致或序列无法推导时抛出 ``TypeError`` 或
        ``ValueError``。
        """
        if isinstance(target, str):
            target = self._worksheet.range(target)
        if not isinstance(target, Range):
            raise TypeError("target 必须是 A1 区域字符串或 Range")
        if target.worksheet is not self._worksheet:
            raise ValueError("自动填充的源区域和目标区域必须属于同一工作表")
        if mode not in {
            AutoFillMode.AUTO, AutoFillMode.COPY, AutoFillMode.SERIES,
            AutoFillMode.FORMATS,
        }:
            raise ValueError("mode 必须是 AutoFillMode 的固定值")
        if (
            target.min_row != self.min_row or target.min_column != self.min_column
            or target.max_row < self.max_row or target.max_column < self.max_column
        ):
            raise ValueError("target 必须从源区域左上角开始并完整包含源区域")
        source_height = self.max_row - self.min_row + 1
        source_width = self.max_column - self.min_column + 1
        target_height = target.max_row - target.min_row + 1
        target_width = target.max_column - target.min_column + 1
        if source_height > 1 and source_width > 1 and (
            target_height % source_height or target_width % source_width
        ):
            raise ValueError("二维源区域的目标尺寸必须是源区域尺寸的整数倍")
        snapshots = self._auto_fill_snapshots()
        is_series = mode == AutoFillMode.SERIES or (
            mode == AutoFillMode.AUTO and self._can_fill_series()
        )
        if mode == AutoFillMode.SERIES and not self._can_fill_series():
            raise ValueError("SERIES 模式要求源区域是一行或一列数值、日期或日期时间序列")
        for row in range(target.min_row, target.max_row + 1):
            for column in range(target.min_column, target.max_column + 1):
                if row <= self.max_row and column <= self.max_column:
                    continue
                source_row = self.min_row + (row - self.min_row) % source_height
                source_column = self.min_column + (column - self.min_column) % source_width
                source = snapshots[(source_row, source_column)]
                coordinate = (row, column)
                if mode != AutoFillMode.FORMATS:
                    value, formula, hyperlink, note = source[:4]
                    if is_series:
                        value = self._series_value(row, column)
                        formula = None
                    if formula is not None:
                        from ..template import _translate_formula
                        formula = _translate_formula(
                            formula, row - source_row, column - source_column
                        )
                        self._worksheet._set_formula(row, column, formula)
                    else:
                        self._worksheet._set_value(row, column, deepcopy(value))
                    if hyperlink is None:
                        self._worksheet._hyperlinks.pop(coordinate, None)
                    else:
                        self._worksheet._hyperlinks[coordinate] = deepcopy(hyperlink)
                    if note is None:
                        self._worksheet._notes.pop(coordinate, None)
                    else:
                        self._worksheet._notes[coordinate] = deepcopy(note)
                style = source[4]
                if style is None:
                    self._worksheet._styles.pop(coordinate, None)
                else:
                    self._worksheet._styles[coordinate] = style
                    self._worksheet._touch(row, column)
        return target

    def _auto_fill_snapshots(self) -> dict[tuple[int, int], tuple[Any, Any, Any, Any, Any]]:
        """功能：冻结自动填充源区域状态，避免目标覆盖源数据影响后续计算。

        使用方法：由 :meth:`auto_fill` 在写入目标区域之前内部调用。
        参数：无；读取当前区域边界内的值、公式、超链接、批注和样式。
        返回：以 0-based ``(row, column)`` 为键、五类内容快照为值的字典。
        """
        return {
            (row, column): (
                deepcopy(self._worksheet._values.get(row, column)),
                self._worksheet._formulas.get((row, column)),
                deepcopy(self._worksheet._hyperlinks.get((row, column))),
                deepcopy(self._worksheet._notes.get((row, column))),
                self._worksheet._styles.get((row, column)),
            )
            for row in range(self.min_row, self.max_row + 1)
            for column in range(self.min_column, self.max_column + 1)
        }

    def _can_fill_series(self) -> bool:
        """功能：判断源区域是否为可自动扩展的单行或单列数值/日期序列。

        使用方法：由 :meth:`auto_fill` 的自动模式和序列模式内部调用。
        参数：无；判断当前区域的一格或两格普通值。
        返回：可以按序列扩展时返回 ``True``，否则返回 ``False``。
        """
        height = self.max_row - self.min_row + 1
        width = self.max_column - self.min_column + 1
        if height != 1 and width != 1:
            return False
        values = [
            self._worksheet._values.get(row, column)
            for row in range(self.min_row, self.max_row + 1)
            for column in range(self.min_column, self.max_column + 1)
        ]
        return len(values) in {1, 2} and all(
            isinstance(value, (int, float, date, datetime)) and not isinstance(value, bool)
            for value in values
        )

    def _series_value(self, row: int, column: int) -> Any:
        """功能：根据一格或两格数值/日期源计算目标位置的序列值。

        使用方法：由 :meth:`auto_fill` 为每个目标位置内部调用。
        参数：``row``、``column`` 为目标单元格的 0-based 索引，顺序为先行后列。
        返回：按源间距推算并复制得到的数值、日期或日期时间。
        """
        source_values = [
            self._worksheet._values.get(source_row, source_column)
            for source_row in range(self.min_row, self.max_row + 1)
            for source_column in range(self.min_column, self.max_column + 1)
        ]
        offset = (row - self.min_row) if self.max_row > self.min_row else (column - self.min_column)
        if len(source_values) == 1:
            return deepcopy(source_values[0])
        first, second = source_values
        delta = second - first
        return first + delta * offset

    def merge(self) -> "Range":
        """功能：合并当前矩形区域并保留左上角单元格内容。

        使用方法：``worksheet.range("A1:C1").merge()``。
        参数：无，区域边界在创建 ``Range`` 时已经验证。
        返回：当前 :class:`Range`，支持链式调用。
        异常：区域只有一个单元格、与已有合并区域重叠，或除左上角外存在值或公式时
        抛出 ``ValueError``，失败时工作表保持不变。
        """
        self._worksheet._merge_range(
            self._min_row, self._min_column, self._max_row, self._max_column
        )
        return self

    def unmerge(self) -> "Range":
        """功能：取消与当前边界完全相同的合并区域。

        使用方法：``worksheet.range("A1:C1").unmerge()``。
        参数：无。
        返回：当前 :class:`Range`，左上角内容保持不变。
        异常：当前边界不是一个完整合并区域时抛出 ``ValueError``。
        """
        self._worksheet._unmerge_range(
            self._min_row, self._min_column, self._max_row, self._max_column
        )
        return self

    def remove_duplicates(
        self, columns: Sequence[int] | None = None, *, has_header: bool = False
    ) -> int:
        """功能：删除区域内键列值完全相同的重复行，并把保留行向上连续排列。

        使用方法：``removed = ws.range('A1:D100').remove_duplicates([0], has_header=True)``。
        参数：``columns`` 为相对于当前区域左侧的 0-based 列索引序列，省略时比较
        全部列；``has_header`` 为真时首行不参与去重。
        返回：实际删除的行数。
        异常：列索引、开关或包含合并单元格的区域无效时抛出 ``TypeError`` 或
        ``ValueError``。
        """
        selected = self._selected_columns(columns)
        rows = list(range(self.min_row + int(has_header), self.max_row + 1))
        seen: set[tuple[Any, ...]] = set()
        kept: list[int] = []
        for row in rows:
            key = tuple(self._worksheet._values.get(row, self.min_column + col) for col in selected)
            if key not in seen:
                seen.add(key)
                kept.append(row)
        self._rewrite_rows(kept, has_header=has_header)
        return len(rows) - len(kept)

    def remove_blank_rows(self) -> int:
        """功能：删除区域内完全没有值和公式的空白行并向上收紧数据。

        使用方法：``removed = ws.range('A1:F200').remove_blank_rows()``。
        参数：无。
        返回：实际删除的空白行数。
        异常：区域与合并单元格相交时抛出 ``ValueError``。
        """
        kept = [
            row for row in range(self.min_row, self.max_row + 1)
            if any(
                self._worksheet._values.get(row, column) is not None
                or (row, column) in self._worksheet._formulas
                for column in range(self.min_column, self.max_column + 1)
            )
        ]
        self._rewrite_rows(kept, has_header=False)
        return self.max_row - self.min_row + 1 - len(kept)

    def _selected_columns(self, columns: Sequence[int] | None) -> tuple[int, ...]:
        """功能：验证并返回区域去重所使用的相对列索引。

        使用方法：由 :meth:`remove_duplicates` 内部调用。
        参数：``columns`` 为相对于当前区域左侧的 0-based 列索引序列或 ``None``。
        返回：验证后的索引元组；``None`` 转换为当前区域全部列。
        异常：序列为空、类型错误、包含布尔值或索引越界时抛出异常。
        """
        width = self.max_column - self.min_column + 1
        if columns is None:
            return tuple(range(width))
        if isinstance(columns, (str, bytes)):
            raise TypeError("columns 必须是 0-based 列索引序列或 None")
        result = tuple(columns)
        if not result or any(isinstance(value, bool) or not isinstance(value, int) or not 0 <= value < width for value in result):
            raise ValueError("columns 必须是区域范围内的非空 0-based 列索引序列")
        return result

    def _rewrite_rows(self, source_rows: list[int], *, has_header: bool) -> None:
        """功能：按原始行快照重写区域行，用于去重和清除空白行。

        使用方法：由 :meth:`remove_duplicates` 和 :meth:`remove_blank_rows` 内部调用。
        参数：``source_rows`` 为按新顺序保留的绝对 0-based 行索引；``has_header``
        控制是否保留区域首行不参与重写。
        返回：``None``；值、公式、样式和附属对象同步移动，剩余位置清空。
        异常：当前区域与合并单元格相交时抛出 ``ValueError``。
        """
        if any(
            not (merged.max_row < self.min_row or merged.min_row > self.max_row
                 or merged.max_column < self.min_column or merged.min_column > self.max_column)
            for merged in self._worksheet.merged_ranges
        ):
            raise ValueError("包含合并单元格的区域不能删除行")
        start = self.min_row + int(has_header)
        maps = (
            self._worksheet._values._values, self._worksheet._formulas,
            self._worksheet._formula_values, self._worksheet._formula_errors,
            self._worksheet._hyperlinks, self._worksheet._notes, self._worksheet._styles,
        )
        # 分映射保存完整行快照，避免清空原区域后丢失待保留的值、公式和样式。
        row_snapshots = []
        for source_row in source_rows:
            row_snapshots.append([
                {
                    column: deepcopy(mapping[(source_row, column)])
                    for column in range(self.min_column, self.max_column + 1)
                    if (source_row, column) in mapping
                }
                for mapping in maps
            ])
        for mapping in maps:
            for row in range(start, self.max_row + 1):
                for column in range(self.min_column, self.max_column + 1):
                    mapping.pop((row, column), None)
        for target_row, snapshot in zip(range(start, self.max_row + 1), row_snapshots):
            for mapping, values in zip(maps, snapshot):
                for column, value in values.items():
                    mapping[(target_row, column)] = value
            self._worksheet._touch(target_row, self.max_column)
        self._worksheet._workbook._invalidate_formula_caches()


class RangeFormat:
    """表示一个区域的可批量修改格式属性。"""

    __slots__ = ("_range",)

    def __init__(self, area: Range) -> None:
        """功能：绑定由 ``Range.format`` 创建的格式代理。"""
        self._range = area

    @property
    def number(self) -> str | None:
        """功能：读取区域统一数字格式；不一致时返回 ``None``。

        使用方法：``number_format = worksheet.range(''C2:C20'').format.number``。
        参数：无，只读时不需要参数。
        返回：全部单元格一致时返回 Excel 数字格式字符串，否则返回 ``None``。
        """
        values = {
            self._range._worksheet._styles.get((row, column), DEFAULT_STYLE).number_format
            for row in range(self._range.min_row, self._range.max_row + 1)
            for column in range(self._range.min_column, self._range.max_column + 1)
        }
        return values.pop() if len(values) == 1 else None

    @number.setter
    def number(self, value: str) -> None:
        """功能：设置整个区域的 Excel 数字格式。

        使用方法：``worksheet.range(''C2:C20'').format.number = ''0.00''``。
        参数：``value`` 为非空 Excel 数字格式字符串。
        返回：``None``；区域中每个单元格保留其他样式字段。
        异常：值不是非空字符串时抛出 ``ValueError``。
        """
        if not isinstance(value, str) or not value:
            raise ValueError("number 必须是非空 Excel 数字格式字符串")
        for row in range(self._range.min_row, self._range.max_row + 1):
            for column in range(self._range.min_column, self._range.max_column + 1):
                coordinate = (row, column)
                style = self._range._worksheet._styles.get(coordinate, DEFAULT_STYLE)
                self._range._worksheet._styles[coordinate] = replace(style, number_format=value)
                self._range._worksheet._touch(row, column)
