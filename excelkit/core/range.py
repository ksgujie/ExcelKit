"""连续矩形区域的普通值批量访问对象。"""

from __future__ import annotations

from collections.abc import Iterable
from copy import deepcopy
from typing import TYPE_CHECKING, Any, List

from ..address import cell_address

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
        """功能：读取区域内的全部普通值。

        使用方法：``values = worksheet.range("A1:C2").values``。
        参数：无。
        返回：按行组织的二维 ``list``；空单元格和公式单元格返回 ``None``。
        """
        return [
            [
                self._worksheet._values.get(row, column)
                for column in range(self._min_column, self._max_column + 1)
            ]
            for row in range(self._min_row, self._max_row + 1)
        ]

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

    def clear_values(self) -> "Range":
        """功能：清除区域内普通值、公式、缓存结果和计算错误并保留样式。

        使用方法：``worksheet.range("A1:C10").clear_values()``。
        参数：无。
        返回：当前 :class:`Range`，支持链式调用。
        """
        self._worksheet._workbook._invalidate_formula_caches()
        for row in range(self._min_row, self._max_row + 1):
            for column in range(self._min_column, self._max_column + 1):
                coordinate = (row, column)
                self._worksheet._values.set(row, column, None)
                self._worksheet._formulas.pop(coordinate, None)
                self._worksheet._formula_values.pop(coordinate, None)
                self._worksheet._formula_errors.pop(coordinate, None)
        return self

    def clear_styles(self) -> "Range":
        """功能：把区域内全部单元格恢复为默认样式并保留值和公式。

        使用方法：``worksheet.range("A1:C10").clear_styles()``。
        参数：无。
        返回：当前 :class:`Range`，支持链式调用。
        """
        for row in range(self._min_row, self._max_row + 1):
            for column in range(self._min_column, self._max_column + 1):
                self._worksheet._styles.pop((row, column), None)
        return self

    def clear(self) -> "Range":
        """功能：同时清除区域内值、公式、缓存结果、计算错误和样式。

        使用方法：``worksheet.range("A1:C10").clear()``。
        参数：无；合并关系和行列尺寸不受影响。
        返回：当前 :class:`Range`，支持链式调用。
        """
        self.clear_values()
        self.clear_styles()
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
