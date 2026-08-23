"""连续矩形区域的普通值批量访问对象。"""

from __future__ import annotations

from collections.abc import Iterable
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
