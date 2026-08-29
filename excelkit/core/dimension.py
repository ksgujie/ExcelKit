"""工作表0-based行高、列宽和隐藏状态对象。"""

from __future__ import annotations

from typing import Optional

from ..address import validate_column_index, validate_row_index


def _size(value: Optional[float], name: str, maximum: float) -> Optional[float]:
    """功能：验证并规范化可恢复默认值的行高或列宽。

    使用方法：行高、列宽属性设置器内部调用。
    参数：``value`` 为正数或 ``None``；``name`` 为错误信息字段名称；
    ``maximum`` 为 Excel 格式允许的上限。
    返回：规范化后的 ``float`` 或 ``None``。
    异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
    """
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError(f"{name} 必须是数字或 None")
    converted = float(value)
    if not 0 < converted <= maximum:
        raise ValueError(f"{name} 必须大于 0 且不超过 {maximum:g}")
    return converted


def _outline_level(value: int) -> int:
    """功能：验证 Excel 行列分组的大纲层级。

    使用方法：由行列尺寸对象的 ``outline_level`` 属性设置器调用。
    参数：``value`` 为 0～7 的整数；0 表示未分组。
    返回：验证后的整数层级。
    异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
    """
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError("outline_level 必须是整数")
    if not 0 <= value <= 7:
        raise ValueError("outline_level 必须在 0～7 之间")
    return value


class RowDimension:
    """表示一行的0-based索引、尺寸、隐藏状态和大纲分组状态。"""

    __slots__ = ("_index", "_height", "_hidden", "_outline_level", "_collapsed")

    def __init__(
        self, index: int, height: Optional[float] = None, hidden: bool = False
    ) -> None:
        """功能：创建并验证行尺寸对象。

        使用方法：通常通过 ``worksheet.row(index)`` 获取，不直接实例化。
        参数：``index`` 为0-based行索引；``height`` 为磅值或 ``None``；
        ``hidden`` 为隐藏开关。
        返回：无。
        """
        self._index = validate_row_index(index)
        self._height: Optional[float] = None
        self._hidden = False
        self._outline_level = 0
        self._collapsed = False
        self.height = height
        self.hidden = hidden

    @property
    def index(self) -> int:
        """功能：读取不可修改的0-based行索引。

        使用方法：``index = worksheet.row(0).index``。
        参数：无。
        返回：有效0-based整数行索引。
        """
        return self._index

    @property
    def height(self) -> Optional[float]:
        """功能：读取自定义行高。

        使用方法：``height = worksheet.row(0).height``。
        参数：无。
        返回：以磅为单位的 ``float``；使用默认行高时为 ``None``。
        """
        return self._height

    @height.setter
    def height(self, value: Optional[float]) -> None:
        """功能：设置行高或恢复默认行高。

        使用方法：``worksheet.row(0).height = 28``；赋值 ``None`` 恢复默认。
        参数：``value`` 为大于0且不超过409的磅值或 ``None``。
        返回：``None``。
        异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        self._height = _size(value, "行高", 409)

    @property
    def hidden(self) -> bool:
        """功能：读取当前行是否隐藏。

        使用方法：``hidden = worksheet.row(5).hidden``。
        参数：无。
        返回：布尔值。
        """
        return self._hidden

    @hidden.setter
    def hidden(self, value: bool) -> None:
        """功能：隐藏或显示当前行。

        使用方法：``worksheet.row(5).hidden = True``。
        参数：``value`` 必须是布尔值。
        返回：``None``。
        异常：类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("RowDimension.hidden 必须是布尔值")
        self._hidden = value

    @property
    def outline_level(self) -> int:
        """功能：读取当前行的 Excel 大纲层级。

        使用方法：``level = worksheet.row(3).outline_level``。
        参数：无。
        返回：0～7 的整数；0 表示当前行没有分组。
        """
        return self._outline_level

    @outline_level.setter
    def outline_level(self, value: int) -> None:
        """功能：设置当前行的 Excel 大纲层级。

        使用方法：通常由 ``worksheet.group_rows()`` 调用，也可直接设置。
        参数：``value`` 为 0～7 的整数。
        返回：``None``。
        异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        self._outline_level = _outline_level(value)

    @property
    def collapsed(self) -> bool:
        """功能：读取当前行是否带有 Excel 大纲折叠标志。

        使用方法：``collapsed = worksheet.row(3).collapsed``。
        参数：无。
        返回：布尔值。
        """
        return self._collapsed

    @collapsed.setter
    def collapsed(self, value: bool) -> None:
        """功能：设置当前行的大纲折叠标志。

        使用方法：通常由 ``worksheet.group_rows(..., collapsed=True)`` 调用。
        参数：``value`` 必须是布尔值。
        返回：``None``。
        异常：类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("RowDimension.collapsed 必须是布尔值")
        self._collapsed = value

    def _is_default(self) -> bool:
        """功能：判断当前行是否没有任何自定义尺寸设置。

        使用方法：写出器筛选需要序列化的行时内部调用。
        参数：无。
        返回：行高为默认且未隐藏时返回 ``True``。
        """
        return (
            self._height is None and not self._hidden
            and self._outline_level == 0 and not self._collapsed
        )


class ColumnDimension:
    """表示一列的0-based索引、尺寸、隐藏状态和大纲分组状态。"""

    __slots__ = ("_index", "_width", "_hidden", "_outline_level", "_collapsed")

    def __init__(
        self, index: int, width: Optional[float] = None, hidden: bool = False
    ) -> None:
        """功能：创建并验证列尺寸对象。

        使用方法：通常通过 ``worksheet.column(index)`` 获取，不直接实例化。
        参数：``index`` 为0-based列索引；``width`` 为Excel字符宽度或 ``None``；
        ``hidden`` 为隐藏开关。
        返回：无。
        """
        self._index = validate_column_index(index)
        self._width: Optional[float] = None
        self._hidden = False
        self._outline_level = 0
        self._collapsed = False
        self.width = width
        self.hidden = hidden

    @property
    def index(self) -> int:
        """功能：读取不可修改的0-based列索引。

        使用方法：``index = worksheet.column(0).index``。
        参数：无。
        返回：有效0-based整数列索引。
        """
        return self._index

    @property
    def width(self) -> Optional[float]:
        """功能：读取自定义Excel列宽。

        使用方法：``width = worksheet.column(0).width``。
        参数：无。
        返回：Excel字符宽度 ``float``；使用默认列宽时为 ``None``。
        """
        return self._width

    @width.setter
    def width(self, value: Optional[float]) -> None:
        """功能：设置列宽或恢复默认列宽。

        使用方法：``worksheet.column(0).width = 20``；赋值 ``None`` 恢复默认。
        参数：``value`` 为大于0且不超过255的数字或 ``None``。
        返回：``None``。
        异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        self._width = _size(value, "列宽", 255)

    @property
    def hidden(self) -> bool:
        """功能：读取当前列是否隐藏。

        使用方法：``hidden = worksheet.column(2).hidden``。
        参数：无。
        返回：布尔值。
        """
        return self._hidden

    @hidden.setter
    def hidden(self, value: bool) -> None:
        """功能：隐藏或显示当前列。

        使用方法：``worksheet.column(2).hidden = True``。
        参数：``value`` 必须是布尔值。
        返回：``None``。
        异常：类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("ColumnDimension.hidden 必须是布尔值")
        self._hidden = value

    @property
    def outline_level(self) -> int:
        """功能：读取当前列的 Excel 大纲层级。

        使用方法：``level = worksheet.column(2).outline_level``。
        参数：无。
        返回：0～7 的整数；0 表示当前列没有分组。
        """
        return self._outline_level

    @outline_level.setter
    def outline_level(self, value: int) -> None:
        """功能：设置当前列的 Excel 大纲层级。

        使用方法：通常由 ``worksheet.group_columns()`` 调用，也可直接设置。
        参数：``value`` 为 0～7 的整数。
        返回：``None``。
        异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        self._outline_level = _outline_level(value)

    @property
    def collapsed(self) -> bool:
        """功能：读取当前列是否带有 Excel 大纲折叠标志。

        使用方法：``collapsed = worksheet.column(2).collapsed``。
        参数：无。
        返回：布尔值。
        """
        return self._collapsed

    @collapsed.setter
    def collapsed(self, value: bool) -> None:
        """功能：设置当前列的大纲折叠标志。

        使用方法：通常由 ``worksheet.group_columns(..., collapsed=True)`` 调用。
        参数：``value`` 必须是布尔值。
        返回：``None``。
        异常：类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("ColumnDimension.collapsed 必须是布尔值")
        self._collapsed = value

    def _is_default(self) -> bool:
        """功能：判断当前列是否没有任何自定义尺寸设置。

        使用方法：写出器筛选需要序列化的列时内部调用。
        参数：无。
        返回：列宽为默认且未隐藏时返回 ``True``。
        """
        return (
            self._width is None and not self._hidden
            and self._outline_level == 0 and not self._collapsed
        )


__all__ = ["RowDimension", "ColumnDimension"]
