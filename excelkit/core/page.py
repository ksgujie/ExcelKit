"""工作表页面布局、打印边距及页眉页脚对象。"""

from __future__ import annotations

import math
from copy import deepcopy
from dataclasses import dataclass
from typing import ClassVar, Optional, Tuple, cast

from ..address import range_index, validate_column_index, validate_row_index

_PAPER_SIZES = {
    "a3": "A3",
    "a4": "A4",
    "a5": "A5",
    "letter": "Letter",
    "legal": "Legal",
}
_ORIENTATIONS = {"portrait", "landscape"}
_ORDERS = {"down_then_over", "over_then_down"}


def _positive_number(value: float, name: str, allow_zero: bool = False) -> float:
    """功能：验证打印边距等有限非负或正数。

    使用方法：页面值对象初始化时内部调用。
    参数：``value`` 为待验证数字；``name`` 为字段名称；``allow_zero`` 控制是否
    接受0。
    返回：规范化后的 ``float``。
    异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError(f"{name} 必须是数字")
    converted = float(value)
    if not math.isfinite(converted) or (
        converted < 0 if allow_zero else converted <= 0
    ):
        comparison = "大于等于0" if allow_zero else "大于0"
        raise ValueError(f"{name} 必须{comparison}")
    return converted


@dataclass(frozen=True)
class PageMargins:
    """以厘米为单位保存左、右、上、下、页眉和页脚边距。"""

    left: float = 1.78
    right: float = 1.78
    top: float = 1.91
    bottom: float = 1.91
    header: float = 0.76
    footer: float = 0.76

    def __post_init__(self) -> None:
        """功能：验证并把六个厘米边距规范化为浮点数。

        使用方法：创建 ``PageMargins(...)`` 时自动调用。
        参数：所有字段均以厘米为单位，允许0但不允许负数或布尔值。
        返回：``None``；字段在不可变实例中规范化。
        异常：类型或范围无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        for name in ("left", "right", "top", "bottom", "header", "footer"):
            object.__setattr__(
                self, name, _positive_number(getattr(self, name), name, True)
            )


@dataclass(frozen=True)
class HeaderFooter:
    """保存页眉或页脚的左、中、右三个文本区域。"""

    left: str = ""
    center: str = ""
    right: str = ""

    def __post_init__(self) -> None:
        """功能：验证页眉页脚三个区域均为字符串。

        使用方法：创建 ``HeaderFooter(...)`` 时自动调用。
        参数：``left``、``center``、``right`` 可包含Excel页码等控制代码。
        返回：``None``。
        异常：任一字段不是字符串时抛出 ``TypeError``。
        """
        for name in ("left", "center", "right"):
            if not isinstance(getattr(self, name), str):
                raise TypeError(f"HeaderFooter.{name} 必须是字符串")


def header_footer_text(value: HeaderFooter) -> str:
    """功能：把页眉页脚三个区域组合为Excel通用控制代码字符串。

    使用方法：XLSX和XLS写出器共用，避免两种格式重复实现区域拼接。
    参数：``value`` 为 :class:`HeaderFooter`。
    返回：按 ``&L``、``&C``、``&R`` 分段的字符串；全部为空时返回空字符串。
    """
    parts = []
    if value.left:
        parts.append(f"&L{value.left}")
    if value.center:
        parts.append(f"&C{value.center}")
    if value.right:
        parts.append(f"&R{value.right}")
    return "".join(parts)


def _index_pair(
    value: Optional[Tuple[int, int]], name: str, *, rows: bool
) -> Optional[Tuple[int, int]]:
    """功能：验证重复打印行或列的0-based包含式起止索引。

    使用方法：``PageSettings`` 设置重复标题时内部调用。
    参数：``value`` 为两个整数的元组或 ``None``；``name`` 为字段名称；
    ``rows`` 指示校验行边界还是列边界。
    返回：验证后的元组或 ``None``。
    异常：结构、索引类型、范围或顺序无效时抛出 ``TypeError`` 或 ``ValueError``。
    """
    if value is None:
        return None
    if not isinstance(value, tuple) or len(value) != 2:
        raise TypeError(f"{name} 必须是包含两个0-based索引的元组或 None")
    start, end = value
    validator = validate_row_index if rows else validate_column_index
    validator(start)
    validator(end)
    if start > end:
        raise ValueError(f"{name} 的起始索引不能大于结束索引")
    return start, end


class PageSettings:
    """集中保存工作表页面、缩放、打印区域及页眉页脚设置。"""

    PORTRAIT: ClassVar[str] = "portrait"
    LANDSCAPE: ClassVar[str] = "landscape"
    A3: ClassVar[str] = "A3"
    A4: ClassVar[str] = "A4"
    A5: ClassVar[str] = "A5"
    LETTER: ClassVar[str] = "Letter"
    LEGAL: ClassVar[str] = "Legal"
    DOWN_THEN_OVER: ClassVar[str] = "down_then_over"
    OVER_THEN_DOWN: ClassVar[str] = "over_then_down"

    __slots__ = (
        "orientation", "paper_size", "scale", "_fit_width", "_fit_height",
        "first_page_number", "black_and_white", "draft", "print_order", "margins",
        "print_area", "repeat_rows", "repeat_columns", "center_horizontal",
        "center_vertical", "print_gridlines", "print_headings", "header", "footer",
    )

    def __init__(self) -> None:
        """功能：创建采用常用A4纵向打印默认值的页面设置。

        使用方法：由每张 ``Worksheet`` 自动创建，通过 ``worksheet.page`` 获取。
        参数：无。
        返回：无。
        """
        object.__setattr__(self, "orientation", "portrait")
        object.__setattr__(self, "paper_size", "A4")
        object.__setattr__(self, "scale", 100)
        object.__setattr__(self, "_fit_width", None)
        object.__setattr__(self, "_fit_height", None)
        object.__setattr__(self, "first_page_number", None)
        object.__setattr__(self, "black_and_white", False)
        object.__setattr__(self, "draft", False)
        object.__setattr__(self, "print_order", self.DOWN_THEN_OVER)
        object.__setattr__(self, "margins", PageMargins())
        object.__setattr__(self, "print_area", None)
        object.__setattr__(self, "repeat_rows", None)
        object.__setattr__(self, "repeat_columns", None)
        object.__setattr__(self, "center_horizontal", False)
        object.__setattr__(self, "center_vertical", False)
        object.__setattr__(self, "print_gridlines", False)
        object.__setattr__(self, "print_headings", False)
        object.__setattr__(self, "header", HeaderFooter())
        object.__setattr__(self, "footer", HeaderFooter())

    def __setattr__(self, name: str, value: object) -> None:
        """功能：集中验证页面属性并维护缩放与适应页数的互斥关系。

        使用方法：用户执行 ``worksheet.page.orientation = 'landscape'`` 等赋值时
        自动调用。
        参数：``name`` 为固定页面字段；``value`` 为待设置值。
        返回：``None``。
        异常：字段类型、范围或枚举值无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if name == "orientation":
            if value not in _ORIENTATIONS:
                raise ValueError("orientation 必须是 'portrait' 或 'landscape'")
        elif name == "paper_size":
            if not isinstance(value, str) or value.casefold() not in _PAPER_SIZES:
                raise ValueError("paper_size 必须是 A3、A4、A5、Letter 或 Legal")
            value = _PAPER_SIZES[value.casefold()]
        elif name == "scale":
            if (
                value is None
                or isinstance(value, bool)
                or not isinstance(value, int)
                or not 10 <= value <= 400
            ):
                raise ValueError("scale 必须是10～400的整数；适应页数请使用 fit()")
            object.__setattr__(self, "_fit_width", None)
            object.__setattr__(self, "_fit_height", None)
        elif name == "first_page_number":
            if value is not None and (
                isinstance(value, bool) or not isinstance(value, int) or value <= 0
            ):
                raise ValueError("first_page_number 必须是正整数或 None")
        elif name in {
            "black_and_white", "draft", "center_horizontal", "center_vertical",
            "print_gridlines", "print_headings",
        }:
            if not isinstance(value, bool):
                raise TypeError(f"{name} 必须是布尔值")
        elif name == "print_order":
            if value not in _ORDERS:
                raise ValueError(
                    "print_order 必须是 'down_then_over' 或 'over_then_down'"
                )
        elif name == "margins":
            if not isinstance(value, PageMargins):
                raise TypeError("margins 必须是 PageMargins")
        elif name in {"header", "footer"}:
            if not isinstance(value, HeaderFooter):
                raise TypeError(f"{name} 必须是 HeaderFooter")
        elif name == "print_area":
            if value is not None:
                if not isinstance(value, str):
                    raise TypeError("print_area 必须是A1区域字符串或 None")
                range_index(value)
                value = value.upper()
        elif name == "repeat_rows":
            value = _index_pair(
                cast(Optional[Tuple[int, int]], value), name, rows=True
            )
        elif name == "repeat_columns":
            value = _index_pair(
                cast(Optional[Tuple[int, int]], value), name, rows=False
            )
        object.__setattr__(self, name, value)

    def fit(
        self, width: Optional[int] = 1, height: Optional[int] = None
    ) -> "PageSettings":
        """功能：一次性设置打印内容适应的宽度和高度页数。

        使用方法：``page.fit(width=1)`` 表示一页宽且高度不限；
        ``page.fit(width=1, height=1)`` 表示整体适应一页。
        参数：``width``、``height`` 为正整数页数或 ``None``，至少一项非空。
        返回：当前 :class:`PageSettings`，支持继续设置其他页面属性。
        异常：类型、范围无效或两项均为 ``None`` 时抛出 ``ValueError``。
        """
        for name, value in (("width", width), ("height", height)):
            if value is not None and (
                isinstance(value, bool) or not isinstance(value, int) or value <= 0
            ):
                raise ValueError(f"{name} 必须是正整数或 None")
        if width is None and height is None:
            raise ValueError("fit() 的 width 和 height 不能同时为 None")
        object.__setattr__(self, "scale", None)
        object.__setattr__(self, "_fit_width", width)
        object.__setattr__(self, "_fit_height", height)
        return self

    def __deepcopy__(self, memo: dict[int, object]) -> "PageSettings":
        """功能：完整复制页面公开配置和内部缩放模式状态。

        使用方法：由 ``Workbook.copy_sheet()`` 通过 ``deepcopy(page)`` 自动调用。
        参数：``memo`` 为 Python 深复制协议维护的对象映射。
        返回：与当前对象状态相同、后续可独立修改的新 :class:`PageSettings`。
        异常：嵌套值无法深复制时透传对应异常；不会修改原页面设置。
        """
        copied = object.__new__(type(self))
        memo[id(self)] = copied
        for name in self.__slots__:
            object.__setattr__(copied, name, deepcopy(getattr(self, name), memo))
        return copied

    def __repr__(self) -> str:
        """功能：生成用于调试的页面设置文本表示。

        使用方法：``repr(worksheet.page)``。
        参数：无。
        返回：包含纸张、方向及缩放模式的字符串。
        """
        mode = f"scale={self.scale}" if self.scale is not None else (
            f"fit=({self._fit_width}, {self._fit_height})"
        )
        return f"<PageSettings {self.paper_size} {self.orientation} {mode}>"


__all__ = ["PageSettings", "PageMargins", "HeaderFooter"]
