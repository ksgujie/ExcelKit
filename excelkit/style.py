"""单元格常用字体、填充、边框、对齐和数字格式类型。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

_COLOR_PATTERN = re.compile(r"^[0-9A-Fa-f]{6}(?:[0-9A-Fa-f]{2})?$")
_BORDER_STYLES = {
    None,
    "thin",
    "medium",
    "thick",
    "dashed",
    "dotted",
    "double",
    "hair",
    "dashDot",
    "dashDotDot",
    "mediumDashed",
    "mediumDashDot",
    "mediumDashDotDot",
    "slantDashDot",
}
_HORIZONTAL_ALIGNMENTS = {
    None,
    "general",
    "left",
    "center",
    "right",
    "fill",
    "justify",
    "centerContinuous",
    "distributed",
}
_VERTICAL_ALIGNMENTS = {None, "top", "center", "bottom", "justify", "distributed"}


def _color(value: Optional[str]) -> Optional[str]:
    """功能：验证并规范化 RGB 或 ARGB 颜色字符串。

    使用方法：由字体、填充和边框类型初始化时内部调用。
    参数：``value`` 为 ``RRGGBB``、``AARRGGBB`` 字符串或 ``None``。
    返回：统一大写的 8 位 ARGB 字符串；6 位 RGB 自动补 ``FF`` 不透明度。
    异常：类型、长度或十六进制字符无效时抛出 ``ValueError``。
    """
    if value is None:
        return None
    if not isinstance(value, str) or _COLOR_PATTERN.fullmatch(value) is None:
        raise ValueError("颜色必须是 6 位 RGB 或 8 位 ARGB 十六进制字符串")
    normalized = value.upper()
    return f"FF{normalized}" if len(normalized) == 6 else normalized


@dataclass(frozen=True)
class Font:
    """字体样式；支持名称、字号、粗体、斜体、下划线和 RGB/ARGB 颜色。"""

    name: str = "Calibri"
    size: float = 11.0
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None

    def __post_init__(self) -> None:
        """功能：验证字体字段并规范化颜色。

        使用方法：创建 ``Font(...)`` 时由 dataclass 自动调用。
        参数：字段来自当前实例；名称必须非空，字号必须为正数，三个开关必须为布尔值。
        返回：``None``；颜色会原地规范化为不可变 ARGB 字符串。
        异常：字段类型或取值无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(self.name, str) or not self.name:
            raise ValueError("字体名称必须是非空字符串")
        if isinstance(self.size, bool) or not isinstance(self.size, (int, float)):
            raise TypeError("字体大小必须是正数")
        if self.size <= 0:
            raise ValueError("字体大小必须大于 0")
        for field_name in ("bold", "italic", "underline"):
            if not isinstance(getattr(self, field_name), bool):
                raise TypeError(f"字体字段 {field_name} 必须是布尔值")
        object.__setattr__(self, "size", float(self.size))
        object.__setattr__(self, "color", _color(self.color))


@dataclass(frozen=True)
class Fill:
    """单色实心填充；颜色为 ``None`` 时表示无填充。"""

    color: Optional[str] = None

    def __post_init__(self) -> None:
        """功能：验证并规范化填充颜色。

        使用方法：创建 ``Fill(color=...)`` 时自动调用。
        参数：``color`` 为 6 位 RGB、8 位 ARGB 或 ``None``。
        返回：``None``；非空颜色规范化为大写 ARGB。
        异常：颜色无效时抛出 ``ValueError``。
        """
        object.__setattr__(self, "color", _color(self.color))


@dataclass(frozen=True)
class Side:
    """边框的一条边；包含线型和可选颜色。"""

    style: Optional[str] = None
    color: Optional[str] = None

    def __post_init__(self) -> None:
        """功能：验证边框线型并规范化颜色。

        使用方法：创建 ``Side(style="thin", color="000000")`` 时自动调用。
        参数：``style`` 为受支持 Excel 线型或 ``None``；``color`` 为颜色或
        ``None``。
        返回：``None``。
        异常：线型或颜色无效时抛出 ``ValueError``。
        """
        if self.style not in _BORDER_STYLES:
            raise ValueError(f"不支持的边框线型：{self.style!r}")
        object.__setattr__(self, "color", _color(self.color))


@dataclass(frozen=True)
class Border:
    """由左、右、上、下四条不可变 Side 组成的单元格边框。"""

    left: Side = field(default_factory=Side)
    right: Side = field(default_factory=Side)
    top: Side = field(default_factory=Side)
    bottom: Side = field(default_factory=Side)

    def __post_init__(self) -> None:
        """功能：验证四个边框字段均为 Side。

        使用方法：创建 ``Border(...)`` 时自动调用。
        参数：``left``、``right``、``top``、``bottom`` 为 :class:`Side`。
        返回：``None``。
        异常：任一字段类型错误时抛出 ``TypeError``。
        """
        for field_name in ("left", "right", "top", "bottom"):
            if not isinstance(getattr(self, field_name), Side):
                raise TypeError(f"边框字段 {field_name} 必须是 Side")


@dataclass(frozen=True)
class Alignment:
    """水平、垂直对齐及自动换行设置。"""

    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    wrap_text: bool = False

    def __post_init__(self) -> None:
        """功能：验证对齐方式和自动换行开关。

        使用方法：创建 ``Alignment(horizontal="center")`` 时自动调用。
        参数：``horizontal``、``vertical`` 为受支持 Excel 对齐值或 ``None``；
        ``wrap_text`` 为布尔值。
        返回：``None``。
        异常：值无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if self.horizontal not in _HORIZONTAL_ALIGNMENTS:
            raise ValueError(f"不支持的水平对齐方式：{self.horizontal!r}")
        if self.vertical not in _VERTICAL_ALIGNMENTS:
            raise ValueError(f"不支持的垂直对齐方式：{self.vertical!r}")
        if not isinstance(self.wrap_text, bool):
            raise TypeError("wrap_text 必须是布尔值")


@dataclass(frozen=True)
class Style:
    """组合字体、填充、边框、对齐和数字格式的完整单元格样式。"""

    font: Font = field(default_factory=Font)
    fill: Fill = field(default_factory=Fill)
    border: Border = field(default_factory=Border)
    alignment: Alignment = field(default_factory=Alignment)
    number_format: str = "General"

    def __post_init__(self) -> None:
        """功能：验证组合样式字段类型和数字格式。

        使用方法：创建 ``Style(...)`` 时自动调用。
        参数：``font``、``fill``、``border``、``alignment`` 为对应样式类型；
        ``number_format`` 为非空 Excel 数字格式字符串。
        返回：``None``。
        异常：字段类型或数字格式无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        expected = {
            "font": Font,
            "fill": Fill,
            "border": Border,
            "alignment": Alignment,
        }
        for field_name, field_type in expected.items():
            if not isinstance(getattr(self, field_name), field_type):
                raise TypeError(f"样式字段 {field_name} 必须是 {field_type.__name__}")
        if not isinstance(self.number_format, str) or not self.number_format:
            raise ValueError("number_format 必须是非空字符串")


DEFAULT_STYLE = Style()

__all__ = ["Font", "Fill", "Side", "Border", "Alignment", "Style"]
