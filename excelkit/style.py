"""单元格样式系统：Font、Fill、Border、Alignment、Style 均为不可改对象，可安全复用。

XLSX 支持样式完整往返；XLS 仅映射基础样式。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Font:
    """字体样式。

    参数:
        name: 字体名称，如 "微软雅黑"、"Calibri"。
        size: 字号，如 11、12。
        bold: 是否加粗。
        italic: 是否斜体。
        color: ARGB 颜色字符串，如 "FF000000" 或 6 位 RGB "000000"。
        underline: 是否下划线。
    """

    name: str | None = None
    size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None
    underline: bool | None = None


@dataclass(frozen=True)
class Fill:
    """填充样式。

    参数:
        fg_color: 前景色 ARGB 或 6 位 RGB 字符串，如 "4472C4"。
        pattern_type: 填充模式，如 "solid"。
    """

    fg_color: str | None = None
    pattern_type: str | None = "solid"


@dataclass(frozen=True)
class BorderSide:
    """单边框样式。

    参数:
        style: 边框样式，如 "thin"、"medium"、"dashed"、"dotted"、"double"。
        color: 边框颜色，如 "000000"。
    """

    style: str | None = None
    color: str | None = None


class Border:
    """边框，THIN 等常量供 IDE 自动补全。"""

    THIN = "thin"
    MEDIUM = "medium"
    DASHED = "dashed"
    DOTTED = "dotted"
    DOUBLE = "double"
    HAIR = "hair"

    def __init__(
        self,
        top: BorderSide | None = None,
        bottom: BorderSide | None = None,
        left: BorderSide | None = None,
        right: BorderSide | None = None,
    ) -> None:
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right


@dataclass(frozen=True)
class Alignment:
    """对齐方式。

    参数:
        horizontal: 水平对齐，如 "left"、"center"、"right"、"fill"、"justify"。
        vertical: 垂直对齐，如 "top"、"center"、"bottom"、"justify"。
        wrap_text: 是否自动换行。
        indent: 缩进。
    """

    HORIZONTAL_LEFT = "left"
    HORIZONTAL_CENTER = "center"
    HORIZONTAL_RIGHT = "right"
    HORIZONTAL_FILL = "fill"
    HORIZONTAL_JUSTIFY = "justify"

    VERTICAL_TOP = "top"
    VERTICAL_CENTER = "center"
    VERTICAL_BOTTOM = "bottom"
    VERTICAL_JUSTIFY = "justify"

    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool | None = None
    indent: int | None = 0


@dataclass(frozen=True)
class Style:
    """完整单元格样式，不可变，可安全复用于多个单元格。"""

    font: Font | None = None
    fill: Fill | None = None
    border: Border | None = None
    alignment: Alignment | None = None
    number_format: str | None = None

    def copy(self) -> "Style":
        """返回当前样式的副本（由于不可变，直接返回自身）。"""
        return self


@dataclass(frozen=True)
class NumberFormat:
    """常用数字格式预设。"""

    GENERAL = "General"
    NUMBER = "0"
    NUMBER_2DEC = "0.00"
    CURRENCY = '"$"#,##0.00'
    PERCENT = "0%"
    PERCENT_2DEC = "0.00%"
    DATE = "yyyy-m-d"
    DATETIME = "yyyy-m-d h:mm:ss"
    TEXT = "@"


@dataclass(frozen=True)
class ReportStyle:
    """报表常用样式预设。"""

    HEADER = Style(
        font=Font(bold=True, color="FFFFFF"),
        fill=Fill("4472C4"),
        alignment=Alignment(
            horizontal=Alignment.HORIZONTAL_CENTER,
            vertical=Alignment.VERTICAL_CENTER,
            wrap_text=False,
        ),
    )


__all__ = [
    "Font",
    "Fill",
    "BorderSide",
    "Border",
    "Alignment",
    "Style",
    "NumberFormat",
    "ReportStyle",
]
