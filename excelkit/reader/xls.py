"""基于 xlrd 的 Excel 97–2003 XLS 数据与样式读取实现。"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Dict, Optional, Type, TypeVar

import xlrd

from ..errors import InvalidFileError
from ..style import Alignment, Border, DEFAULT_STYLE, Fill, Font, Side, Style

if TYPE_CHECKING:
    from ..core.workbook import Workbook

_WorkbookType = TypeVar("_WorkbookType", bound="Workbook")
_BORDER_STYLES = {
    0: None, 1: "thin", 2: "medium", 3: "dashed", 4: "dotted",
    5: "thick", 6: "double", 7: "hair", 8: "mediumDashed",
    9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot",
    12: "mediumDashDotDot", 13: "slantDashDot",
}
_HORIZONTAL_ALIGNMENTS = {
    0: "general", 1: "left", 2: "center", 3: "right", 4: "fill",
    5: "justify", 6: "centerContinuous", 7: "distributed",
}
_VERTICAL_ALIGNMENTS = {
    0: "top", 1: "center", 2: "bottom", 3: "justify", 4: "distributed",
}


def _color(book: xlrd.book.Book, index: int) -> Optional[str]:
    """功能：把 XLS 调色板索引转换为 ExcelKit ARGB 颜色。

    使用方法：读取字体、填充和边框颜色时内部调用。
    参数：``book`` 为 xlrd 工作簿；``index`` 为 XLS 调色板索引。
    返回：8 位 ARGB 字符串；自动色或未知颜色返回 ``None``。
    """
    rgb = book.colour_map.get(index)
    if not rgb:
        return None
    return "FF{:02X}{:02X}{:02X}".format(*rgb)


def _side(book: xlrd.book.Book, line_style: int, color_index: int) -> Side:
    """功能：把 XLS 边框线记录转换为 ExcelKit Side。

    使用方法：构造四边 Border 时内部调用。
    参数：``book`` 为 xlrd 工作簿；``line_style`` 为线型编号；
    ``color_index`` 为调色板索引。
    返回：对应的 :class:`Side`；未知线型按无边框处理。
    """
    style = _BORDER_STYLES.get(line_style)
    return Side(style=style, color=_color(book, color_index) if style else None)


def _style(book: xlrd.book.Book, xf_index: int) -> Style:
    """功能：把 XLS XF 格式记录转换为 ExcelKit Style。

    使用方法：遍历单元格并遇到新 XF 索引时调用。
    参数：``book`` 为以 ``formatting_info=True`` 打开的 xlrd 工作簿；
    ``xf_index`` 为 0-based XF 索引。
    返回：字体、填充、边框、对齐和数字格式组合样式。
    异常：XF 索引损坏时抛出 :class:`InvalidFileError`。
    """
    try:
        xf = book.xf_list[xf_index]
        source_font = book.font_list[xf.font_index]
    except IndexError as error:
        raise InvalidFileError(f"无效的 XLS 样式索引：{xf_index}") from error
    font = Font(
        name=source_font.name or "Arial",
        size=max(source_font.height / 20.0, 1.0),
        bold=bool(source_font.bold),
        italic=bool(source_font.italic),
        underline=bool(source_font.underlined or source_font.underline_type),
        color=_color(book, source_font.colour_index),
    )
    background = xf.background
    fill = Fill(
        _color(book, background.pattern_colour_index)
        if background.fill_pattern
        else None
    )
    source_border = xf.border
    border = Border(
        left=_side(book, source_border.left_line_style, source_border.left_colour_index),
        right=_side(book, source_border.right_line_style, source_border.right_colour_index),
        top=_side(book, source_border.top_line_style, source_border.top_colour_index),
        bottom=_side(book, source_border.bottom_line_style, source_border.bottom_colour_index),
    )
    source_alignment = xf.alignment
    alignment = Alignment(
        horizontal=_HORIZONTAL_ALIGNMENTS.get(source_alignment.hor_align),
        vertical=_VERTICAL_ALIGNMENTS.get(source_alignment.vert_align),
        wrap_text=bool(source_alignment.text_wrapped),
    )
    source_format = book.format_map.get(xf.format_key)
    number_format = source_format.format_str if source_format else "General"
    return Style(font, fill, border, alignment, number_format or "General")


def _value(book: xlrd.book.Book, cell: xlrd.sheet.Cell) -> Any:
    """功能：把 xlrd 单元格值转换为 ExcelKit 使用的 Python 类型。

    使用方法：XLS 工作表遍历时为每个单元格调用。
    参数：``book`` 为 xlrd 工作簿；``cell`` 为 xlrd 单元格对象。
    返回：``None``、字符串、布尔值、整数、浮点数、日期时间或错误文本。
    """
    if cell.ctype in {xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK}:
        return None
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        number = float(cell.value)
        return int(number) if number.is_integer() else number
    if cell.ctype == xlrd.XL_CELL_DATE:
        result = xlrd.xldate_as_datetime(cell.value, book.datemode)
        return result.date() if result.time() == datetime.min.time() else result
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(cell.value, f"#ERROR({cell.value})")
    return cell.value


def _load_xls(
    workbook_class: Type[_WorkbookType], filename: os.PathLike | str
) -> _WorkbookType:
    """功能：读取二进制 XLS 文件的数据和受支持样式并构造 Workbook。

    使用方法：仅由 ``Workbook.load('input.xls')`` 的格式分派器调用。
    参数：``workbook_class`` 为 Workbook 类或子类；``filename`` 为源文件路径。
    返回：包含全部工作表、计算结果、日期和基础样式的新工作簿。
    异常：文件不存在时抛出 ``FileNotFoundError``；文件损坏、加密或不是 XLS 时
    抛出 :class:`InvalidFileError`。
    """
    path = Path(filename)
    if not path.exists():
        raise FileNotFoundError(path)
    try:
        source = xlrd.open_workbook(str(path), formatting_info=True)
    except (xlrd.XLRDError, OSError) as error:
        raise InvalidFileError(f"无法读取 XLS 文件：{path}") from error
    workbook = workbook_class()
    style_cache: Dict[int, Style] = {}
    for source_sheet in source.sheets():
        worksheet = workbook.add_sheet(source_sheet.name)
        for row in range(source_sheet.nrows):
            for column in range(source_sheet.ncols):
                cell = source_sheet.cell(row, column)
                xf_index = source_sheet.cell_xf_index(row, column)
                if xf_index not in style_cache:
                    style_cache[xf_index] = _style(source, xf_index)
                style = style_cache[xf_index]
                value = _value(source, cell)
                if value is not None or style != DEFAULT_STYLE:
                    worksheet._touch(row, column)
                if value is not None:
                    worksheet._values.set(row, column, value)
                if style != DEFAULT_STYLE:
                    worksheet._styles[(row, column)] = style
    return workbook


__all__ = []
