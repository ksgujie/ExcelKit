"""基于 xlwt 的 Excel 97–2003 XLS 写出实现。"""

from __future__ import annotations

import os
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Dict, Optional, Tuple, Union

import xlwt

from ..core.page import header_footer_text
from ..style import DEFAULT_STYLE, Style

if TYPE_CHECKING:
    from ..core.workbook import Workbook

_BORDER_STYLES = {
    None: xlwt.Borders.NO_LINE,
    "thin": xlwt.Borders.THIN,
    "medium": xlwt.Borders.MEDIUM,
    "dashed": xlwt.Borders.DASHED,
    "dotted": xlwt.Borders.DOTTED,
    "thick": xlwt.Borders.THICK,
    "double": xlwt.Borders.DOUBLE,
    "hair": xlwt.Borders.HAIR,
    "mediumDashed": xlwt.Borders.MEDIUM_DASHED,
    "dashDot": xlwt.Borders.THIN_DASH_DOTTED,
    "mediumDashDot": xlwt.Borders.MEDIUM_DASH_DOTTED,
    "dashDotDot": xlwt.Borders.THIN_DASH_DOT_DOTTED,
    "mediumDashDotDot": xlwt.Borders.MEDIUM_DASH_DOT_DOTTED,
    "slantDashDot": xlwt.Borders.SLANTED_MEDIUM_DASH_DOTTED,
}
_HORIZONTAL_ALIGNMENTS = {
    None: xlwt.Alignment.HORZ_GENERAL,
    "general": xlwt.Alignment.HORZ_GENERAL,
    "left": xlwt.Alignment.HORZ_LEFT,
    "center": xlwt.Alignment.HORZ_CENTER,
    "right": xlwt.Alignment.HORZ_RIGHT,
    "fill": xlwt.Alignment.HORZ_FILLED,
    "justify": xlwt.Alignment.HORZ_JUSTIFIED,
    "centerContinuous": xlwt.Alignment.HORZ_CENTER_ACROSS_SEL,
    "distributed": xlwt.Alignment.HORZ_DISTRIBUTED,
}
_VERTICAL_ALIGNMENTS = {
    None: xlwt.Alignment.VERT_BOTTOM,
    "top": xlwt.Alignment.VERT_TOP,
    "center": xlwt.Alignment.VERT_CENTER,
    "bottom": xlwt.Alignment.VERT_BOTTOM,
    "justify": xlwt.Alignment.VERT_JUSTIFIED,
    "distributed": xlwt.Alignment.VERT_DISTRIBUTED,
}
_PAPER_SIZE_CODES = {"Letter": 1, "Legal": 5, "A3": 8, "A4": 9, "A5": 11}


class _StyleRegistry:
    """把 ExcelKit Style 转换并缓存为 xlwt XFStyle。"""

    def __init__(self, workbook: xlwt.Workbook) -> None:
        """功能：初始化样式缓存和 XLS 调色板分配状态。

        使用方法：每个 XLS 写出任务创建一个注册器。
        参数：``workbook`` 为目标 xlwt 工作簿，用于设置自定义调色板。
        返回：无。
        """
        self.workbook = workbook
        self.styles: Dict[Tuple[Style, Optional[str]], xlwt.XFStyle] = {}
        self.colors: Dict[str, int] = {}
        self.next_color_index = 8

    def color(self, value: Optional[str]) -> int:
        """功能：取得 ARGB 颜色在 XLS 调色板中的索引。

        使用方法：转换字体、填充和边框时内部调用。
        参数：``value`` 为 8 位 ARGB 字符串或 ``None``；XLS 不支持透明度，
        因此只使用末 6 位 RGB。
        返回：0 表示自动色，或 8～63 范围内的自定义调色板索引。
        异常：同一工作簿自定义颜色超过 56 种时抛出 ``ValueError``。
        """
        if value is None:
            return 0
        rgb = value[-6:]
        if rgb not in self.colors:
            if self.next_color_index > 63:
                raise ValueError("XLS 单个工作簿最多支持 56 种自定义颜色")
            red, green, blue = (int(rgb[index:index + 2], 16) for index in (0, 2, 4))
            self.workbook.set_colour_RGB(self.next_color_index, red, green, blue)
            self.colors[rgb] = self.next_color_index
            self.next_color_index += 1
        return self.colors[rgb]

    def convert(self, style: Style, date_format: Optional[str] = None) -> xlwt.XFStyle:
        """功能：取得 Style 与可选日期格式对应的 xlwt XFStyle。

        使用方法：写出每个值、公式或纯样式单元格前调用。
        参数：``style`` 为 ExcelKit 样式；``date_format`` 仅在样式仍为 General
        时提供日期或日期时间默认格式。
        返回：可传给 ``xlwt.Worksheet.write`` 的缓存 XFStyle。
        """
        key = (style, date_format)
        if key in self.styles:
            return self.styles[key]
        result = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = style.font.name
        font.height = int(round(style.font.size * 20))
        font.bold = style.font.bold
        font.italic = style.font.italic
        font.underline = (
            xlwt.Font.UNDERLINE_SINGLE if style.font.underline else xlwt.Font.UNDERLINE_NONE
        )
        font.colour_index = self.color(style.font.color)
        result.font = font

        pattern = xlwt.Pattern()
        if style.fill.color:
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = self.color(style.fill.color)
        result.pattern = pattern

        borders = xlwt.Borders()
        for name in ("left", "right", "top", "bottom"):
            side = getattr(style.border, name)
            setattr(borders, name, _BORDER_STYLES[side.style])
            setattr(borders, f"{name}_colour", self.color(side.color))
        result.borders = borders

        alignment = xlwt.Alignment()
        alignment.horz = _HORIZONTAL_ALIGNMENTS[style.alignment.horizontal]
        alignment.vert = _VERTICAL_ALIGNMENTS[style.alignment.vertical]
        alignment.wrap = int(style.alignment.wrap_text)
        result.alignment = alignment
        result.num_format_str = (
            date_format if date_format and style.number_format == "General"
            else style.number_format
        )
        self.styles[key] = result
        return result


def _date_format(value: Any) -> Optional[str]:
    """功能：为日期值选择 XLS 默认数字格式。

    使用方法：单元格样式未指定数字格式时由写出器调用。
    参数：``value`` 为当前普通值。
    返回：日期格式、日期时间格式或非日期值的 ``None``。
    """
    if isinstance(value, datetime):
        return "yyyy-mm-dd hh:mm:ss"
    if isinstance(value, date):
        return "yyyy-mm-dd"
    return None


class XlsWriter:
    """把 ExcelKit Workbook 写成 Excel 97–2003 二进制 XLS 文件。"""

    def __init__(self, workbook: "Workbook") -> None:
        """功能：创建绑定到指定工作簿的 XLS 写出器。

        使用方法：业务代码使用 ``workbook.save('output.xls')`` 间接创建。
        参数：``workbook`` 为待写出的 ExcelKit 工作簿。
        返回：无。
        """
        self.workbook = workbook

    def write(self, filename: Union[str, os.PathLike]) -> None:
        """功能：以临时文件和原子替换方式写出 XLS 文件。

        使用方法：``XlsWriter(workbook).write('demo.xls')``。
        参数：``filename`` 为字符串或 ``os.PathLike`` 路径；父目录必须存在。
        返回：``None``。
        异常：类型、路径、XLS 行列上限、公式、颜色或文件系统错误会抛出相应异常；
        已有目标文件在写出失败时保持不变。
        """
        if not isinstance(filename, (str, os.PathLike)):
            raise TypeError("filename 必须是字符串或 PathLike 对象")
        target = Path(filename)
        if not target.name:
            raise ValueError("filename 必须指定文件名")
        if not self.workbook.sheets:
            self.workbook.active

        output = xlwt.Workbook(encoding="utf-8")
        registry = _StyleRegistry(output)
        for worksheet in self.workbook.sheets:
            if worksheet.max_row >= 65536 or worksheet.max_column >= 256:
                raise ValueError("XLS 仅支持 65536 行和 256 列，当前数据已超过格式上限")
            target_sheet = output.add_sheet(worksheet.name)
            target_sheet.set_show_grid(worksheet.show_gridlines)
            for row_index, dimension in worksheet._rows.items():
                if dimension._is_default():
                    continue
                target_row = target_sheet.row(row_index)
                if dimension.height is not None:
                    target_row.height = int(round(dimension.height * 20))
                    target_row.height_mismatch = True
                target_row.hidden = int(dimension.hidden)
            for column_index, dimension in worksheet._columns.items():
                if dimension._is_default():
                    continue
                target_column = target_sheet.col(column_index)
                if dimension.width is not None:
                    target_column.width = int(round(dimension.width * 256))
                target_column.hidden = int(dimension.hidden)
            if worksheet.freeze_panes is not None:
                from ..address import cell_index

                freeze_row, freeze_column = cell_index(worksheet.freeze_panes)
                target_sheet.set_panes_frozen(True)
                target_sheet.set_horz_split_pos(freeze_row)
                target_sheet.set_vert_split_pos(freeze_column)
                target_sheet.set_horz_split_first_visible(freeze_row)
                target_sheet.set_vert_split_first_visible(freeze_column)

            page = worksheet.page
            target_sheet.set_portrait(page.orientation == "portrait")
            target_sheet.set_paper_size_code(_PAPER_SIZE_CODES[page.paper_size])
            target_sheet.set_print_in_rows(page.print_order == "over_then_down")
            target_sheet.set_print_colour(not page.black_and_white)
            target_sheet.set_print_draft(page.draft)
            target_sheet.set_print_centered_horz(page.center_horizontal)
            target_sheet.set_print_centered_vert(page.center_vertical)
            target_sheet.set_print_grid(page.print_gridlines)
            target_sheet.set_print_headers(page.print_headings)
            if worksheet.horizontal_page_breaks:
                target_sheet.set_horz_page_breaks(
                    [(row, 0, 255) for row in worksheet.horizontal_page_breaks]
                )
            if page.scale is not None:
                target_sheet.set_print_scaling(page.scale)
            else:
                target_sheet.set_fit_width_to_pages(page._fit_width or 0)
                target_sheet.set_fit_height_to_pages(page._fit_height or 0)
                target_sheet.set_fit_num_pages(1)
            margins = page.margins
            target_sheet.set_left_margin(margins.left / 2.54)
            target_sheet.set_right_margin(margins.right / 2.54)
            target_sheet.set_top_margin(margins.top / 2.54)
            target_sheet.set_bottom_margin(margins.bottom / 2.54)
            target_sheet.set_header_margin(margins.header / 2.54)
            target_sheet.set_footer_margin(margins.footer / 2.54)
            target_sheet._Worksheet__header_str = header_footer_text(page.header)
            target_sheet._Worksheet__footer_str = header_footer_text(page.footer)
            coordinates = (
                {coordinate for coordinate, _value in worksheet._values.items()}
                | set(worksheet._formulas)
                | set(worksheet._styles)
            )
            for row, column in sorted(coordinates):
                anchor = worksheet._merged_anchor(row, column)
                if anchor is not None and anchor != (row, column):
                    continue
                value = worksheet._values.get(row, column)
                style = worksheet._styles.get((row, column), DEFAULT_STYLE)
                target_style = registry.convert(style, _date_format(value))
                formula = worksheet._formulas.get((row, column))
                if formula is not None:
                    target_sheet.write(row, column, xlwt.Formula(formula.lstrip("=")), target_style)
                elif value is None:
                    target_sheet.write(row, column, None, target_style)
                elif isinstance(value, (str, bool, int, float, date, datetime)):
                    target_sheet.write(row, column, value, target_style)
                else:
                    target_sheet.write(row, column, str(value), target_style)
            for min_row, min_column, max_row, max_column in worksheet._merged_ranges:
                anchor_value = worksheet._values.get(min_row, min_column)
                anchor_style = worksheet._styles.get(
                    (min_row, min_column), DEFAULT_STYLE
                )
                target_sheet.merge(
                    min_row,
                    max_row,
                    min_column,
                    max_column,
                    registry.convert(anchor_style, _date_format(anchor_value)),
                )

        temporary_name = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb", prefix=f".{target.name}.", suffix=".tmp",
                dir=target.parent, delete=False
            ) as temporary:
                temporary_name = temporary.name
            output.save(temporary_name)
            os.replace(temporary_name, target)
            temporary_name = None
        finally:
            if temporary_name is not None:
                try:
                    os.unlink(temporary_name)
                except FileNotFoundError:
                    pass


__all__ = ["XlsWriter"]
