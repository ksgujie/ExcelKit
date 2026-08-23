"""XLSX 样式表到 ExcelKit 不可变样式对象的内部转换。"""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Set, Tuple

from ..errors import InvalidFileError
from ..style import Alignment, Border, BorderSide, DEFAULT_STYLE, Fill, Font, Style

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_BUILTIN_NUMBER_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
}
_BUILTIN_DATE_FORMAT_IDS = {
    14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33,
    34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
}


def _tag(local_name: str) -> str:
    """功能：生成 SpreadsheetML 主命名空间限定名称。

    使用方法：样式读取函数查找 XML 元素时内部调用。
    参数：``local_name`` 为 XML 元素本地名称。
    返回：``{主命名空间}本地名称`` 字符串。
    """
    return f"{{{_MAIN_NS}}}{local_name}"


def _rgb(element: Optional[ET.Element]) -> Optional[str]:
    """功能：读取样式颜色元素中的 RGB 或 ARGB 值。

    使用方法：解析字体、填充和边框颜色时内部调用。
    参数：``element`` 为 ``color``、``fgColor`` 元素或 ``None``。
    返回：可由 ExcelKit 样式接受的颜色字符串；主题色和索引色返回 ``None``。
    """
    if element is None:
        return None
    value = element.get("rgb")
    return value if value and len(value) in {6, 8} else None


def _is_date_format(format_code: str) -> bool:
    """功能：判断自定义 Excel 数字格式是否表示日期或时间。

    使用方法：解析 ``numFmt`` 与 ``cellXfs`` 时内部调用。
    参数：``format_code`` 为 Excel 数字格式代码。
    返回：包含有效日期或时间标记时返回 ``True``，否则返回 ``False``。
    """
    cleaned = re.sub(r'"[^"\\]*(?:\\.[^"\\]*)*"', "", format_code.lower())
    cleaned = re.sub(r"\\.", "", cleaned)
    cleaned = re.sub(r"\[(?!h+\]|m+\]|s+\])[^\]]*\]", "", cleaned)
    return bool(re.search(r"y+|d+|h+|s+|\[(?:h+|m+|s+)\]", cleaned))


def _font(element: ET.Element) -> Font:
    """功能：把一个 XLSX font 元素转换为 ExcelKit Font。

    使用方法：读取字体集合时内部调用。
    参数：``element`` 为 ``font`` XML 元素。
    返回：包含受支持字体名称、字号、字形和颜色的 :class:`Font`。
    """
    name_element = element.find(_tag("name"))
    size_element = element.find(_tag("sz"))
    try:
        size = float(size_element.get("val", "11")) if size_element is not None else 11.0
    except ValueError:
        size = 11.0
    return Font(
        name=name_element.get("val", "Calibri") if name_element is not None else "Calibri",
        size=size if size > 0 else 11.0,
        bold=element.find(_tag("b")) is not None,
        italic=element.find(_tag("i")) is not None,
        underline=element.find(_tag("u")) is not None,
        color=_rgb(element.find(_tag("color"))),
    )


def _fill(element: ET.Element) -> Fill:
    """功能：把一个 XLSX fill 元素转换为 ExcelKit Fill。

    使用方法：读取填充集合时内部调用。
    参数：``element`` 为 ``fill`` XML 元素。
    返回：实心前景色填充；其他填充类型返回默认无填充。
    """
    pattern = element.find(_tag("patternFill"))
    if pattern is None or pattern.get("patternType") != "solid":
        return Fill()
    return Fill(_rgb(pattern.find(_tag("fgColor"))))


def _side(element: Optional[ET.Element]) -> BorderSide:
    """功能：把 XLSX 边框边元素转换为 ExcelKit BorderSide。

    使用方法：读取 Border 的四条边时内部调用。
    参数：``element`` 为 left、right、top、bottom 元素或 ``None``。
    返回：受支持线型和颜色组成的 :class:`BorderSide`；不支持线型回退为空边。
    """
    if element is None:
        return BorderSide()
    try:
        return BorderSide(element.get("style"), _rgb(element.find(_tag("color"))))
    except ValueError:
        return BorderSide()


def _border(element: ET.Element) -> Border:
    """功能：把一个 XLSX border 元素转换为 ExcelKit Border。

    使用方法：读取边框集合时内部调用。
    参数：``element`` 为 ``border`` XML 元素。
    返回：包含左、右、上、下四边的 :class:`Border`。
    """
    return Border(
        left=_side(element.find(_tag("left"))),
        right=_side(element.find(_tag("right"))),
        top=_side(element.find(_tag("top"))),
        bottom=_side(element.find(_tag("bottom"))),
    )


def _item(items: List[object], index_text: str, default: object) -> object:
    """功能：按 XML 索引安全取得样式组件。

    使用方法：组合 cellXfs 样式时内部调用。
    参数：``items`` 为组件列表；``index_text`` 为索引文本；``default`` 为回退值。
    返回：索引存在时返回对应组件，索引无效或越界时返回默认组件。
    """
    try:
        index = int(index_text)
        return items[index]
    except (ValueError, IndexError):
        return default


def _read_styles(package: zipfile.ZipFile) -> Tuple[List[Style], Set[int]]:
    """功能：读取 XLSX 样式表并生成单元格样式及日期样式索引。

    使用方法：XLSX 主读取流程在读取工作表前调用一次。
    参数：``package`` 为已经打开的 :class:`zipfile.ZipFile`。
    返回：二元组；第一项是按 0-based cellXfs 索引排列的 Style，第二项是日期
    或时间数字格式的样式索引集合。没有样式表时返回默认样式列表和空集合。
    异常：``styles.xml`` 不是合法 XML 时抛出 :class:`InvalidFileError`。
    """
    member = "xl/styles.xml"
    if member not in package.namelist():
        return [DEFAULT_STYLE], set()
    try:
        root = ET.fromstring(package.read(member))
    except (KeyError, ET.ParseError) as error:
        raise InvalidFileError("XLSX 样式表损坏") from error

    custom_formats: Dict[int, str] = {}
    for element in root.findall(f"{_tag('numFmts')}/{_tag('numFmt')}"):
        try:
            custom_formats[int(element.get("numFmtId", "-1"))] = element.get(
                "formatCode", "General"
            )
        except ValueError:
            continue

    fonts_parent = root.find(_tag("fonts"))
    fills_parent = root.find(_tag("fills"))
    borders_parent = root.find(_tag("borders"))
    fonts = [_font(item) for item in fonts_parent] if fonts_parent is not None else [Font()]
    fills = [_fill(item) for item in fills_parent] if fills_parent is not None else [Fill()]
    borders = (
        [_border(item) for item in borders_parent]
        if borders_parent is not None
        else [Border()]
    )

    styles: List[Style] = []
    date_styles: Set[int] = set()
    cell_formats = root.find(_tag("cellXfs"))
    if cell_formats is None:
        return [DEFAULT_STYLE], set()
    for style_index, cell_format in enumerate(cell_formats):
        try:
            number_format_id = int(cell_format.get("numFmtId", "0"))
        except ValueError:
            number_format_id = 0
        number_format = custom_formats.get(
            number_format_id, _BUILTIN_NUMBER_FORMATS.get(number_format_id, "General")
        )
        alignment_element = cell_format.find(_tag("alignment"))
        alignment = Alignment()
        if alignment_element is not None:
            horizontal = alignment_element.get("horizontal")
            vertical = alignment_element.get("vertical")
            try:
                alignment = Alignment(
                    horizontal=horizontal,
                    vertical=vertical,
                    wrap_text=alignment_element.get("wrapText", "0") in {"1", "true", "True"},
                )
            except ValueError:
                alignment = Alignment()
        styles.append(
            Style(
                font=_item(fonts, cell_format.get("fontId", "0"), Font()),
                fill=_item(fills, cell_format.get("fillId", "0"), Fill()),
                border=_item(borders, cell_format.get("borderId", "0"), Border()),
                alignment=alignment,
                number_format=number_format,
            )
        )
        if number_format_id in _BUILTIN_DATE_FORMAT_IDS or _is_date_format(number_format):
            date_styles.add(style_index)
    return styles or [DEFAULT_STYLE], date_styles


__all__ = []
