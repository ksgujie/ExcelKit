"""Open XML Excel 工作簿的完整内部读取实现。"""

from __future__ import annotations

import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import TYPE_CHECKING, Any, Dict, List, Sequence, Set, Type, TypeVar

from ..address import cell_address, cell_index, column_to_index, range_index
from ..core.page import HeaderFooter, PageMargins
from ..errors import InvalidFileError
from ..style import DEFAULT_STYLE, Style
from .styles import _read_styles

if TYPE_CHECKING:
    from ..core.workbook import Workbook
    from ..core.worksheet import Worksheet

_WorkbookType = TypeVar("_WorkbookType", bound="Workbook")
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_INTEGER_PATTERN = re.compile(r"^[+-]?\d+$")
_ABSOLUTE_AREA_PATTERN = re.compile(
    r"\$([A-Za-z]{1,3})\$(\d+):\$([A-Za-z]{1,3})\$(\d+)"
)
_ABSOLUTE_ROW_PATTERN = re.compile(r"\$(\d+):\$(\d+)")
_ABSOLUTE_COLUMN_PATTERN = re.compile(r"\$([A-Za-z]{1,3}):\$([A-Za-z]{1,3})")
_PAPER_SIZE_NAMES = {1: "Letter", 5: "Legal", 8: "A3", 9: "A4", 11: "A5"}


def _tag(namespace: str, local_name: str) -> str:
    """功能：生成 ElementTree 使用的 XML 限定名称。

    使用方法：读取器查找 XML 元素和属性时内部调用。
    参数：``namespace`` 为命名空间 URI；``local_name`` 为本地名称。
    返回：``{命名空间}本地名称`` 字符串。
    """
    return f"{{{namespace}}}{local_name}"


def _read_xml(package: zipfile.ZipFile, member: str) -> ET.Element:
    """功能：读取并解析 XLSX 文件包中的一个 XML 部件。

    使用方法：内部传入已打开文件包和包内路径。
    参数：``package`` 为 :class:`ZipFile`；``member`` 为正斜杠包内路径。
    返回：解析后的 XML 根元素。
    异常：部件缺失或 XML 损坏时抛出 :class:`InvalidFileError`。
    """
    try:
        return ET.fromstring(package.read(member))
    except (KeyError, ET.ParseError) as error:
        raise InvalidFileError(f"XLSX 部件缺失或损坏：{member}") from error


def _relationship_targets(package: zipfile.ZipFile) -> Dict[str, str]:
    """功能：读取工作簿关系编号到包内目标路径的映射。

    使用方法：定位工作表 XML 时内部调用。
    参数：``package`` 为已打开的 XLSX 文件包。
    返回：键为关系 ID、值为规范化包内路径的字典。
    异常：关系越出文件包根目录时抛出 :class:`InvalidFileError`。
    """
    root = _read_xml(package, "xl/_rels/workbook.xml.rels")
    targets: Dict[str, str] = {}
    for relationship in root.findall(_tag(_PACKAGE_REL_NS, "Relationship")):
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        if not relationship_id or not target or relationship.get("TargetMode") == "External":
            continue
        target = target.replace("\\", "/")
        normalized = (
            posixpath.normpath(target.lstrip("/"))
            if target.startswith("/")
            else posixpath.normpath(posixpath.join("xl", target))
        )
        if normalized == ".." or normalized.startswith("../"):
            raise InvalidFileError(f"非法的 XLSX 关系目标：{target!r}")
        targets[relationship_id] = normalized
    return targets


def _all_text(element: ET.Element) -> str:
    """功能：连接共享字符串或内联富文本中的全部文本节点。

    使用方法：解析字符串表和字符串单元格时内部调用。
    参数：``element`` 为包含一个或多个 ``t`` 节点的 XML 元素。
    返回：按文档顺序拼接的字符串。
    """
    return "".join(text.text or "" for text in element.iter(_tag(_MAIN_NS, "t")))


def _shared_strings(package: zipfile.ZipFile) -> List[str]:
    """功能：读取 XLSX 共享字符串表。

    使用方法：主读取流程在读取工作表前调用一次。
    参数：``package`` 为已打开的 XLSX 文件包。
    返回：按 0-based 索引排列的字符串列表；没有该部件时返回空列表。
    异常：XML 损坏时抛出 :class:`InvalidFileError`。
    """
    if "xl/sharedStrings.xml" not in package.namelist():
        return []
    root = _read_xml(package, "xl/sharedStrings.xml")
    return [_all_text(item) for item in root.findall(_tag(_MAIN_NS, "si"))]


def _number(value: str) -> Any:
    """功能：把数值文本还原为 Python int 或 float。

    使用方法：读取普通数值单元格时内部调用。
    参数：``value`` 为 ``v`` 元素中的数值字符串。
    返回：整数字符串返回 ``int``，其他合法数值返回 ``float``。
    异常：文本不是合法数字时抛出 :class:`InvalidFileError`。
    """
    try:
        return int(value) if _INTEGER_PATTERN.fullmatch(value) else float(value)
    except ValueError as error:
        raise InvalidFileError(f"无效的 XLSX 数值：{value!r}") from error


def _iso_date(value: str) -> Any:
    """功能：把 Open XML ISO 日期文本还原为 Python 日期对象。

    使用方法：读取 ``t='d'`` 单元格时内部调用。
    参数：``value`` 为 ISO 日期或日期时间字符串。
    返回：日期返回 ``date``，含时间返回 ``datetime``。
    异常：内容无效时抛出 :class:`InvalidFileError`。
    """
    try:
        if "T" in value or " " in value:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        return date.fromisoformat(value)
    except ValueError as error:
        raise InvalidFileError(f"无效的 XLSX ISO 日期：{value!r}") from error


def _excel_serial(value: float, date_1904: bool) -> Any:
    """功能：把 Excel 浮点日期序列转换为 Python date 或 datetime。

    使用方法：读取带日期数字格式的数值单元格时内部调用。
    参数：``value`` 为日期序列；``date_1904`` 表示是否使用 1904 日期系统。
    返回：整数天返回 ``date``，包含时间小数时返回 ``datetime``。
    """
    whole_days = int(value)
    fraction = value - whole_days
    epoch = datetime(1904, 1, 1) if date_1904 else datetime(1899, 12, 30)
    if not date_1904 and 0 < whole_days < 60:
        whole_days += 1
    result = epoch + timedelta(days=whole_days, seconds=round(fraction * 86400, 6))
    return result.date() if abs(fraction) < 1e-12 else result


def _cell_value(
    cell: ET.Element,
    shared_strings: Sequence[str],
    style_index: int,
    date_styles: Set[int],
    date_1904: bool,
) -> Any:
    """功能：根据 XLSX 单元格类型和样式还原普通 Python 值。

    使用方法：工作表读取循环为每个非公式单元格调用。
    参数：``cell`` 为单元格 XML；``shared_strings`` 为共享字符串；
    ``style_index`` 为 0-based 样式索引；``date_styles`` 为日期样式集合；
    ``date_1904`` 为日期系统开关。
    返回：字符串、布尔值、数值、日期时间或 ``None``。
    异常：索引或单元格内容无效时抛出 :class:`InvalidFileError`。
    """
    cell_type = cell.get("t", "n")
    if cell_type == "inlineStr":
        inline = cell.find(_tag(_MAIN_NS, "is"))
        return "" if inline is None else _all_text(inline)
    value_element = cell.find(_tag(_MAIN_NS, "v"))
    if value_element is None or value_element.text is None:
        return None
    value = value_element.text
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError) as error:
            raise InvalidFileError(f"无效的共享字符串索引：{value!r}") from error
    if cell_type == "b":
        if value not in {"0", "1"}:
            raise InvalidFileError(f"无效的布尔单元格值：{value!r}")
        return value == "1"
    if cell_type == "d":
        return _iso_date(value)
    if cell_type in {"str", "e"}:
        return value
    number = _number(value)
    return _excel_serial(float(number), date_1904) if style_index in date_styles else number


def _bool_attribute(value: str | None, default: bool = False) -> bool:
    """功能：把Open XML布尔属性文本转换为Python布尔值。

    使用方法：读取工作表视图和打印开关时内部调用。
    参数：``value`` 为 ``1/0``、``true/false`` 或 ``None``；``default`` 为缺失值。
    返回：解析后的布尔值。
    """
    if value is None:
        return default
    return value in {"1", "true", "True"}


def _header_footer(value: str | None) -> HeaderFooter:
    """功能：把Excel的 ``&L/&C/&R`` 文本拆分为HeaderFooter对象。

    使用方法：读取奇数页页眉和页脚时内部调用。
    参数：``value`` 为控制代码字符串或 ``None``；``&&`` 作为字面量与号保留。
    返回：左、中、右区域组成的 :class:`HeaderFooter`。
    """
    if not value:
        return HeaderFooter()
    sections = {"L": [], "C": [], "R": []}
    current = "C"
    index = 0
    while index < len(value):
        if value[index] == "&" and index + 1 < len(value):
            code = value[index + 1]
            if code == "&":
                sections[current].append("&&")
                index += 2
                continue
            if code in sections:
                current = code
                index += 2
                continue
        sections[current].append(value[index])
        index += 1
    return HeaderFooter(
        left="".join(sections["L"]),
        center="".join(sections["C"]),
        right="".join(sections["R"]),
    )


def _load_sheet_layout(root: ET.Element, worksheet: "Worksheet") -> None:
    """功能：读取工作表合并、尺寸、视图、筛选和页面打印设置。

    使用方法：单张工作表值和样式读取完成后调用。
    参数：``root`` 为工作表XML根元素；``worksheet`` 为目标工作表。
    返回：``None``；解析结果写入目标对象。
    异常：地址、数值或页面属性损坏时抛出 :class:`InvalidFileError`。
    """
    try:
        sheet_view = root.find(
            f"{_tag(_MAIN_NS, 'sheetViews')}/{_tag(_MAIN_NS, 'sheetView')}"
        )
        if sheet_view is not None:
            worksheet.show_gridlines = _bool_attribute(
                sheet_view.get("showGridLines"), True
            )
            pane = sheet_view.find(_tag(_MAIN_NS, "pane"))
            if pane is not None and pane.get("state") in {"frozen", "frozenSplit"}:
                top_left = pane.get("topLeftCell")
                if top_left:
                    worksheet.freeze_panes = top_left
                else:
                    row = int(float(pane.get("ySplit", "0")))
                    column = int(float(pane.get("xSplit", "0")))
                    worksheet.freeze_panes = cell_address(row, column)

        columns = root.find(_tag(_MAIN_NS, "cols"))
        if columns is not None:
            for source in columns.findall(_tag(_MAIN_NS, "col")):
                minimum = int(source.get("min", "0")) - 1
                maximum = int(source.get("max", "0")) - 1
                if minimum < 0 or maximum < minimum:
                    raise ValueError("无效的列尺寸范围")
                width = (
                    float(source.get("width", "0"))
                    if "width" in source.attrib else None
                )
                hidden = _bool_attribute(source.get("hidden"))
                for column_index in range(minimum, maximum + 1):
                    dimension = worksheet.column(column_index)
                    dimension.width = width
                    dimension.hidden = hidden

        sheet_data = root.find(_tag(_MAIN_NS, "sheetData"))
        if sheet_data is not None:
            for source in sheet_data.findall(_tag(_MAIN_NS, "row")):
                row_index = int(source.get("r", "0")) - 1
                if row_index < 0:
                    raise ValueError("无效的行尺寸索引")
                if "ht" in source.attrib or _bool_attribute(source.get("hidden")):
                    dimension = worksheet.row(row_index)
                    dimension.height = (
                        float(source.get("ht", "0"))
                        if "ht" in source.attrib else None
                    )
                    dimension.hidden = _bool_attribute(source.get("hidden"))

        merge_cells = root.find(_tag(_MAIN_NS, "mergeCells"))
        if merge_cells is not None:
            for source in merge_cells.findall(_tag(_MAIN_NS, "mergeCell")):
                reference = source.get("ref")
                if not reference:
                    raise ValueError("合并区域缺少地址")
                worksheet._merge_range(*range_index(reference))

        auto_filter = root.find(_tag(_MAIN_NS, "autoFilter"))
        if auto_filter is not None and auto_filter.get("ref"):
            reference = auto_filter.get("ref") or ""
            worksheet.auto_filter_range = (
                reference if ":" in reference else f"{reference}:{reference}"
            )

        options = root.find(_tag(_MAIN_NS, "printOptions"))
        if options is not None:
            worksheet.page.center_horizontal = _bool_attribute(
                options.get("horizontalCentered")
            )
            worksheet.page.center_vertical = _bool_attribute(
                options.get("verticalCentered")
            )
            worksheet.page.print_gridlines = _bool_attribute(options.get("gridLines"))
            worksheet.page.print_headings = _bool_attribute(options.get("headings"))

        margins = root.find(_tag(_MAIN_NS, "pageMargins"))
        if margins is not None:
            defaults = worksheet.page.margins
            worksheet.page.margins = PageMargins(
                **{
                    name: float(margins.get(name, str(getattr(defaults, name) / 2.54)))
                    * 2.54
                    for name in ("left", "right", "top", "bottom", "header", "footer")
                }
            )

        setup = root.find(_tag(_MAIN_NS, "pageSetup"))
        if setup is not None:
            if setup.get("orientation") in {"portrait", "landscape"}:
                worksheet.page.orientation = setup.get("orientation")
            if setup.get("paperSize"):
                paper_code = int(setup.get("paperSize", "9"))
                if paper_code in _PAPER_SIZE_NAMES:
                    worksheet.page.paper_size = _PAPER_SIZE_NAMES[paper_code]
            if setup.get("pageOrder") in {"downThenOver", "overThenDown"}:
                worksheet.page.print_order = (
                    "down_then_over"
                    if setup.get("pageOrder") == "downThenOver"
                    else "over_then_down"
                )
            worksheet.page.black_and_white = _bool_attribute(
                setup.get("blackAndWhite")
            )
            worksheet.page.draft = _bool_attribute(setup.get("draft"))
            fit_width = int(setup.get("fitToWidth", "0")) or None
            fit_height = int(setup.get("fitToHeight", "0")) or None
            if fit_width is not None or fit_height is not None:
                worksheet.page.fit(width=fit_width, height=fit_height)
            elif setup.get("scale"):
                worksheet.page.scale = int(setup.get("scale", "100"))
            if _bool_attribute(setup.get("useFirstPageNumber")) and setup.get(
                "firstPageNumber"
            ):
                worksheet.page.first_page_number = int(
                    setup.get("firstPageNumber", "1")
                )

        header_footer = root.find(_tag(_MAIN_NS, "headerFooter"))
        if header_footer is not None:
            odd_header = header_footer.find(_tag(_MAIN_NS, "oddHeader"))
            odd_footer = header_footer.find(_tag(_MAIN_NS, "oddFooter"))
            worksheet.page.header = _header_footer(
                odd_header.text if odd_header is not None else None
            )
            worksheet.page.footer = _header_footer(
                odd_footer.text if odd_footer is not None else None
            )
    except (TypeError, ValueError, KeyError) as error:
        raise InvalidFileError("工作表布局或打印设置无效") from error


def _load_defined_names(root: ET.Element, workbook: "Workbook") -> None:
    """功能：读取打印区域和重复打印行列的工作簿定义名称。

    使用方法：全部工作表加载完成后调用。
    参数：``root`` 为workbook.xml根元素；``workbook`` 为目标工作簿。
    返回：``None``；只处理带本地工作表索引的标准打印名称。
    异常：本库支持的定义名称内容损坏时抛出 :class:`InvalidFileError`。
    """
    container = root.find(_tag(_MAIN_NS, "definedNames"))
    if container is None:
        return
    try:
        for item in container.findall(_tag(_MAIN_NS, "definedName")):
            if item.text is None or item.get("localSheetId") is None:
                continue
            sheet_index = int(item.get("localSheetId", "-1"))
            if not 0 <= sheet_index < len(workbook.sheets):
                continue
            page = workbook.sheet(sheet_index).page
            if item.get("name") == "_xlnm.Print_Area":
                match = _ABSOLUTE_AREA_PATTERN.search(item.text)
                if match:
                    page.print_area = (
                        f"{match.group(1)}{match.group(2)}:"
                        f"{match.group(3)}{match.group(4)}"
                    )
            elif item.get("name") == "_xlnm.Print_Titles":
                row_match = _ABSOLUTE_ROW_PATTERN.search(item.text)
                column_match = _ABSOLUTE_COLUMN_PATTERN.search(item.text)
                if row_match:
                    page.repeat_rows = (
                        int(row_match.group(1)) - 1,
                        int(row_match.group(2)) - 1,
                    )
                if column_match:
                    page.repeat_columns = (
                        column_to_index(column_match.group(1)),
                        column_to_index(column_match.group(2)),
                    )
    except (TypeError, ValueError) as error:
        raise InvalidFileError("打印区域或重复标题定义无效") from error


def _load_sheet(
    package: zipfile.ZipFile,
    member: str,
    worksheet: "Worksheet",
    shared_strings: Sequence[str],
    styles: Sequence[Style],
    date_styles: Set[int],
    date_1904: bool,
) -> None:
    """功能：读取单张工作表的值、公式、样式和已触及范围。

    使用方法：XLSX 主流程按工作表顺序调用。
    参数：``package`` 为 ZIP 包；``member`` 为工作表路径；``worksheet`` 为目标；
    其余参数分别为共享字符串、样式、日期样式集合和日期系统。
    返回：``None``；读取结果直接写入工作表。
    异常：XML、地址、样式索引或数据损坏时抛出相应异常。
    """
    root = _read_xml(package, member)
    properties = root.find(_tag(_MAIN_NS, "sheetPr"))
    tab_color = (
        properties.find(_tag(_MAIN_NS, "tabColor"))
        if properties is not None
        else None
    )
    if tab_color is not None and tab_color.get("rgb"):
        try:
            worksheet.color = tab_color.get("rgb")
        except ValueError as error:
            raise InvalidFileError("工作表标签颜色不是有效的 RGB 或 ARGB 值") from error
    for cell in root.findall(f".//{_tag(_MAIN_NS, 'sheetData')}//{_tag(_MAIN_NS, 'c')}"):
        address = cell.get("r")
        if not address:
            raise InvalidFileError("工作表中存在缺少地址的单元格")
        row, column = cell_index(address)
        try:
            style_index = int(cell.get("s", "0"))
            style = styles[style_index]
        except (ValueError, IndexError) as error:
            raise InvalidFileError(f"无效的单元格样式索引：{cell.get('s')!r}") from error
        formula = cell.find(_tag(_MAIN_NS, "f"))
        if formula is not None and formula.text:
            worksheet._set_formula(row, column, formula.text)
        else:
            value = _cell_value(cell, shared_strings, style_index, date_styles, date_1904)
            worksheet._touch(row, column)
            worksheet._values.set(row, column, value)
        if style != DEFAULT_STYLE:
            worksheet._styles[(row, column)] = style
    _load_sheet_layout(root, worksheet)


def _load_xlsx(
    workbook_class: Type[_WorkbookType], filename: os.PathLike | str
) -> _WorkbookType:
    """功能：解析 Open XML Excel 文件并构造 Workbook。

    使用方法：仅由 ``Workbook.load(filename)`` 的格式分派器调用。
    参数：``workbook_class`` 为 Workbook 类或子类；``filename`` 为源路径。
    返回：包含工作表、值、公式、日期和受支持样式的新工作簿。
    异常：文件不存在时抛出 ``FileNotFoundError``；文件包或 XML 损坏时抛出
    :class:`InvalidFileError`。
    """
    path = Path(filename)
    try:
        package = zipfile.ZipFile(path)
    except zipfile.BadZipFile as error:
        raise InvalidFileError(f"文件不是有效的 Open XML 工作簿：{path}") from error
    with package:
        workbook_root = _read_xml(package, "xl/workbook.xml")
        targets = _relationship_targets(package)
        shared_strings = _shared_strings(package)
        styles, date_styles = _read_styles(package)
        properties = workbook_root.find(_tag(_MAIN_NS, "workbookPr"))
        date_1904 = properties is not None and properties.get("date1904", "0") in {
            "1", "true", "True"
        }
        workbook = workbook_class()
        sheets = workbook_root.find(_tag(_MAIN_NS, "sheets"))
        if sheets is None:
            raise InvalidFileError("workbook.xml 缺少 sheets 元素")
        for sheet in sheets.findall(_tag(_MAIN_NS, "sheet")):
            name = sheet.get("name")
            relationship_id = sheet.get(_tag(_REL_NS, "id"))
            if not name or not relationship_id or relationship_id not in targets:
                raise InvalidFileError("工作表名称或关系信息不完整")
            worksheet = workbook.add_sheet(name)
            _load_sheet(
                package, targets[relationship_id], worksheet, shared_strings,
                styles, date_styles, date_1904
            )
        _load_defined_names(workbook_root, workbook)
        return workbook


__all__ = []
