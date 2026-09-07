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

from ..address import cell_address, cell_index, column_to_index, range_address, range_index
from ..core.page import HeaderFooter, PageMargins
from ..errors import InvalidFileError
from ..properties import WorkbookProperties
from ..validation import Validation
from ..conditional import ConditionalFormat
from ..style import DEFAULT_STYLE, Style
from .styles import _read_dxf_colors, _read_styles

if TYPE_CHECKING:
    from ..core.workbook import Workbook
    from ..core.worksheet import Worksheet

_WorkbookType = TypeVar("_WorkbookType", bound="Workbook")
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CORE_PROPERTIES_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_DCTERMS_NS = "http://purl.org/dc/terms/"
_DC_NS = "http://purl.org/dc/elements/1.1/"
_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_EXCELKIT_NS = "https://github.com/ksgujie/ExcelKit"
_INTEGER_PATTERN = re.compile(r"^[+-]?\d+$")
_ABSOLUTE_AREA_PATTERN = re.compile(
    r"\$([A-Za-z]{1,3})\$(\d+):\$([A-Za-z]{1,3})\$(\d+)"
)
_ABSOLUTE_ROW_PATTERN = re.compile(r"\$(\d+):\$(\d+)")
_ABSOLUTE_COLUMN_PATTERN = re.compile(r"\$([A-Za-z]{1,3}):\$([A-Za-z]{1,3})")
_NAMED_RANGE_PATTERN = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^!]+))!"
    r"\$([A-Za-z]{1,3})\$(\d+):\$([A-Za-z]{1,3})\$(\d+)$"
)
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


def _load_core_properties(
    package: zipfile.ZipFile, properties: WorkbookProperties
) -> None:
    """功能：读取 XLSX 核心文档属性并写入 WorkbookProperties。

    使用方法：工作簿主结构创建后由内部调用。
    参数：``package`` 为已打开的XLSX包；``properties`` 为目标属性对象。
    返回：``None``；没有核心属性部件时保持默认值。
    异常：属性 XML 损坏或时间格式无效时抛出 ``InvalidFileError``。
    """
    member = "docProps/core.xml"
    if member not in package.namelist():
        return
    root = _read_xml(package, member)

    def text(namespace: str, name: str) -> str:
        """功能：读取一个核心属性 XML 元素的文本。

        使用方法：由外层核心属性读取流程按字段内部调用。
        参数：``namespace`` 为元素命名空间；``name`` 为不带命名空间的元素名。
        返回：元素文本字符串；元素不存在或没有文本时返回空字符串。
        """
        element = root.find(_tag(namespace, name))
        return "" if element is None else (element.text or "")

    properties.title = text(_DC_NS, "title")
    properties.subject = text(_DC_NS, "subject")
    properties.author = text(_DC_NS, "creator")
    properties.comments = text(_DC_NS, "description")
    properties.category = text(_CORE_PROPERTIES_NS, "category")
    properties.keywords = text(_CORE_PROPERTIES_NS, "keywords")
    properties.last_modified_by = text(_CORE_PROPERTIES_NS, "lastModifiedBy")
    for name in ("created", "modified"):
        value = text(_DCTERMS_NS, name)
        if value:
            try:
                parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
            except ValueError as error:
                raise InvalidFileError(f"核心属性 {name} 的时间格式无效") from error
            setattr(properties, name, parsed)


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


def _load_sheet_layout(
    root: ET.Element,
    worksheet: "Worksheet",
    dxf_colors: Sequence[Tuple[Optional[str], Optional[str]]] = (),
) -> None:
    """功能：读取工作表合并、尺寸、视图、筛选和页面打印设置。

    使用方法：单张工作表值和样式读取完成后调用。
    参数：``root`` 为工作表XML根元素；``worksheet`` 为目标工作表；
    ``dxf_colors`` 为按差异样式索引排列的填充色和字体色。
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
                outline_level = int(source.get("outlineLevel", "0"))
                collapsed = _bool_attribute(source.get("collapsed"))
                for column_index in range(minimum, maximum + 1):
                    dimension = worksheet.column(column_index)
                    dimension.width = width
                    dimension.hidden = hidden
                    dimension.outline_level = outline_level
                    dimension.collapsed = collapsed

        sheet_data = root.find(_tag(_MAIN_NS, "sheetData"))
        if sheet_data is not None:
            for source in sheet_data.findall(_tag(_MAIN_NS, "row")):
                row_index = int(source.get("r", "0")) - 1
                if row_index < 0:
                    raise ValueError("无效的行尺寸索引")
                if (
                    "ht" in source.attrib or _bool_attribute(source.get("hidden"))
                    or source.get("outlineLevel") is not None
                    or _bool_attribute(source.get("collapsed"))
                ):
                    dimension = worksheet.row(row_index)
                    dimension.height = (
                        float(source.get("ht", "0"))
                        if "ht" in source.attrib else None
                    )
                    dimension.hidden = _bool_attribute(source.get("hidden"))
                    dimension.outline_level = int(source.get("outlineLevel", "0"))
                    dimension.collapsed = _bool_attribute(source.get("collapsed"))

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
            worksheet.auto_filter.range = (
                reference if ":" in reference else f"{reference}:{reference}"
            )
            for filter_column in auto_filter.findall(_tag(_MAIN_NS, "filterColumn")):
                try:
                    column = int(filter_column.get("colId", "-1"))
                except ValueError:
                    continue
                filters = filter_column.find(_tag(_MAIN_NS, "filters"))
                if filters is not None and column >= 0:
                    worksheet._filter_conditions[column] = tuple(
                        item.get("val", "") for item in filters.findall(_tag(_MAIN_NS, "filter"))
                    )

        data_validations = root.find(_tag(_MAIN_NS, "dataValidations"))
        if data_validations is not None:
            for source in data_validations.findall(_tag(_MAIN_NS, "dataValidation")):
                address = source.get("sqref")
                kind = source.get("type", "custom")
                if not address:
                    raise ValueError("数据有效性缺少 sqref")
                formula1_element = source.find(_tag(_MAIN_NS, "formula1"))
                formula2_element = source.find(_tag(_MAIN_NS, "formula2"))
                formula1 = formula1_element.text if formula1_element is not None else None
                values = None
                if kind == "list" and formula1 and formula1.startswith('"') and formula1.endswith('"'):
                    values = tuple(formula1[1:-1].split(","))
                    formula1 = None
                worksheet._validations.append(Validation(
                    address, kind=kind, operator=source.get("operator"),
                    formula1=formula1,
                    formula2=formula2_element.text if formula2_element is not None else None,
                    values=values,
                    allow_blank=_bool_attribute(source.get("allowBlank")),
                    show_dropdown=source.get("showDropDown", "0") not in {"1", "true", "True"},
                    prompt_title=source.get("promptTitle"), prompt=source.get("prompt"),
                    error_title=source.get("errorTitle"), error=source.get("error"),
                    error_style=source.get("errorStyle", "stop"),
                ))

        for group in root.findall(_tag(_MAIN_NS, "conditionalFormatting")):
            address = group.get("sqref")
            if not address:
                continue
            for source in group.findall(_tag(_MAIN_NS, "cfRule")):
                formula_element = source.find(_tag(_MAIN_NS, "formula"))
                try:
                    dxf_index = int(source.get("dxfId", "-1"))
                except ValueError:
                    dxf_index = -1
                fill, font = (
                    dxf_colors[dxf_index]
                    if 0 <= dxf_index < len(dxf_colors)
                    else (None, None)
                )
                worksheet._conditionals.append(ConditionalFormat(
                    address, rule=source.get("type", "cellIs"),
                    operator=source.get("operator"),
                    formula=formula_element.text if formula_element is not None else None,
                    fill=fill,
                    font=font,
                    priority=int(source.get("priority", "1")),
                    stop_if_true=_bool_attribute(source.get("stopIfTrue")),
                ))

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
    """功能：读取打印设置名称和工作簿级命名区域。

    使用方法：全部工作表加载完成后调用。
    参数：``root`` 为workbook.xml根元素；``workbook`` 为目标工作簿。
    返回：``None``；处理标准打印名称以及指向单张工作表矩形区域的全局名称。
    异常：本库支持的定义名称内容损坏时抛出 :class:`InvalidFileError`。
    """
    container = root.find(_tag(_MAIN_NS, "definedNames"))
    if container is None:
        return
    try:
        for item in container.findall(_tag(_MAIN_NS, "definedName")):
            if item.text is None:
                continue
            if item.get("localSheetId") is None:
                name = item.get("name")
                match = _NAMED_RANGE_PATTERN.fullmatch(item.text)
                if not name or name.startswith("_xlnm.") or match is None:
                    continue
                sheet_name = (match.group(1) or match.group(2)).replace("''", "'")
                address = (
                    f"{match.group(3)}{match.group(4)}:"
                    f"{match.group(5)}{match.group(6)}"
                )
                workbook.add_named_range(
                    name, workbook.sheet(sheet_name).range(address)
                )
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
    except (TypeError, ValueError, KeyError) as error:
        raise InvalidFileError("打印设置或命名区域定义无效") from error


def _load_sheet(
    package: zipfile.ZipFile,
    member: str,
    worksheet: "Worksheet",
    shared_strings: Sequence[str],
    styles: Sequence[Style],
    date_styles: Set[int],
    date_1904: bool,
    dxf_colors: Sequence[Tuple[Optional[str], Optional[str]]],
) -> None:
    """功能：读取单张工作表的值、公式、样式和已触及范围。

    使用方法：XLSX 主流程按工作表顺序调用。
    参数：``package`` 为 ZIP 包；``member`` 为工作表路径；``worksheet`` 为目标；
    其余参数分别为共享字符串、样式、日期样式集合、日期系统和条件格式差异样式颜色。
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
    protection = root.find(_tag(_MAIN_NS, "sheetProtection"))
    if protection is not None:
        worksheet.protection.enabled = True
        worksheet.protection.password = protection.get("password")
        worksheet.protection.select_locked = protection.get("selectLockedCells", "0") not in {"1", "true", "True"}
        worksheet.protection.select_unlocked = protection.get("selectUnlockedCells", "0") not in {"1", "true", "True"}
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
            worksheet._set_formula(row, column, formula.text, invalidate=False)
            cached = cell.find(_tag(_MAIN_NS, "v"))
            if cached is not None and (
                cached.text is not None or cell.get("t") in {"str", "e"}
            ):
                value = (
                    "" if cached.text is None else
                    _cell_value(cell, shared_strings, style_index, date_styles, date_1904)
                )
                worksheet._set_cached_value(row, column, value)
        else:
            value = _cell_value(cell, shared_strings, style_index, date_styles, date_1904)
            worksheet._touch(row, column)
            worksheet._values.set(row, column, value)
        if style != DEFAULT_STYLE:
            worksheet._styles[(row, column)] = style
    _load_sheet_layout(root, worksheet, dxf_colors)
    _load_sheet_hyperlinks(package, member, root, worksheet)
    _load_sheet_notes(package, member, worksheet)
    _load_sheet_images(package, member, root, worksheet)
    _load_sheet_tables(package, member, root, worksheet)


def _sheet_relationships(
    package: zipfile.ZipFile, member: str
) -> Dict[str, Tuple[str, str]]:
    """功能：读取单张工作表关系部件中的内部目标关系。

    使用方法：读取超链接和数据表时由内部调用。
    参数：``package`` 为已打开的XLSX包；``member`` 为工作表包内路径。
    返回：关系编号到 ``(关系类型, 规范化目标路径)`` 的映射；没有关系部件时返回空字典。
    异常：关系文件损坏、目标越出包根目录时抛出 ``InvalidFileError``。
    """
    directory, filename = posixpath.split(member)
    relationships_member = posixpath.join(directory, "_rels", f"{filename}.rels")
    if relationships_member not in package.namelist():
        return {}
    relationships_root = _read_xml(package, relationships_member)
    targets: Dict[str, Tuple[str, str]] = {}
    for relationship in relationships_root.findall(
        _tag(_PACKAGE_REL_NS, "Relationship")
    ):
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        relationship_type = relationship.get("Type", "")
        if not relationship_id or not target:
            continue
        if relationship.get("TargetMode") == "External":
            targets[relationship_id] = (relationship_type, target)
            continue
        target = target.replace("\\", "/")
        normalized = (
            posixpath.normpath(target.lstrip("/"))
            if target.startswith("/")
            else posixpath.normpath(posixpath.join(directory, target))
        )
        if normalized == ".." or normalized.startswith("../"):
            raise InvalidFileError(f"非法的工作表关系目标：{target!r}")
        targets[relationship_id] = (relationship_type, normalized)
    return targets


def _load_sheet_hyperlinks(
    package: zipfile.ZipFile,
    member: str,
    root: ET.Element,
    worksheet: "Worksheet",
) -> None:
    """功能：读取工作表中的外部和内部超链接。

    使用方法：单张工作表布局读取后由内部调用。
    参数：``package`` 为XLSX包；``member`` 为工作表路径；``root`` 为XML根元素；
    ``worksheet`` 为接收链接的工作表。
    返回：``None``；链接通过 ``Cell.hyperlink`` 写入目标工作表。
    异常：链接地址、关系编号或关系类型无效时抛出 ``InvalidFileError``。
    """
    container = root.find(_tag(_MAIN_NS, "hyperlinks"))
    if container is None:
        return
    relationships = _sheet_relationships(package, member)
    hyperlink_type = "/relationships/hyperlink"
    try:
        for item in container.findall(_tag(_MAIN_NS, "hyperlink")):
            address = item.get("ref")
            if not address or ":" in address:
                raise InvalidFileError("超链接 ref 必须是单个A1地址")
            location = item.get("location")
            relationship_id = item.get(_tag(_REL_NS, "id"))
            target = None
            if relationship_id is not None:
                relationship = relationships.get(relationship_id)
                if relationship is None or not relationship[0].endswith(hyperlink_type):
                    raise InvalidFileError("超链接关系缺失或类型错误")
                target = relationship[1]
            if target is None and location is None:
                raise InvalidFileError("超链接缺少外部目标或内部位置")
            from ..hyperlink import Hyperlink

            worksheet.cell(address).hyperlink = Hyperlink(
                target=target,
                location=location,
                display=item.get("display"),
                tooltip=item.get("tooltip"),
            )
    except (TypeError, ValueError, KeyError) as error:
        if isinstance(error, InvalidFileError):
            raise
        raise InvalidFileError("工作表超链接定义无效") from error


def _load_sheet_notes(
    package: zipfile.ZipFile, member: str, worksheet: "Worksheet"
) -> None:
    """功能：读取工作表关系中引用的传统批注内容。

    使用方法：单张工作表的单元格和布局读取完成后内部调用。
    参数：``package`` 为 XLSX ZIP 包；``member`` 为工作表路径；``worksheet`` 为目标表。
    返回：``None``；批注写入对应 ``Cell.note``。
    异常：批注关系或 XML 损坏时抛出 :class:`InvalidFileError`。
    """
    relationships = _sheet_relationships(package, member)
    targets = [
        target for relationship_type, target in relationships.values()
        if relationship_type.endswith("/comments")
    ]
    if not targets:
        return
    try:
        root = _read_xml(package, targets[0])
        authors = [item.text or "ExcelKit" for item in root.findall(
            f"{_tag(_MAIN_NS, 'authors')}/{_tag(_MAIN_NS, 'author')}"
        )]
        comments = root.find(_tag(_MAIN_NS, "commentList"))
        if comments is None:
            return
        from ..note import Note
        for item in comments.findall(_tag(_MAIN_NS, "comment")):
            address = item.get("ref")
            if not address:
                raise InvalidFileError("批注缺少单元格地址")
            try:
                author = authors[int(item.get("authorId", "0"))]
            except (ValueError, IndexError) as error:
                raise InvalidFileError("批注作者索引无效") from error
            text_element = item.find(_tag(_MAIN_NS, "text"))
            text = _all_text(text_element) if text_element is not None else ""
            worksheet.cell(address).note = Note(text or " ", author=author)
    except (TypeError, ValueError, KeyError) as error:
        if isinstance(error, InvalidFileError):
            raise
        raise InvalidFileError("工作表批注定义无效") from error


def _drawing_marker(anchor: ET.Element, name: str) -> tuple[int, int, int, int]:
    """功能：读取 DrawingML 锚点中的行列标记和偏移。

    使用方法：由图片读取器解析 ``from``、``to`` 元素时内部调用。
    参数：``anchor`` 为图片锚点元素；``name`` 为标记名称。
    返回：``(行索引, 列索引, 行偏移EMU, 列偏移EMU)``，行列索引均为0-based。
    异常：标记缺失或数值无效时抛出 ``InvalidFileError``。
    """
    marker = anchor.find(_tag(_SPREADSHEET_DRAWING_NS, name))
    if marker is None:
        raise InvalidFileError(f"图片锚点缺少 {name} 标记")
    try:
        row = int(marker.findtext(_tag(_SPREADSHEET_DRAWING_NS, "row"), "-1"))
        column = int(marker.findtext(_tag(_SPREADSHEET_DRAWING_NS, "col"), "-1"))
        row_offset = int(marker.findtext(_tag(_SPREADSHEET_DRAWING_NS, "rowOff"), "0"))
        column_offset = int(marker.findtext(_tag(_SPREADSHEET_DRAWING_NS, "colOff"), "0"))
    except ValueError as error:
        raise InvalidFileError("图片锚点行列标记不是整数") from error
    if row < 0 or column < 0:
        raise InvalidFileError("图片锚点行列标记不能为负数")
    return row, column, row_offset, column_offset


def _load_sheet_images(
    package: zipfile.ZipFile,
    member: str,
    root: ET.Element,
    worksheet: "Worksheet",
) -> None:
    """功能：读取工作表 DrawingML 中的嵌入 PNG/JPEG 图片。

    使用方法：由单张工作表读取流程自动调用，业务代码无需直接调用。
    参数：``package`` 为已打开的 XLSX ZIP 包；``member`` 为工作表部件路径；
    ``root`` 为工作表 XML 根元素；``worksheet`` 为接收图片的工作表。
    返回：``None``；图片对象按 DrawingML 顺序追加到 ``worksheet.images``。
    异常：关系、媒体、锚点或图片内容损坏时抛出 ``InvalidFileError``；不支持的
    图表和绝对锚点会被忽略，不影响其它单元格数据读取。
    """
    drawing_element = root.find(_tag(_MAIN_NS, "drawing"))
    if drawing_element is None:
        return
    relationship_id = drawing_element.get(_tag(_REL_NS, "id"))
    if not relationship_id:
        raise InvalidFileError("工作表 drawing 缺少关系编号")
    sheet_relationships = _sheet_relationships(package, member)
    drawing_relation = sheet_relationships.get(relationship_id)
    if drawing_relation is None or not drawing_relation[0].endswith("/drawing"):
        raise InvalidFileError("工作表 drawing 关系缺失或类型错误")
    drawing_member = drawing_relation[1]
    drawing_relationships = _sheet_relationships(package, drawing_member)
    try:
        drawing_root = _read_xml(package, drawing_member)
    except InvalidFileError:
        raise
    from ..image import Image, ImageFit, ImagePlacement

    image_relationships = {
        key: target
        for key, (relationship_type, target) in drawing_relationships.items()
        if relationship_type.endswith("/image")
    }
    if not image_relationships:
        return
    anchor_names = ("oneCellAnchor", "twoCellAnchor")
    for anchor_name in anchor_names:
        for anchor in drawing_root.findall(_tag(_SPREADSHEET_DRAWING_NS, anchor_name)):
            picture = anchor.find(_tag(_SPREADSHEET_DRAWING_NS, "pic"))
            if picture is None:
                continue
            blip = picture.find(
                f".//{_tag(_DRAWING_NS, 'blip')}"
            )
            if blip is None:
                continue
            embedded_id = blip.get(_tag(_REL_NS, "embed"))
            target = image_relationships.get(embedded_id or "")
            if target is None or target not in package.namelist() or not target.startswith("xl/media/"):
                raise InvalidFileError("图片媒体关系缺失或目标非法")
            try:
                from_row, from_column, from_row_offset, from_column_offset = _drawing_marker(anchor, "from")
                if anchor_name == "twoCellAnchor":
                    to_row, to_column, _to_row_offset, _to_column_offset = _drawing_marker(anchor, "to")
                    if to_row <= from_row or to_column <= from_column:
                        raise InvalidFileError("twoCellAnchor 的终点必须位于起点右下方")
                    bounds = (from_row, from_column, to_row - 1, to_column - 1)
                    placement = ImagePlacement.CELL
                else:
                    bounds = (from_row, from_column, from_row, from_column)
                    placement = ImagePlacement.FLOATING
                address = range_address(*bounds)
                image = Image(
                    worksheet,
                    package.read(target),
                    anchor=address,
                    name=Path(target).name,
                    placement=placement,
                    fit=ImageFit.STRETCH,
                )
                if anchor_name == "twoCellAnchor":
                    fit_value = anchor.get(_tag(_EXCELKIT_NS, "fit"))
                    if fit_value in {ImageFit.STRETCH, ImageFit.CONTAIN, ImageFit.COVER}:
                        image.fit = fit_value
                else:
                    ext = anchor.find(_tag(_SPREADSHEET_DRAWING_NS, "ext"))
                    if ext is not None:
                        width = int(ext.get("cx", "0")) // 9525
                        height = int(ext.get("cy", "0")) // 9525
                        if width > 0:
                            image.width = width
                        if height > 0:
                            image.height = height
                image.offset_x = int(round(from_column_offset / 9525))
                image.offset_y = int(round(from_row_offset / 9525))
                c_nv_pr = picture.find(
                    f".//{_tag(_SPREADSHEET_DRAWING_NS, 'cNvPr')}"
                )
                if c_nv_pr is not None and c_nv_pr.get("descr"):
                    image.alt_text = c_nv_pr.get("descr") or ""
                worksheet._images.append(image)
            except (OSError, TypeError, ValueError, KeyError) as error:
                if isinstance(error, InvalidFileError):
                    raise
                raise InvalidFileError("DrawingML 图片定义无效") from error


def _load_sheet_tables(
    package: zipfile.ZipFile,
    member: str,
    root: ET.Element,
    worksheet: "Worksheet",
) -> None:
    """功能：读取工作表关系所引用的全部 Excel 数据表定义。

    使用方法：单张工作表完成单元格和布局读取后由内部调用。
    参数：``package`` 为XLSX包；``member`` 为工作表包内路径；``root`` 为工作表
    XML根元素；``worksheet`` 为接收数据表对象的目标工作表。
    返回：``None``；没有 ``tableParts`` 时不做任何修改。
    异常：关系缺失、目标越界或数据表定义无效时抛出 :class:`InvalidFileError`。
    """
    table_parts = root.find(_tag(_MAIN_NS, "tableParts"))
    if table_parts is None:
        return
    relationships = _sheet_relationships(package, member)
    targets = {
        relationship_id: target
        for relationship_id, (relationship_type, target) in relationships.items()
        if relationship_type.endswith("/table")
    }
    if not targets:
        raise InvalidFileError("工作表数据表关系缺失")

    try:
        for table_part in table_parts.findall(_tag(_MAIN_NS, "tablePart")):
            relationship_id = table_part.get(_tag(_REL_NS, "id"))
            if not relationship_id or relationship_id not in targets:
                raise InvalidFileError("工作表数据表关系缺失")
            table_root = _read_xml(package, targets[relationship_id])
            name = table_root.get("displayName") or table_root.get("name")
            address = table_root.get("ref")
            if not name or not address:
                raise InvalidFileError("数据表名称或区域缺失")
            style_info = table_root.find(_tag(_MAIN_NS, "tableStyleInfo"))
            style = (
                style_info.get("name", "TableStyleMedium2")
                if style_info is not None else "TableStyleMedium2"
            )
            table = worksheet.add_table(
                address,
                name=name,
                style=style,
                has_header=table_root.get("headerRowCount", "1") != "0",
                show_row_stripes=(
                    style_info is not None
                    and style_info.get("showRowStripes", "0") in {"1", "true", "True"}
                ),
                show_column_stripes=(
                    style_info is not None
                    and style_info.get("showColumnStripes", "0") in {"1", "true", "True"}
                ),
            )
            if table_root.get("totalsRowShown", "0") in {"1", "true", "True"}:
                # 读取文件时 ref 已经包含实体汇总行，不能再次扩展边界。
                table._show_totals = True
                table._data_max_row = table.range.max_row - 1
            columns = table_root.find(_tag(_MAIN_NS, "tableColumns"))
            if columns is not None:
                names = table.columns
                for column in columns.findall(_tag(_MAIN_NS, "tableColumn")):
                    function = column.get("totalsRowFunction")
                    index = int(column.get("id", "0")) - 1
                    if function and 0 <= index < len(names):
                        table._totals[names[index]] = function
    except (TypeError, ValueError, KeyError) as error:
        if isinstance(error, InvalidFileError):
            raise
        raise InvalidFileError("工作表数据表定义无效") from error


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
        dxf_colors = _read_dxf_colors(package)
        properties = workbook_root.find(_tag(_MAIN_NS, "workbookPr"))
        date_1904 = properties is not None and properties.get("date1904", "0") in {
            "1", "true", "True"
        }
        workbook = workbook_class()
        _load_core_properties(package, workbook.properties)
        workbook_protection = workbook_root.find(_tag(_MAIN_NS, "workbookProtection"))
        if workbook_protection is not None:
            workbook.protection.enabled = True
            workbook.protection.password = workbook_protection.get("workbookPassword")
        sheets = workbook_root.find(_tag(_MAIN_NS, "sheets"))
        if sheets is None:
            raise InvalidFileError("workbook.xml 缺少 sheets 元素")
        for sheet in sheets.findall(_tag(_MAIN_NS, "sheet")):
            name = sheet.get("name")
            relationship_id = sheet.get(_tag(_REL_NS, "id"))
            if not name or not relationship_id or relationship_id not in targets:
                raise InvalidFileError("工作表名称或关系信息不完整")
            worksheet = workbook.add_sheet(name)
            state = sheet.get("state", "visible")
            if state == "hidden":
                worksheet.visibility = worksheet.HIDDEN
            elif state == "veryHidden":
                worksheet.visibility = worksheet.VERY_HIDDEN
            elif state != "visible":
                raise InvalidFileError(f"工作表可见状态无效：{state!r}")
            _load_sheet(
                package, targets[relationship_id], worksheet, shared_strings,
                styles, date_styles, date_1904, dxf_colors
            )
        _load_defined_names(workbook_root, workbook)
        return workbook


__all__ = []
