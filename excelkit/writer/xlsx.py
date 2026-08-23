"""基础 XLSX 文件包的唯一写出实现。"""

from __future__ import annotations

import math
import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterable, Optional, Sequence, Union

from ..address import cell_address, index_to_column, range_index
from ..core.page import header_footer_text
from ..style import DEFAULT_STYLE
from .styles import StyleRegistry

if TYPE_CHECKING:
    from ..core.table import Table
    from ..core.workbook import Workbook
    from ..core.worksheet import Worksheet

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("", _MAIN_NS)
ET.register_namespace("r", _REL_NS)

_ILLEGAL_XML_CHARACTERS = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]"
)
_EXCEL_ESCAPE_PATTERN = re.compile(r"_x[0-9A-Fa-f]{4}_")
_PAPER_SIZE_CODES = {"Letter": 1, "Legal": 5, "A3": 8, "A4": 9, "A5": 11}


def _qname(namespace: str, local_name: str) -> str:
    """功能：生成 ElementTree 使用的 XML 限定名称。

    使用方法：写出器内部调用 ``_qname(namespace, local_name)``。
    参数：``namespace`` 为命名空间 URI；``local_name`` 为元素或属性本地名称。
    返回：``{命名空间}本地名称`` 形式的字符串。
    """
    return f"{{{namespace}}}{local_name}"


def _xml_bytes(element: ET.Element) -> bytes:
    """功能：把 XML 根元素序列化为 UTF-8 字节串。

    使用方法：各底层 XML 生成函数在返回前调用。
    参数：``element`` 为 ``xml.etree.ElementTree.Element`` 根元素。
    返回：包含 XML 声明的 UTF-8 ``bytes``。
    """
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def _escape_text(value: str) -> str:
    """功能：按 Excel 规则转义 XML 1.0 不能表示的文本字符。

    使用方法：字符串单元格和公式写出前由内部调用。
    参数：``value`` 为原始字符串；已有 ``_xNNNN_`` 形式会保护其下划线。
    返回：可安全写入 SpreadsheetML 文本节点的字符串。
    """
    protected = _EXCEL_ESCAPE_PATTERN.sub(
        lambda match: "_x005F_" + match.group(0), value
    )
    return _ILLEGAL_XML_CHARACTERS.sub(
        lambda match: f"_x{ord(match.group(0)):04X}_", protected
    )


def _inline_string_cell(address: str, value: str) -> ET.Element:
    """功能：创建内联字符串类型的 SpreadsheetML 单元格元素。

    使用方法：由 :func:`cell_xml` 处理文本及日期时间时调用。
    参数：``address`` 为规范化 A1 地址；``value`` 为待写出的字符串。
    返回：类型为 ``inlineStr`` 的 ``c`` 元素。
    """
    cell = ET.Element(_qname(_MAIN_NS, "c"), {"r": address, "t": "inlineStr"})
    inline_string = ET.SubElement(cell, _qname(_MAIN_NS, "is"))
    text = ET.SubElement(inline_string, _qname(_MAIN_NS, "t"))
    if value[:1].isspace() or value[-1:].isspace():
        text.set(_XML_SPACE, "preserve")
    text.text = _escape_text(value)
    return cell


def content_types(sheet_count: int, table_count: int = 0) -> bytes:
    """功能：根据工作表数量生成 XLSX 内容类型声明。

    使用方法：``content_types(len(workbook.sheets))``；由 :class:`XlsxWriter`
    写入 ``[Content_Types].xml``。
    参数：``sheet_count`` 为大于等于 1 的整数工作表数量；``table_count`` 为非负
    Excel数据表数量，两者都属于计数而非索引。
    返回：UTF-8 XML 字节串。
    异常：参数不是正整数或是布尔值时抛出 ``ValueError``。
    """
    if (
        isinstance(sheet_count, bool)
        or not isinstance(sheet_count, int)
        or sheet_count < 1
    ):
        raise ValueError("工作表数量必须是大于等于 1 的整数")
    if (
        isinstance(table_count, bool)
        or not isinstance(table_count, int)
        or table_count < 0
    ):
        raise ValueError("数据表数量必须是非负整数")

    root = ET.Element(_qname(_CONTENT_TYPES_NS, "Types"))
    ET.SubElement(
        root,
        _qname(_CONTENT_TYPES_NS, "Default"),
        {
            "Extension": "rels",
            "ContentType": "application/vnd.openxmlformats-package.relationships+xml",
        },
    )
    ET.SubElement(
        root,
        _qname(_CONTENT_TYPES_NS, "Default"),
        {"Extension": "xml", "ContentType": "application/xml"},
    )
    ET.SubElement(
        root,
        _qname(_CONTENT_TYPES_NS, "Override"),
        {
            "PartName": "/xl/workbook.xml",
            "ContentType": (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            ),
        },
    )
    for sheet_index in range(sheet_count):
        ET.SubElement(
            root,
            _qname(_CONTENT_TYPES_NS, "Override"),
            {
                "PartName": f"/xl/worksheets/sheet{sheet_index + 1}.xml",
                "ContentType": (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                ),
            },
        )
    ET.SubElement(
        root,
        _qname(_CONTENT_TYPES_NS, "Override"),
        {
            "PartName": "/xl/styles.xml",
            "ContentType": (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
            ),
        },
    )
    for table_id in range(1, table_count + 1):
        ET.SubElement(
            root,
            _qname(_CONTENT_TYPES_NS, "Override"),
            {
                "PartName": f"/xl/tables/table{table_id}.xml",
                "ContentType": (
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.table+xml"
                ),
            },
        )
    return _xml_bytes(root)


def workbook_xml(
    sheets: Sequence["Worksheet"], named_ranges: Sequence[object] = ()
) -> bytes:
    """功能：生成包含工作表名称、顺序和关系编号的工作簿 XML。

    使用方法：``workbook_xml(workbook.sheets)``。
    参数：``sheets`` 为按创建顺序排列的工作表序列；``named_ranges`` 为工作簿级
    命名区域序列。
    返回：可写入 ``xl/workbook.xml`` 的 UTF-8 XML 字节串。
    """
    root = ET.Element(_qname(_MAIN_NS, "workbook"))
    sheet_elements = ET.SubElement(root, _qname(_MAIN_NS, "sheets"))
    for sheet_index, worksheet in enumerate(sheets):
        ET.SubElement(
            sheet_elements,
            _qname(_MAIN_NS, "sheet"),
            {
                "name": worksheet.name,
                "sheetId": str(sheet_index + 1),
                _qname(_REL_NS, "id"): f"rId{sheet_index + 1}",
            },
        )
    defined_names = []
    for sheet_index, worksheet in enumerate(sheets):
        sheet_name = worksheet.name.replace("'", "''")
        prefix = f"'{sheet_name}'!"
        if worksheet.page.print_area is not None:
            min_row, min_column, max_row, max_column = range_index(
                worksheet.page.print_area
            )
            area = (
                f"${index_to_column(min_column)}${min_row + 1}:"
                f"${index_to_column(max_column)}${max_row + 1}"
            )
            defined_names.append(
                ("_xlnm.Print_Area", sheet_index, prefix + area)
            )
        titles = []
        if worksheet.page.repeat_rows is not None:
            start, end = worksheet.page.repeat_rows
            titles.append(f"${start + 1}:${end + 1}")
        if worksheet.page.repeat_columns is not None:
            start, end = worksheet.page.repeat_columns
            titles.append(
                f"${index_to_column(start)}:${index_to_column(end)}"
            )
        if titles:
            defined_names.append(
                (
                    "_xlnm.Print_Titles",
                    sheet_index,
                    ",".join(prefix + title for title in titles),
                )
            )
    for named_range in named_ranges:
        area = named_range.range
        sheet_name = named_range.worksheet.name.replace("'", "''")
        reference = (
            f"'{sheet_name}'!"
            f"${index_to_column(area.min_column)}${area.min_row + 1}:"
            f"${index_to_column(area.max_column)}${area.max_row + 1}"
        )
        defined_names.append((named_range.name, None, reference))
    if defined_names:
        container = ET.SubElement(root, _qname(_MAIN_NS, "definedNames"))
        for name, sheet_index, value in defined_names:
            attributes = {"name": name}
            if sheet_index is not None:
                attributes["localSheetId"] = str(sheet_index)
            ET.SubElement(
                container, _qname(_MAIN_NS, "definedName"), attributes
            ).text = value
    ET.SubElement(
        root,
        _qname(_MAIN_NS, "calcPr"),
        {"calcMode": "auto", "fullCalcOnLoad": "1", "forceFullCalc": "1"},
    )
    return _xml_bytes(root)


def workbook_rels(sheets: Sequence["Worksheet"]) -> bytes:
    """功能：生成工作簿到各工作表 XML 部件的关系清单。

    使用方法：``workbook_rels(workbook.sheets)``。
    参数：``sheets`` 为按创建顺序排列的工作表序列。
    返回：可写入 ``xl/_rels/workbook.xml.rels`` 的 UTF-8 XML 字节串。
    """
    root = ET.Element(_qname(_PACKAGE_REL_NS, "Relationships"))
    relationship_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    )
    for sheet_index, _worksheet in enumerate(sheets):
        ET.SubElement(
            root,
            _qname(_PACKAGE_REL_NS, "Relationship"),
            {
                "Id": f"rId{sheet_index + 1}",
                "Type": relationship_type,
                "Target": f"worksheets/sheet{sheet_index + 1}.xml",
            },
        )
    ET.SubElement(
        root,
        _qname(_PACKAGE_REL_NS, "Relationship"),
        {
            "Id": f"rId{len(sheets) + 1}",
            "Type": (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
            ),
            "Target": "styles.xml",
        },
    )
    return _xml_bytes(root)


def cell_xml(address: str, value: Any) -> Optional[ET.Element]:
    """功能：把普通单元格值转换为 SpreadsheetML 单元格元素。

    使用方法：由 :func:`sheet_xml` 调用 ``cell_xml(address, value)``。
    参数：``address`` 为规范化 A1 地址；``value`` 为普通 Python 值。``None``
    不生成元素，布尔值和有限整数、浮点数按原生类型写出，日期时间使用 Open XML
    ISO 日期类型，
    其他对象使用 ``str(value)`` 转为内联文本。
    返回：``c`` 元素；``value is None`` 时返回 ``None``。
    """
    if value is None:
        return None
    if isinstance(value, str):
        return _inline_string_cell(address, value)
    if isinstance(value, (datetime, date)):
        cell = ET.Element(_qname(_MAIN_NS, "c"), {"r": address, "t": "d"})
        ET.SubElement(cell, _qname(_MAIN_NS, "v")).text = value.isoformat()
        return cell
    if isinstance(value, bool):
        cell = ET.Element(_qname(_MAIN_NS, "c"), {"r": address, "t": "b"})
        ET.SubElement(cell, _qname(_MAIN_NS, "v")).text = "1" if value else "0"
        return cell
    if isinstance(value, int):
        cell = ET.Element(_qname(_MAIN_NS, "c"), {"r": address})
        ET.SubElement(cell, _qname(_MAIN_NS, "v")).text = str(value)
        return cell
    if isinstance(value, float) and math.isfinite(value):
        cell = ET.Element(_qname(_MAIN_NS, "c"), {"r": address})
        ET.SubElement(cell, _qname(_MAIN_NS, "v")).text = repr(value)
        return cell
    return _inline_string_cell(address, str(value))


def formula_xml(address: str, formula: str, cached_value: Any = None) -> ET.Element:
    """功能：把公式转换为 SpreadsheetML 公式单元格元素。

    使用方法：由 :func:`sheet_xml` 调用 ``formula_xml("C2", "=SUM(A2:B2)", 95)``。
    参数：``address`` 为规范化 A1 地址；``formula`` 为包含表达式的公式字符串，
    可以带或不带前导 ``=``。
    参数：``cached_value`` 为 Excel/WPS 最近一次保存或 ``Workbook.calculate()``
    生成的计算结果；为 ``None`` 时不写入 ``v`` 元素。
    返回：包含 ``f`` 子元素以及可选 ``v`` 缓存结果的 ``c`` 元素。
    异常：公式不是非空字符串或没有表达式时抛出 ``TypeError``。
    """
    if not isinstance(formula, str):
        raise TypeError("公式必须是非空字符串")
    expression = formula.strip()
    if expression.startswith("="):
        expression = expression[1:].strip()
    if not expression:
        raise TypeError("公式必须包含表达式")
    cell = ET.Element(_qname(_MAIN_NS, "c"), {"r": address})
    ET.SubElement(cell, _qname(_MAIN_NS, "f")).text = _escape_text(expression)
    if cached_value is not None:
        if isinstance(cached_value, str):
            cell.set("t", "str")
            cached_text = cached_value
        elif isinstance(cached_value, (datetime, date)):
            cell.set("t", "d")
            cached_text = cached_value.isoformat()
        elif isinstance(cached_value, bool):
            cell.set("t", "b")
            cached_text = "1" if cached_value else "0"
        elif isinstance(cached_value, int):
            cached_text = str(cached_value)
        elif isinstance(cached_value, float) and math.isfinite(cached_value):
            cached_text = repr(cached_value)
        else:
            cell.set("t", "str")
            cached_text = str(cached_value)
        ET.SubElement(cell, _qname(_MAIN_NS, "v")).text = _escape_text(cached_text)
    return cell


def _table_column_names(table: "Table") -> list[str]:
    """功能：从数据表首行生成非空且大小写不敏感唯一的列名称。

    使用方法：由 :func:`table_xml` 写出 ``tableColumns`` 时调用。
    参数：``table`` 为待写出的数据表。
    返回：与表格列数一致的名称列表；空值和重复名称自动生成稳定后缀。
    """
    area = table.range
    names: list[str] = []
    used: set[str] = set()
    for offset, column in enumerate(
        range(area.min_column, area.max_column + 1), 1
    ):
        value = (
            table.worksheet._values.get(area.min_row, column)
            if table.has_header else None
        )
        base = str(value) if value not in (None, "") else f"Column{offset}"
        candidate = base
        suffix = 2
        while candidate.casefold() in used:
            candidate = f"{base}_{suffix}"
            suffix += 1
        used.add(candidate.casefold())
        names.append(candidate)
    return names


def table_xml(table: "Table", table_id: int) -> bytes:
    """功能：生成一个标准 SpreadsheetML 数据表定义部件。

    使用方法：XLSX打包器为每个工作簿数据表调用一次。
    参数：``table`` 为数据表对象；``table_id`` 为从1开始的包内唯一整数。
    返回：可写入 ``xl/tables/tableN.xml`` 的UTF-8 XML字节串。
    """
    area = table.range
    attributes = {
        "id": str(table_id),
        "name": table.name,
        "displayName": table.name,
        "ref": area.address,
        "totalsRowShown": "0",
    }
    if not table.has_header:
        attributes["headerRowCount"] = "0"
    root = ET.Element(_qname(_MAIN_NS, "table"), attributes)
    if table.has_header:
        ET.SubElement(root, _qname(_MAIN_NS, "autoFilter"), {"ref": area.address})
    names = _table_column_names(table)
    columns = ET.SubElement(
        root, _qname(_MAIN_NS, "tableColumns"), {"count": str(len(names))}
    )
    for column_id, name in enumerate(names, 1):
        ET.SubElement(
            columns,
            _qname(_MAIN_NS, "tableColumn"),
            {"id": str(column_id), "name": _escape_text(name)},
        )
    ET.SubElement(
        root,
        _qname(_MAIN_NS, "tableStyleInfo"),
        {
            "name": table.style,
            "showFirstColumn": "0",
            "showLastColumn": "0",
            "showRowStripes": "1" if table.show_row_stripes else "0",
            "showColumnStripes": "1" if table.show_column_stripes else "0",
        },
    )
    return _xml_bytes(root)


def worksheet_rels(table_ids: Sequence[int]) -> bytes:
    """功能：生成单张工作表到其数据表部件的关系清单。

    使用方法：工作表包含数据表时由XLSX打包器调用。
    参数：``table_ids`` 为当前工作表数据表对应的全局1-based编号。
    返回：工作表 ``.rels`` XML字节串。
    """
    root = ET.Element(_qname(_PACKAGE_REL_NS, "Relationships"))
    for relationship_index, table_id in enumerate(table_ids, 1):
        ET.SubElement(
            root,
            _qname(_PACKAGE_REL_NS, "Relationship"),
            {
                "Id": f"rId{relationship_index}",
                "Type": (
                    "http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships/table"
                ),
                "Target": f"../tables/table{table_id}.xml",
            },
        )
    return _xml_bytes(root)


def sheet_xml(
    sheet: "Worksheet",
    style_registry: Optional[StyleRegistry] = None,
    table_ids: Sequence[int] = (),
) -> bytes:
    """功能：生成单张工作表的完整 SpreadsheetML XML。

    使用方法：由 :class:`XlsxWriter` 对工作簿中的每张工作表调用。
    参数：``sheet`` 为待写出的工作表；``style_registry`` 为工作簿共享样式注册器，
    省略时为当前单表临时创建；``table_ids`` 为本表数据表对应的全局1-based编号。
    普通值、公式和样式按0-based行列索引合并排序。
    返回：包含精确数据边界和 ``sheetData`` 的 UTF-8 XML 字节串。
    """
    root = ET.Element(_qname(_MAIN_NS, "worksheet"))
    page = sheet.page
    fit_mode = page.scale is None and (
        page._fit_width is not None or page._fit_height is not None
    )
    if sheet.color is not None or fit_mode:
        properties = ET.SubElement(root, _qname(_MAIN_NS, "sheetPr"))
        if sheet.color is not None:
            ET.SubElement(
                properties,
                _qname(_MAIN_NS, "tabColor"),
                {"rgb": sheet.color},
            )
        if fit_mode:
            ET.SubElement(
                properties, _qname(_MAIN_NS, "pageSetUpPr"), {"fitToPage": "1"}
            )
    if style_registry is None:
        style_registry = StyleRegistry([sheet])
    values = dict(sheet._values.items())
    coordinates = set(values) | set(sheet._formulas) | set(sheet._styles)
    if coordinates:
        rows = [row for row, _column in coordinates]
        columns = [column for _row, column in coordinates]
        start = cell_address(min(rows), min(columns))
        end = cell_address(max(rows), max(columns))
        dimension = start if start == end else f"{start}:{end}"
    else:
        dimension = "A1"
    ET.SubElement(root, _qname(_MAIN_NS, "dimension"), {"ref": dimension})

    sheet_views = ET.SubElement(root, _qname(_MAIN_NS, "sheetViews"))
    view_attributes = {"workbookViewId": "0"}
    if not sheet.show_gridlines:
        view_attributes["showGridLines"] = "0"
    sheet_view = ET.SubElement(
        sheet_views, _qname(_MAIN_NS, "sheetView"), view_attributes
    )
    if sheet.freeze_panes is not None:
        freeze_row, freeze_column = range_index(
            f"{sheet.freeze_panes}:{sheet.freeze_panes}"
        )[:2]
        pane_attributes = {"state": "frozen", "topLeftCell": sheet.freeze_panes}
        if freeze_column:
            pane_attributes["xSplit"] = str(freeze_column)
        if freeze_row:
            pane_attributes["ySplit"] = str(freeze_row)
        pane = (
            "bottomRight" if freeze_row and freeze_column
            else "bottomLeft" if freeze_row else "topRight"
        )
        pane_attributes["activePane"] = pane
        ET.SubElement(sheet_view, _qname(_MAIN_NS, "pane"), pane_attributes)
        ET.SubElement(
            sheet_view,
            _qname(_MAIN_NS, "selection"),
            {"pane": pane, "activeCell": sheet.freeze_panes, "sqref": sheet.freeze_panes},
        )

    custom_columns = [
        dimension for _index, dimension in sorted(sheet._columns.items())
        if not dimension._is_default()
    ]
    if custom_columns:
        cols = ET.SubElement(root, _qname(_MAIN_NS, "cols"))
        for column in custom_columns:
            attributes = {
                "min": str(column.index + 1),
                "max": str(column.index + 1),
            }
            if column.width is not None:
                attributes.update({"width": str(column.width), "customWidth": "1"})
            if column.hidden:
                attributes["hidden"] = "1"
            ET.SubElement(cols, _qname(_MAIN_NS, "col"), attributes)

    sheet_data = ET.SubElement(root, _qname(_MAIN_NS, "sheetData"))
    row_indexes = sorted(
        {row for row, _column in coordinates}
        | {
            index for index, dimension in sheet._rows.items()
            if not dimension._is_default()
        }
    )
    for row in row_indexes:
        row_attributes = {"r": str(row + 1)}
        row_dimension = sheet._rows.get(row)
        if row_dimension is not None:
            if row_dimension.height is not None:
                row_attributes.update(
                    {"ht": str(row_dimension.height), "customHeight": "1"}
                )
            if row_dimension.hidden:
                row_attributes["hidden"] = "1"
        row_element = ET.SubElement(
            sheet_data, _qname(_MAIN_NS, "row"), row_attributes
        )
        columns = sorted(column for item_row, column in coordinates if item_row == row)
        for column in columns:
            address = cell_address(row, column)
            style = sheet._styles.get((row, column), DEFAULT_STYLE)
            style_id = style_registry.style_id(style)
            formula = sheet._formulas.get((row, column))
            if formula is not None:
                element = formula_xml(
                    address, formula, sheet._formula_values.get((row, column))
                )
            else:
                element = cell_xml(address, values.get((row, column)))
            if element is None and style_id:
                element = ET.Element(_qname(_MAIN_NS, "c"), {"r": address})
            if element is not None:
                if style_id:
                    element.set("s", str(style_id))
                row_element.append(element)
    if sheet.auto_filter_range is not None:
        ET.SubElement(
            root, _qname(_MAIN_NS, "autoFilter"), {"ref": sheet.auto_filter_range}
        )
    if sheet._merged_ranges:
        merge_cells = ET.SubElement(
            root,
            _qname(_MAIN_NS, "mergeCells"),
            {"count": str(len(sheet._merged_ranges))},
        )
        for min_row, min_column, max_row, max_column in sheet._merged_ranges:
            ET.SubElement(
                merge_cells,
                _qname(_MAIN_NS, "mergeCell"),
                {
                    "ref": (
                        f"{cell_address(min_row, min_column)}:"
                        f"{cell_address(max_row, max_column)}"
                    )
                },
            )

    print_options = {
        "horizontalCentered": "1" if page.center_horizontal else "0",
        "verticalCentered": "1" if page.center_vertical else "0",
        "gridLines": "1" if page.print_gridlines else "0",
        "headings": "1" if page.print_headings else "0",
    }
    ET.SubElement(root, _qname(_MAIN_NS, "printOptions"), print_options)
    margins = page.margins
    ET.SubElement(
        root,
        _qname(_MAIN_NS, "pageMargins"),
        {
            name: str(getattr(margins, name) / 2.54)
            for name in ("left", "right", "top", "bottom", "header", "footer")
        },
    )
    setup_attributes = {
        "orientation": page.orientation,
        "paperSize": str(_PAPER_SIZE_CODES[page.paper_size]),
        "pageOrder": (
            "downThenOver" if page.print_order == "down_then_over" else "overThenDown"
        ),
        "blackAndWhite": "1" if page.black_and_white else "0",
        "draft": "1" if page.draft else "0",
    }
    if page.scale is not None:
        setup_attributes["scale"] = str(page.scale)
    else:
        setup_attributes["fitToWidth"] = str(page._fit_width or 0)
        setup_attributes["fitToHeight"] = str(page._fit_height or 0)
    if page.first_page_number is not None:
        setup_attributes["firstPageNumber"] = str(page.first_page_number)
        setup_attributes["useFirstPageNumber"] = "1"
    ET.SubElement(root, _qname(_MAIN_NS, "pageSetup"), setup_attributes)

    header_text = header_footer_text(page.header)
    footer_text = header_footer_text(page.footer)
    if header_text or footer_text:
        header_footer = ET.SubElement(root, _qname(_MAIN_NS, "headerFooter"))
        if header_text:
            ET.SubElement(
                header_footer, _qname(_MAIN_NS, "oddHeader")
            ).text = header_text
        if footer_text:
            ET.SubElement(
                header_footer, _qname(_MAIN_NS, "oddFooter")
            ).text = footer_text
    if table_ids:
        table_parts = ET.SubElement(
            root,
            _qname(_MAIN_NS, "tableParts"),
            {"count": str(len(table_ids))},
        )
        for relationship_index, _table_id in enumerate(table_ids, 1):
            ET.SubElement(
                table_parts,
                _qname(_MAIN_NS, "tablePart"),
                {_qname(_REL_NS, "id"): f"rId{relationship_index}"},
            )
    return _xml_bytes(root)


def _root_rels() -> bytes:
    """功能：生成 XLSX 根目录到工作簿部件的关系声明。

    使用方法：由 :class:`XlsxWriter` 写入 ``_rels/.rels``。
    参数：无。
    返回：UTF-8 XML 字节串。
    """
    root = ET.Element(_qname(_PACKAGE_REL_NS, "Relationships"))
    ET.SubElement(
        root,
        _qname(_PACKAGE_REL_NS, "Relationship"),
        {
            "Id": "rId1",
            "Type": (
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                "officeDocument"
            ),
            "Target": "xl/workbook.xml",
        },
    )
    return _xml_bytes(root)


class XlsxWriter:
    """把 :class:`Workbook` 写成基础且符合标准的 XLSX 文件包。"""

    def __init__(self, workbook: "Workbook") -> None:
        """功能：创建绑定到指定工作簿的 XLSX 写出器。

        使用方法：``writer = XlsxWriter(workbook)``；业务代码优先使用
        ``workbook.save(filename)``。
        参数：``workbook`` 为待写出的 :class:`Workbook`。
        返回：无。
        """
        self.workbook = workbook

    def write(self, filename: Union[str, os.PathLike]) -> None:
        """功能：以临时文件和原子替换方式写出 XLSX 文件包。

        使用方法：``XlsxWriter(workbook).write("demo.xlsx")``。
        参数：``filename`` 为字符串或 ``os.PathLike`` 路径；父目录必须已经存在。
        返回：``None``。
        异常：参数类型错误时抛出 ``TypeError``；路径没有文件名时抛出
        ``ValueError``；文件系统失败时透传相应异常，且已有目标文件保持不变。
        """
        if not isinstance(filename, (str, os.PathLike)):
            raise TypeError("filename 必须是字符串或 PathLike 对象")
        target = Path(filename)
        if not target.name:
            raise ValueError("filename 必须指定文件名")
        if not self.workbook.sheets:
            self.workbook.active

        temporary_name = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                prefix=f".{target.name}.",
                suffix=".tmp",
                dir=target.parent,
                delete=False,
            ) as temporary:
                temporary_name = temporary.name
            with zipfile.ZipFile(
                temporary_name, "w", compression=zipfile.ZIP_DEFLATED
            ) as package:
                sheets = self.workbook.sheets
                style_registry = StyleRegistry(sheets)
                table_entries = []
                sheet_table_ids = {}
                next_table_id = 1
                for sheet_index, worksheet in enumerate(sheets):
                    current_ids = []
                    for table in worksheet.tables:
                        current_ids.append(next_table_id)
                        table_entries.append((next_table_id, table))
                        next_table_id += 1
                    sheet_table_ids[sheet_index] = tuple(current_ids)
                package.writestr(
                    "[Content_Types].xml",
                    content_types(len(sheets), len(table_entries)),
                )
                package.writestr("_rels/.rels", _root_rels())
                package.writestr(
                    "xl/workbook.xml",
                    workbook_xml(sheets, self.workbook.named_ranges),
                )
                package.writestr("xl/_rels/workbook.xml.rels", workbook_rels(sheets))
                package.writestr("xl/styles.xml", style_registry.xml())
                for sheet_index, worksheet in enumerate(sheets):
                    table_ids = sheet_table_ids[sheet_index]
                    package.writestr(
                        f"xl/worksheets/sheet{sheet_index + 1}.xml",
                        sheet_xml(worksheet, style_registry, table_ids),
                    )
                    if table_ids:
                        package.writestr(
                            f"xl/worksheets/_rels/sheet{sheet_index + 1}.xml.rels",
                            worksheet_rels(table_ids),
                        )
                for table_id, table in table_entries:
                    package.writestr(
                        f"xl/tables/table{table_id}.xml",
                        table_xml(table, table_id),
                    )
            os.replace(temporary_name, target)
            temporary_name = None
        finally:
            if temporary_name is not None:
                try:
                    os.unlink(temporary_name)
                except FileNotFoundError:
                    pass


__all__ = [
    "XlsxWriter",
    "content_types",
    "workbook_xml",
    "workbook_rels",
    "cell_xml",
    "formula_xml",
    "table_xml",
    "worksheet_rels",
    "sheet_xml",
]
