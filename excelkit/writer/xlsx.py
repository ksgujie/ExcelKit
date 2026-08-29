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

from ..address import cell_address, cell_index, index_to_column, range_index
from ..core.page import header_footer_text
from ..properties import WorkbookProperties
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
_CORE_PROPERTIES_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_DCTERMS_NS = "http://purl.org/dc/terms/"
_DC_NS = "http://purl.org/dc/elements/1.1/"
_XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_VML_NS = "urn:schemas-microsoft-com:vml"
_OFFICE_NS = "urn:schemas-microsoft-com:office:office"
_EXCEL_NS = "urn:schemas-microsoft-com:office:excel"

ET.register_namespace("", _MAIN_NS)
ET.register_namespace("r", _REL_NS)
ET.register_namespace("cp", _CORE_PROPERTIES_NS)
ET.register_namespace("dcterms", _DCTERMS_NS)
ET.register_namespace("dc", _DC_NS)
ET.register_namespace("a", _DRAWING_NS)
ET.register_namespace("xdr", _SPREADSHEET_DRAWING_NS)
ET.register_namespace("c", _CHART_NS)
ET.register_namespace("v", _VML_NS)
ET.register_namespace("o", _OFFICE_NS)
ET.register_namespace("x", _EXCEL_NS)
ET.register_namespace("xsi", _XSI_NS)

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


def content_types(
    sheet_count: int,
    table_count: int = 0,
    note_count: int = 0,
    drawing_count: int = 0,
    chart_count: int = 0,
    image_extensions: Sequence[str] = (),
) -> bytes:
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
    for name, value in (("table_count", table_count), ("note_count", note_count),
                        ("drawing_count", drawing_count), ("chart_count", chart_count)):
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise ValueError(f"{name} 必须是非负整数")

    root = ET.Element(_qname(_CONTENT_TYPES_NS, "Types"))
    ET.SubElement(
        root,
        _qname(_CONTENT_TYPES_NS, "Default"),
        {
            "Extension": "rels",
            "ContentType": "application/vnd.openxmlformats-package.relationships+xml",
        },
    )
    for extension in sorted(set(image_extensions)):
        if extension not in {"png", "jpeg"}:
            raise ValueError("图片扩展名只能是 png 或 jpeg")
        ET.SubElement(
            root,
            _qname(_CONTENT_TYPES_NS, "Default"),
            {"Extension": extension, "ContentType": f"image/{extension}"},
        )
    if note_count:
        ET.SubElement(
            root,
            _qname(_CONTENT_TYPES_NS, "Default"),
            {"Extension": "vml", "ContentType": "application/vnd.openxmlformats-officedocument.vmlDrawing"},
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
    ET.SubElement(
        root,
        _qname(_CONTENT_TYPES_NS, "Override"),
        {
            "PartName": "/docProps/core.xml",
            "ContentType": "application/vnd.openxmlformats-package.core-properties+xml",
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
    for note_id in range(1, note_count + 1):
        ET.SubElement(
            root,
            _qname(_CONTENT_TYPES_NS, "Override"),
            {
                "PartName": f"/xl/comments{note_id}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            },
        )
    for drawing_id in range(1, drawing_count + 1):
        ET.SubElement(
            root,
            _qname(_CONTENT_TYPES_NS, "Override"),
            {
                "PartName": f"/xl/drawings/drawing{drawing_id}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
            },
        )
    for chart_id in range(1, chart_count + 1):
        ET.SubElement(
            root,
            _qname(_CONTENT_TYPES_NS, "Override"),
            {
                "PartName": f"/xl/charts/chart{chart_id}.xml",
                "ContentType": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            },
        )
    return _xml_bytes(root)


def workbook_xml(
    sheets: Sequence["Worksheet"], named_ranges: Sequence[object] = (), protection=None
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
        attributes = {
            "name": worksheet.name,
            "sheetId": str(sheet_index + 1),
            _qname(_REL_NS, "id"): f"rId{sheet_index + 1}",
        }
        if worksheet.visibility == worksheet.HIDDEN:
            attributes["state"] = "hidden"
        elif worksheet.visibility == worksheet.VERY_HIDDEN:
            attributes["state"] = "veryHidden"
        ET.SubElement(sheet_elements, _qname(_MAIN_NS, "sheet"), attributes)
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
    if protection is not None and protection.enabled:
        attributes = {}
        if protection.password:
            attributes["workbookPassword"] = str(protection.password)
        ET.SubElement(root, _qname(_MAIN_NS, "workbookProtection"), attributes)
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
        "totalsRowShown": "1" if table.show_totals else "0",
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
        element = ET.SubElement(
            columns,
            _qname(_MAIN_NS, "tableColumn"),
            {"id": str(column_id), "name": _escape_text(name)},
        )
        function = table.totals.get(name)
        if function:
            element.set("totalsRowFunction", function)
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


def comments_xml(notes: Sequence[tuple[str, object]]) -> bytes:
    """功能：生成一个工作表传统批注 comments XML 部件。

    使用方法：XLSX 打包器为包含 ``Cell.note`` 的工作表调用。
    参数：``notes`` 为按 A1 地址排序的 ``(地址, Note)`` 序列。
    返回：可写入 ``xl/commentsN.xml`` 的 UTF-8 XML 字节串。
    """
    root = ET.Element(_qname(_MAIN_NS, "comments"))
    authors: list[str] = []
    for _address, note in notes:
        if note.author not in authors:
            authors.append(note.author)
    author_list = ET.SubElement(root, _qname(_MAIN_NS, "authors"))
    for author in authors:
        ET.SubElement(author_list, _qname(_MAIN_NS, "author")).text = _escape_text(author)
    comment_list = ET.SubElement(root, _qname(_MAIN_NS, "commentList"))
    for address, note in notes:
        comment = ET.SubElement(
            comment_list,
            _qname(_MAIN_NS, "comment"),
            {"ref": address, "authorId": str(authors.index(note.author))},
        )
        text = ET.SubElement(comment, _qname(_MAIN_NS, "text"))
        run = ET.SubElement(text, _qname(_MAIN_NS, "r"))
        value = ET.SubElement(run, _qname(_MAIN_NS, "t"))
        if note.text[:1].isspace() or note.text[-1:].isspace():
            value.set(_XML_SPACE, "preserve")
        value.text = _escape_text(note.text)
    return _xml_bytes(root)


def vml_comments_xml(notes: Sequence[tuple[str, object]]) -> bytes:
    """功能：生成传统批注显示框所需的 VML 绘图部件。

    使用方法：与 :func:`comments_xml` 成对写入 XLSX 包。
    参数：``notes`` 为按 A1 地址排序的批注序列。
    返回：可写入 ``xl/drawings/vmlDrawingN.vml`` 的 XML 字节串。
    """
    root = ET.Element("xml")
    shape_type = ET.SubElement(
        root,
        _qname(_VML_NS, "shapetype"),
        {"id": "_x0000_t202", "coordsize": "21600,21600", _qname(_OFFICE_NS, "spt"): "202", "path": "m,l,21600r21600,l21600,xe"},
    )
    ET.SubElement(shape_type, _qname(_VML_NS, "stroke"), {"joinstyle": "miter"})
    ET.SubElement(shape_type, _qname(_VML_NS, "path"), {"gradientshapeok": "t", _qname(_OFFICE_NS, "connecttype"): "rect"})
    for index, (address, _note) in enumerate(notes, 1):
        row, column = cell_index(address)
        shape = ET.SubElement(
            root,
            _qname(_VML_NS, "shape"),
            {
                "id": f"_x0000_s{1024 + index}",
                "type": "#_x0000_t202",
                "style": "position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden",
                "fillcolor": "#ffffe1",
                _qname(_OFFICE_NS, "insetmode"): "auto",
            },
        )
        ET.SubElement(shape, _qname(_VML_NS, "fill"), {"color2": "#ffffe1"})
        ET.SubElement(shape, _qname(_VML_NS, "shadow"), {"on": "t", "color": "black", "obscured": "t"})
        ET.SubElement(shape, _qname(_VML_NS, "path"), {_qname(_OFFICE_NS, "connecttype"): "none"})
        textbox = ET.SubElement(shape, _qname(_VML_NS, "textbox"), {"style": "mso-direction-alt:auto"})
        ET.SubElement(textbox, "div", {"style": "text-align:left"})
        client_data = ET.SubElement(shape, _qname(_EXCEL_NS, "ClientData"), {"ObjectType": "Note"})
        ET.SubElement(client_data, _qname(_EXCEL_NS, "MoveWithCells"))
        ET.SubElement(client_data, _qname(_EXCEL_NS, "SizeWithCells"))
        ET.SubElement(client_data, _qname(_EXCEL_NS, "AutoFill")).text = "False"
        ET.SubElement(client_data, _qname(_EXCEL_NS, "Row")).text = str(row)
        ET.SubElement(client_data, _qname(_EXCEL_NS, "Column")).text = str(column)
    return _xml_bytes(root)


def _chart_reference(worksheet: "Worksheet", address: str) -> str:
    """功能：把图表区域地址转换为带工作表名称的绝对公式引用。"""
    area = worksheet.range(address)
    name = worksheet.name.replace("'", "''")
    return (
        f"'{name}'!${index_to_column(area.min_column)}${area.min_row + 1}:"
        f"${index_to_column(area.max_column)}${area.max_row + 1}"
    )


def chart_xml(chart: object) -> bytes:
    """功能：生成柱状、条形、折线或饼图的 ChartML 部件。

    使用方法：XLSX 打包器为每个 ``Chart`` 调用。
    参数：``chart`` 为绑定工作表、类型和系列的图表对象。
    返回：可写入 ``xl/charts/chartN.xml`` 的 XML 字节串。
    异常：图表没有数据系列时抛出 ``ValueError``。
    """
    if not chart.series:
        raise ValueError("图表至少需要一个数据系列")
    root = ET.Element(_qname(_CHART_NS, "chartSpace"))
    chart_element = ET.SubElement(root, _qname(_CHART_NS, "chart"))
    if chart.title:
        title = ET.SubElement(chart_element, _qname(_CHART_NS, "title"))
        tx = ET.SubElement(title, _qname(_CHART_NS, "tx"))
        rich = ET.SubElement(tx, _qname(_CHART_NS, "rich"))
        ET.SubElement(rich, _qname(_DRAWING_NS, "bodyPr"))
        ET.SubElement(rich, _qname(_DRAWING_NS, "lstStyle"))
        paragraph = ET.SubElement(rich, _qname(_DRAWING_NS, "p"))
        run = ET.SubElement(paragraph, _qname(_DRAWING_NS, "r"))
        ET.SubElement(run, _qname(_DRAWING_NS, "t")).text = _escape_text(chart.title)
        ET.SubElement(title, _qname(_CHART_NS, "layout"))
    plot_area = ET.SubElement(chart_element, _qname(_CHART_NS, "plotArea"))
    ET.SubElement(plot_area, _qname(_CHART_NS, "layout"))
    chart_name = {"column": "barChart", "bar": "barChart", "line": "lineChart", "pie": "pieChart"}[chart.type]
    plot = ET.SubElement(plot_area, _qname(_CHART_NS, chart_name))
    if chart.type in {"column", "bar"}:
        ET.SubElement(plot, _qname(_CHART_NS, "barDir"), {"val": "col" if chart.type == "column" else "bar"})
        ET.SubElement(plot, _qname(_CHART_NS, "grouping"), {"val": "clustered"})
    if chart.type == "line":
        ET.SubElement(plot, _qname(_CHART_NS, "grouping"), {"val": "standard"})
    for index, series in enumerate(chart.series):
        series_element = ET.SubElement(plot, _qname(_CHART_NS, "ser"))
        ET.SubElement(series_element, _qname(_CHART_NS, "idx"), {"val": str(index)})
        ET.SubElement(series_element, _qname(_CHART_NS, "order"), {"val": str(index)})
        if series.name:
            tx = ET.SubElement(series_element, _qname(_CHART_NS, "tx"))
            ET.SubElement(tx, _qname(_CHART_NS, "v")).text = _escape_text(series.name)
        if series.categories:
            category = ET.SubElement(series_element, _qname(_CHART_NS, "cat"))
            reference = ET.SubElement(category, _qname(_CHART_NS, "strRef"))
            ET.SubElement(reference, _qname(_CHART_NS, "f")).text = _chart_reference(chart.worksheet, series.categories)
        value = ET.SubElement(series_element, _qname(_CHART_NS, "val"))
        reference = ET.SubElement(value, _qname(_CHART_NS, "numRef"))
        ET.SubElement(reference, _qname(_CHART_NS, "f")).text = _chart_reference(chart.worksheet, series.values)
    if chart.type != "pie":
        ET.SubElement(plot, _qname(_CHART_NS, "axId"), {"val": "48650112"})
        ET.SubElement(plot, _qname(_CHART_NS, "axId"), {"val": "48672768"})
        category_axis = ET.SubElement(plot_area, _qname(_CHART_NS, "catAx"))
        ET.SubElement(category_axis, _qname(_CHART_NS, "axId"), {"val": "48650112"})
        ET.SubElement(category_axis, _qname(_CHART_NS, "scaling"))
        ET.SubElement(category_axis, _qname(_CHART_NS, "delete"), {"val": "0"})
        ET.SubElement(category_axis, _qname(_CHART_NS, "axPos"), {"val": "b"})
        ET.SubElement(category_axis, _qname(_CHART_NS, "crossAx"), {"val": "48672768"})
        value_axis = ET.SubElement(plot_area, _qname(_CHART_NS, "valAx"))
        ET.SubElement(value_axis, _qname(_CHART_NS, "axId"), {"val": "48672768"})
        ET.SubElement(value_axis, _qname(_CHART_NS, "scaling"))
        ET.SubElement(value_axis, _qname(_CHART_NS, "delete"), {"val": "0"})
        ET.SubElement(value_axis, _qname(_CHART_NS, "axPos"), {"val": "l"})
        ET.SubElement(value_axis, _qname(_CHART_NS, "crossAx"), {"val": "48650112"})
    if chart.legend.position != "none":
        legend = ET.SubElement(chart_element, _qname(_CHART_NS, "legend"))
        position = {"bottom": "b", "top": "t", "left": "l", "right": "r"}[chart.legend.position]
        ET.SubElement(legend, _qname(_CHART_NS, "legendPos"), {"val": position})
        ET.SubElement(legend, _qname(_CHART_NS, "layout"))
    ET.SubElement(chart_element, _qname(_CHART_NS, "plotVisOnly"), {"val": "1"})
    return _xml_bytes(root)


def drawing_xml(
    images: Sequence[tuple[int, object]], charts: Sequence[tuple[int, object]]
) -> bytes:
    """功能：生成承载图片和图表的一张工作表 DrawingML 部件。

    使用方法：XLSX 打包器为包含图片或图表的工作表调用。
    参数：``images`` 为 ``(全局图片编号, Image)`` 序列；``charts`` 为
    ``(全局图表编号, Chart)`` 序列。
    返回：可写入 ``xl/drawings/drawingN.xml`` 的 XML 字节串。
    """
    root = ET.Element(_qname(_SPREADSHEET_DRAWING_NS, "wsDr"))
    relationship_id = 1
    object_id = 1
    for image_id, image in images:
        row, column = cell_index(image.anchor)
        anchor = ET.SubElement(root, _qname(_SPREADSHEET_DRAWING_NS, "oneCellAnchor"))
        start = ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "from"))
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "col")).text = str(column)
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "colOff")).text = str(int(image.offset_x) * 9525)
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "row")).text = str(row)
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "rowOff")).text = str(int(image.offset_y) * 9525)
        ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "ext"), {"cx": str(int(image.width) * 9525), "cy": str(int(image.height) * 9525)})
        picture = ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "pic"))
        non_visual = ET.SubElement(picture, _qname(_SPREADSHEET_DRAWING_NS, "nvPicPr"))
        ET.SubElement(non_visual, _qname(_SPREADSHEET_DRAWING_NS, "cNvPr"), {"id": str(object_id), "name": f"图片 {object_id}", "descr": image.alt_text})
        ET.SubElement(non_visual, _qname(_SPREADSHEET_DRAWING_NS, "cNvPicPr"))
        blip_fill = ET.SubElement(picture, _qname(_SPREADSHEET_DRAWING_NS, "blipFill"))
        ET.SubElement(blip_fill, _qname(_DRAWING_NS, "blip"), {_qname(_REL_NS, "embed"): f"rId{relationship_id}"})
        stretch = ET.SubElement(blip_fill, _qname(_DRAWING_NS, "stretch"))
        ET.SubElement(stretch, _qname(_DRAWING_NS, "fillRect"))
        shape_properties = ET.SubElement(picture, _qname(_SPREADSHEET_DRAWING_NS, "spPr"))
        transform = ET.SubElement(shape_properties, _qname(_DRAWING_NS, "xfrm"))
        ET.SubElement(transform, _qname(_DRAWING_NS, "off"), {"x": "0", "y": "0"})
        ET.SubElement(transform, _qname(_DRAWING_NS, "ext"), {"cx": str(int(image.width) * 9525), "cy": str(int(image.height) * 9525)})
        geometry = ET.SubElement(shape_properties, _qname(_DRAWING_NS, "prstGeom"), {"prst": "rect"})
        ET.SubElement(geometry, _qname(_DRAWING_NS, "avLst"))
        ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "clientData"))
        relationship_id += 1
        object_id += 1
    for chart_id, chart in charts:
        row, column = cell_index(chart.anchor)
        anchor = ET.SubElement(root, _qname(_SPREADSHEET_DRAWING_NS, "oneCellAnchor"))
        start = ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "from"))
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "col")).text = str(column)
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "colOff")).text = "0"
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "row")).text = str(row)
        ET.SubElement(start, _qname(_SPREADSHEET_DRAWING_NS, "rowOff")).text = "0"
        ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "ext"), {"cx": str(int(chart.width * 914400)), "cy": str(int(chart.height * 914400))})
        frame = ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "graphicFrame"))
        non_visual = ET.SubElement(frame, _qname(_SPREADSHEET_DRAWING_NS, "nvGraphicFramePr"))
        ET.SubElement(non_visual, _qname(_SPREADSHEET_DRAWING_NS, "cNvPr"), {"id": str(object_id), "name": f"图表 {object_id}"})
        ET.SubElement(non_visual, _qname(_SPREADSHEET_DRAWING_NS, "cNvGraphicFramePr"))
        transform = ET.SubElement(frame, _qname(_SPREADSHEET_DRAWING_NS, "xfrm"))
        ET.SubElement(transform, _qname(_DRAWING_NS, "off"), {"x": "0", "y": "0"})
        ET.SubElement(transform, _qname(_DRAWING_NS, "ext"), {"cx": str(int(chart.width * 914400)), "cy": str(int(chart.height * 914400))})
        graphic = ET.SubElement(frame, _qname(_DRAWING_NS, "graphic"))
        graphic_data = ET.SubElement(graphic, _qname(_DRAWING_NS, "graphicData"), {"uri": _CHART_NS})
        ET.SubElement(graphic_data, _qname(_CHART_NS, "chart"), {_qname(_REL_NS, "id"): f"rId{relationship_id}"})
        ET.SubElement(anchor, _qname(_SPREADSHEET_DRAWING_NS, "clientData"))
        relationship_id += 1
        object_id += 1
    return _xml_bytes(root)


def drawing_rels_xml(
    images: Sequence[tuple[int, object]], charts: Sequence[tuple[int, object]]
) -> bytes:
    """功能：生成 DrawingML 到图片和图表部件的关系清单。"""
    root = ET.Element(_qname(_PACKAGE_REL_NS, "Relationships"))
    relationship_id = 1
    for image_id, image in images:
        ET.SubElement(root, _qname(_PACKAGE_REL_NS, "Relationship"), {
            "Id": f"rId{relationship_id}",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            "Target": f"../media/image{image_id}.{image.format}",
        })
        relationship_id += 1
    for chart_id, _chart in charts:
        ET.SubElement(root, _qname(_PACKAGE_REL_NS, "Relationship"), {
            "Id": f"rId{relationship_id}",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            "Target": f"../charts/chart{chart_id}.xml",
        })
        relationship_id += 1
    return _xml_bytes(root)


def worksheet_rels(
    table_ids: Sequence[int] = (),
    hyperlink_targets: Sequence[str] = (),
    note_id: Optional[int] = None,
    drawing_id: Optional[int] = None,
) -> bytes:
    """功能：生成单张工作表到超链接和数据表部件的关系清单。

    使用方法：工作表包含超链接或数据表时由XLSX打包器调用。
    参数：``table_ids`` 为当前工作表数据表对应的全局1-based编号；
    ``hyperlink_targets`` 为按单元格顺序排列的外部链接目标。
    返回：工作表 ``.rels`` XML字节串。
    """
    root = ET.Element(_qname(_PACKAGE_REL_NS, "Relationships"))
    for relationship_index, target in enumerate(hyperlink_targets, 1):
        ET.SubElement(
            root,
            _qname(_PACKAGE_REL_NS, "Relationship"),
            {
                "Id": f"rId{relationship_index}",
                "Type": (
                    "http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships/hyperlink"
                ),
                "Target": target,
                "TargetMode": "External",
            },
        )
    relationship_offset = len(hyperlink_targets)
    for relationship_index, table_id in enumerate(table_ids, relationship_offset + 1):
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
    relationship_offset += len(table_ids)
    if note_id is not None:
        ET.SubElement(root, _qname(_PACKAGE_REL_NS, "Relationship"), {
            "Id": f"rId{relationship_offset + 1}",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            "Target": f"../comments{note_id}.xml",
        })
        ET.SubElement(root, _qname(_PACKAGE_REL_NS, "Relationship"), {
            "Id": f"rId{relationship_offset + 2}",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            "Target": f"../drawings/vmlDrawing{note_id}.vml",
        })
        relationship_offset += 2
    if drawing_id is not None:
        ET.SubElement(root, _qname(_PACKAGE_REL_NS, "Relationship"), {
            "Id": f"rId{relationship_offset + 1}",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            "Target": f"../drawings/drawing{drawing_id}.xml",
        })
    return _xml_bytes(root)


def sheet_xml(
    sheet: "Worksheet",
    style_registry: Optional[StyleRegistry] = None,
    table_ids: Sequence[int] = (),
    hyperlink_count: int = 0,
    note_rel_id: Optional[int] = None,
    drawing_rel_id: Optional[int] = None,
) -> bytes:
    """功能：生成单张工作表的完整 SpreadsheetML XML。

    使用方法：由 :class:`XlsxWriter` 对工作簿中的每张工作表调用。
    参数：``sheet`` 为待写出的工作表；``style_registry`` 为工作簿共享样式注册器，
    省略时为当前单表临时创建；``table_ids`` 为本表数据表对应的全局1-based编号；
    ``hyperlink_count`` 为本表外部超链接关系数量；``note_rel_id``、
    ``drawing_rel_id`` 分别为传统批注 VML 和 DrawingML 关系编号。
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
    if sheet.protection.enabled:
        protection_attributes = {}
        if sheet.protection.password:
            protection_attributes["password"] = str(sheet.protection.password)
        if not sheet.protection.select_locked:
            protection_attributes["selectLockedCells"] = "0"
        if not sheet.protection.select_unlocked:
            protection_attributes["selectUnlockedCells"] = "0"
        ET.SubElement(root, _qname(_MAIN_NS, "sheetProtection"), protection_attributes)
    if style_registry is None:
        style_registry = StyleRegistry([sheet])
    values = dict(sheet._values.items())
    coordinates = (
        set(values)
        | set(sheet._formulas)
        | set(sheet._styles)
        | set(sheet._hyperlinks)
    )
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
        auto_filter = ET.SubElement(root, _qname(_MAIN_NS, "autoFilter"), {"ref": sheet.auto_filter_range})
        for column, values in sorted(sheet._filter_conditions.items()):
            filter_column = ET.SubElement(auto_filter, _qname(_MAIN_NS, "filterColumn"), {"colId": str(column)})
            custom = ET.SubElement(filter_column, _qname(_MAIN_NS, "filters"))
            for value in values:
                ET.SubElement(custom, _qname(_MAIN_NS, "filter"), {"val": value})
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

    if sheet.hyperlinks:
        hyperlinks = ET.SubElement(root, _qname(_MAIN_NS, "hyperlinks"))
        external_index = 0
        for address, link in sheet.hyperlinks:
            attributes = {"ref": address}
            if link.target is not None:
                external_index += 1
                attributes[_qname(_REL_NS, "id")] = f"rId{external_index}"
            if link.location is not None:
                attributes["location"] = link.location
            if link.display is not None:
                attributes["display"] = link.display
            if link.tooltip is not None:
                attributes["tooltip"] = link.tooltip
            ET.SubElement(hyperlinks, _qname(_MAIN_NS, "hyperlink"), attributes)

    if drawing_rel_id is not None:
        ET.SubElement(root, _qname(_MAIN_NS, "drawing"), {_qname(_REL_NS, "id"): f"rId{drawing_rel_id}"})
    if note_rel_id is not None:
        ET.SubElement(root, _qname(_MAIN_NS, "legacyDrawing"), {_qname(_REL_NS, "id"): f"rId{note_rel_id}"})

    if sheet.validations:
        validations = ET.SubElement(
            root, _qname(_MAIN_NS, "dataValidations"),
            {"count": str(len(sheet.validations))},
        )
        for validation in sheet.validations:
            attributes = {
                "sqref": validation.range,
                "type": validation.kind,
                "allowBlank": "1" if validation.allow_blank else "0",
                "showDropDown": "0" if validation.show_dropdown else "1",
                "errorStyle": validation.error_style,
            }
            if validation.operator:
                attributes["operator"] = validation.operator
            for key, value in (("promptTitle", validation.prompt_title),
                               ("prompt", validation.prompt),
                               ("errorTitle", validation.error_title),
                               ("error", validation.error)):
                if value is not None:
                    attributes[key] = str(value)
            item = ET.SubElement(root.find(_qname(_MAIN_NS, "dataValidations")),
                                 _qname(_MAIN_NS, "dataValidation"), attributes)
            if validation.kind == "list" and validation.values is not None:
                formula1 = '"' + ",".join(validation.values) + '"'
            else:
                formula1 = validation.formula1
            if formula1 is not None:
                ET.SubElement(item, _qname(_MAIN_NS, "formula1")).text = str(formula1)
            if validation.formula2 is not None:
                ET.SubElement(item, _qname(_MAIN_NS, "formula2")).text = str(validation.formula2)

    for conditional in sheet.conditional_formats:
        group = ET.SubElement(root, _qname(_MAIN_NS, "conditionalFormatting"),
                              {"sqref": conditional.range})
        attributes = {
            "type": conditional.rule,
            "priority": str(conditional.priority),
            "stopIfTrue": "1" if conditional.stop_if_true else "0",
        }
        if conditional.operator:
            attributes["operator"] = conditional.operator
        if conditional.fill or conditional.font:
            dxf_id = style_registry.dxf_ids.get(id(conditional))
            if dxf_id is not None:
                attributes["dxfId"] = str(dxf_id)
        rule = ET.SubElement(group, _qname(_MAIN_NS, "cfRule"), attributes)
        if conditional.rule == "colorScale":
            scale = ET.SubElement(rule, _qname(_MAIN_NS, "colorScale"))
            colors = conditional.options["colors"]
            kinds = ("min", "max") if len(colors) == 2 else ("min", "percentile", "max")
            for index, kind in enumerate(kinds):
                attributes = {"type": kind}
                if kind == "percentile":
                    attributes["val"] = "50"
                ET.SubElement(scale, _qname(_MAIN_NS, "cfvo"), attributes)
                ET.SubElement(scale, _qname(_MAIN_NS, "color"), {"rgb": colors[index]})
        elif conditional.rule == "dataBar":
            bar = ET.SubElement(
                rule, _qname(_MAIN_NS, "dataBar"),
                {"showValue": "1" if conditional.options["show_value"] else "0"},
            )
            ET.SubElement(bar, _qname(_MAIN_NS, "cfvo"), {"type": "min"})
            ET.SubElement(bar, _qname(_MAIN_NS, "cfvo"), {"type": "max"})
            ET.SubElement(bar, _qname(_MAIN_NS, "color"), {"rgb": conditional.options["color"]})
        elif conditional.rule == "iconSet":
            icon_set = ET.SubElement(
                rule, _qname(_MAIN_NS, "iconSet"), {"iconSet": conditional.options["style"]}
            )
            count = int(conditional.options["style"][0])
            for index in range(count):
                attributes = {"type": "percent", "val": str(index * 100 // count)}
                if index:
                    attributes["gte"] = "0"
                ET.SubElement(icon_set, _qname(_MAIN_NS, "cfvo"), attributes)
        elif conditional.formula is not None:
            ET.SubElement(rule, _qname(_MAIN_NS, "formula")).text = str(conditional.formula)

    if sheet.horizontal_page_breaks:
        breaks = ET.SubElement(
            root, _qname(_MAIN_NS, "rowBreaks"),
            {"count": str(len(sheet.horizontal_page_breaks)), "manualBreakCount": str(len(sheet.horizontal_page_breaks))},
        )
        for row in sheet.horizontal_page_breaks:
            ET.SubElement(
                breaks, _qname(_MAIN_NS, "brk"),
                {"id": str(row), "min": "0", "max": "16383", "man": "1"},
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
                {_qname(_REL_NS, "id"): f"rId{hyperlink_count + relationship_index}"},
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
    ET.SubElement(
        root,
        _qname(_PACKAGE_REL_NS, "Relationship"),
        {
            "Id": "rId2",
            "Type": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "Target": "docProps/core.xml",
        },
    )
    return _xml_bytes(root)


def core_properties_xml(properties: WorkbookProperties) -> bytes:
    """功能：把工作簿文档属性写成 Open Packaging 核心属性 XML。

    使用方法：由 ``XlsxWriter`` 内部写入 ``docProps/core.xml``；二次开发也可直接调用。
    参数：``properties`` 为 ``WorkbookProperties`` 对象。
    返回：带 XML 声明的 UTF-8 XML 字节串。
    异常：参数不是 ``WorkbookProperties`` 时抛出 ``TypeError``。
    """
    if not isinstance(properties, WorkbookProperties):
        raise TypeError("properties 必须是 WorkbookProperties")
    root = ET.Element(_qname(_CORE_PROPERTIES_NS, "coreProperties"))
    fields = (
        (_DC_NS, "title", properties.title),
        (_DC_NS, "subject", properties.subject),
        (_DC_NS, "creator", properties.author),
        (_DC_NS, "description", properties.comments),
        (_CORE_PROPERTIES_NS, "keywords", properties.keywords),
        (_CORE_PROPERTIES_NS, "category", properties.category),
        (_CORE_PROPERTIES_NS, "lastModifiedBy", properties.last_modified_by),
    )
    for namespace, name, value in fields:
        if value:
            ET.SubElement(root, _qname(namespace, name)).text = _escape_text(value)
    for name, value in (("created", properties.created), ("modified", properties.modified)):
        if value is not None:
            element = ET.SubElement(
                root,
                _qname(_DCTERMS_NS, name),
                {_qname(_XSI_NS, "type"): "dcterms:W3CDTF"},
            )
            element.text = value.isoformat()
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
                note_entries = []
                sheet_note_ids = {}
                drawing_entries = []
                sheet_drawing_ids = {}
                image_entries = []
                chart_entries = []
                next_table_id = 1
                next_note_id = 1
                next_drawing_id = 1
                next_image_id = 1
                next_chart_id = 1
                for sheet_index, worksheet in enumerate(sheets):
                    current_ids = []
                    for table in worksheet.tables:
                        current_ids.append(next_table_id)
                        table_entries.append((next_table_id, table))
                        next_table_id += 1
                    sheet_table_ids[sheet_index] = tuple(current_ids)
                    notes = tuple(
                        (cell_address(row, column), note)
                        for (row, column), note in sorted(worksheet._notes.items())
                    )
                    if notes:
                        sheet_note_ids[sheet_index] = next_note_id
                        note_entries.append((next_note_id, notes))
                        next_note_id += 1
                    images = []
                    for image in worksheet.images:
                        images.append((next_image_id, image))
                        image_entries.append((next_image_id, image))
                        next_image_id += 1
                    charts = []
                    for chart in worksheet.charts:
                        charts.append((next_chart_id, chart))
                        chart_entries.append((next_chart_id, chart))
                        next_chart_id += 1
                    if images or charts:
                        sheet_drawing_ids[sheet_index] = next_drawing_id
                        drawing_entries.append((next_drawing_id, tuple(images), tuple(charts)))
                        next_drawing_id += 1
                package.writestr(
                    "[Content_Types].xml",
                    content_types(
                        len(sheets), len(table_entries), len(note_entries),
                        len(drawing_entries), len(chart_entries),
                        tuple(image.format for _image_id, image in image_entries),
                    ),
                )
                package.writestr("_rels/.rels", _root_rels())
                package.writestr(
                    "docProps/core.xml",
                    core_properties_xml(self.workbook.properties),
                )
                package.writestr(
                    "xl/workbook.xml",
                    workbook_xml(sheets, self.workbook.named_ranges, self.workbook.protection),
                )
                package.writestr("xl/_rels/workbook.xml.rels", workbook_rels(sheets))
                package.writestr("xl/styles.xml", style_registry.xml())
                for sheet_index, worksheet in enumerate(sheets):
                    table_ids = sheet_table_ids[sheet_index]
                    note_id = sheet_note_ids.get(sheet_index)
                    drawing_id = sheet_drawing_ids.get(sheet_index)
                    hyperlink_targets = tuple(
                        link.target
                        for _address, link in worksheet.hyperlinks
                        if link.target is not None
                    )
                    package.writestr(
                        f"xl/worksheets/sheet{sheet_index + 1}.xml",
                        sheet_xml(
                            worksheet,
                            style_registry,
                            table_ids,
                            len(hyperlink_targets),
                            (len(hyperlink_targets) + len(table_ids) + 2) if note_id is not None else None,
                            (
                                len(hyperlink_targets) + len(table_ids)
                                + (2 if note_id is not None else 0) + 1
                            ) if drawing_id is not None else None,
                        ),
                    )
                    if table_ids or hyperlink_targets or note_id is not None or drawing_id is not None:
                        package.writestr(
                            f"xl/worksheets/_rels/sheet{sheet_index + 1}.xml.rels",
                            worksheet_rels(table_ids, hyperlink_targets, note_id, drawing_id),
                        )
                for table_id, table in table_entries:
                    package.writestr(
                        f"xl/tables/table{table_id}.xml",
                        table_xml(table, table_id),
                    )
                for note_id, notes in note_entries:
                    package.writestr(f"xl/comments{note_id}.xml", comments_xml(notes))
                    package.writestr(
                        f"xl/drawings/vmlDrawing{note_id}.vml", vml_comments_xml(notes)
                    )
                for drawing_id, images, charts in drawing_entries:
                    package.writestr(
                        f"xl/drawings/drawing{drawing_id}.xml", drawing_xml(images, charts)
                    )
                    package.writestr(
                        f"xl/drawings/_rels/drawing{drawing_id}.xml.rels",
                        drawing_rels_xml(images, charts),
                    )
                for image_id, image in image_entries:
                    package.writestr(
                        f"xl/media/image{image_id}.{image.format}", image.payload
                    )
                for chart_id, chart in chart_entries:
                    package.writestr(f"xl/charts/chart{chart_id}.xml", chart_xml(chart))
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
    "core_properties_xml",
    "workbook_xml",
    "workbook_rels",
    "cell_xml",
    "formula_xml",
    "table_xml",
    "worksheet_rels",
    "sheet_xml",
]
