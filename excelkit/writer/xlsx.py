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

from ..address import cell_address
from ..style import DEFAULT_STYLE
from .styles import StyleRegistry

if TYPE_CHECKING:
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


def content_types(sheet_count: int) -> bytes:
    """功能：根据工作表数量生成 XLSX 内容类型声明。

    使用方法：``content_types(len(workbook.sheets))``；由 :class:`XlsxWriter`
    写入 ``[Content_Types].xml``。
    参数：``sheet_count`` 为大于等于 1 的整数工作表数量，属于计数而非索引。
    返回：UTF-8 XML 字节串。
    异常：参数不是正整数或是布尔值时抛出 ``ValueError``。
    """
    if (
        isinstance(sheet_count, bool)
        or not isinstance(sheet_count, int)
        or sheet_count < 1
    ):
        raise ValueError("工作表数量必须是大于等于 1 的整数")

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
    return _xml_bytes(root)


def workbook_xml(sheets: Sequence["Worksheet"]) -> bytes:
    """功能：生成包含工作表名称、顺序和关系编号的工作簿 XML。

    使用方法：``workbook_xml(workbook.sheets)``。
    参数：``sheets`` 为按创建顺序排列的工作表序列。
    返回：可写入 ``xl/workbook.xml`` 的 UTF-8 XML 字节串。
    """
    root = ET.Element(_qname(_MAIN_NS, "workbook"))
    sheet_elements = ET.SubElement(root, _qname(_MAIN_NS, "sheets"))
    for sheet_index, worksheet in enumerate(sheets):
        ET.SubElement(
            sheet_elements,
            _qname(_MAIN_NS, "sheet"),
            {
                "name": worksheet.label,
                "sheetId": str(sheet_index + 1),
                _qname(_REL_NS, "id"): f"rId{sheet_index + 1}",
            },
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


def formula_xml(address: str, formula: str) -> ET.Element:
    """功能：把公式转换为 SpreadsheetML 公式单元格元素。

    使用方法：由 :func:`sheet_xml` 调用 ``formula_xml("C2", "=SUM(A2:B2)")``。
    参数：``address`` 为规范化 A1 地址；``formula`` 为包含表达式的公式字符串，
    可以带或不带前导 ``=``。
    返回：包含 ``f`` 子元素且不带缓存计算值的 ``c`` 元素。
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
    return cell


def sheet_xml(
    sheet: "Worksheet", style_registry: Optional[StyleRegistry] = None
) -> bytes:
    """功能：生成单张工作表的完整 SpreadsheetML XML。

    使用方法：由 :class:`XlsxWriter` 对工作簿中的每张工作表调用。
    参数：``sheet`` 为待写出的工作表；``style_registry`` 为工作簿共享样式注册器，
    省略时为当前单表临时创建。普通值、公式和样式按 0-based 行列索引合并排序。
    返回：包含精确数据边界和 ``sheetData`` 的 UTF-8 XML 字节串。
    """
    root = ET.Element(_qname(_MAIN_NS, "worksheet"))
    if sheet.label_color is not None:
        properties = ET.SubElement(root, _qname(_MAIN_NS, "sheetPr"))
        ET.SubElement(
            properties,
            _qname(_MAIN_NS, "tabColor"),
            {"rgb": sheet.label_color},
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

    sheet_data = ET.SubElement(root, _qname(_MAIN_NS, "sheetData"))
    row_indexes = sorted({row for row, _column in coordinates})
    for row in row_indexes:
        row_element = ET.SubElement(
            sheet_data, _qname(_MAIN_NS, "row"), {"r": str(row + 1)}
        )
        columns = sorted(column for item_row, column in coordinates if item_row == row)
        for column in columns:
            address = cell_address(row, column)
            style = sheet._styles.get((row, column), DEFAULT_STYLE)
            style_id = style_registry.style_id(style)
            formula = sheet._formulas.get((row, column))
            if formula is not None:
                element = formula_xml(address, formula)
            else:
                element = cell_xml(address, values.get((row, column)))
            if element is None and style_id:
                element = ET.Element(_qname(_MAIN_NS, "c"), {"r": address})
            if element is not None:
                if style_id:
                    element.set("s", str(style_id))
                row_element.append(element)
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
                package.writestr("[Content_Types].xml", content_types(len(sheets)))
                package.writestr("_rels/.rels", _root_rels())
                package.writestr("xl/workbook.xml", workbook_xml(sheets))
                package.writestr("xl/_rels/workbook.xml.rels", workbook_rels(sheets))
                package.writestr("xl/styles.xml", style_registry.xml())
                for sheet_index, worksheet in enumerate(sheets):
                    package.writestr(
                        f"xl/worksheets/sheet{sheet_index + 1}.xml",
                        sheet_xml(worksheet, style_registry),
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
    "sheet_xml",
]
