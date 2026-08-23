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

from ..address import cell_index
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
            worksheet.label_color = tab_color.get("rgb")
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
        return workbook


__all__ = []
