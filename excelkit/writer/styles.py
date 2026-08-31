"""单元格 Style 到 XLSX styles.xml 的内部注册与序列化。"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Dict, Iterable, List, Sequence, Tuple, TypeVar

from ..style import Alignment, Border, BorderSide, DEFAULT_STYLE, Fill, Font, Style

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_BUILTIN_NUMBER_FORMATS = {
    "General": 0,
    "0": 1,
    "0.00": 2,
    "#,##0": 3,
    "#,##0.00": 4,
    "0%": 9,
    "0.00%": 10,
    "0.00E+00": 11,
    "# ?/?": 12,
    "# ??/??": 13,
    "mm-dd-yy": 14,
    "h:mm": 20,
    "h:mm:ss": 21,
    "m/d/yy h:mm": 22,
}
_ValueType = TypeVar("_ValueType")


def _tag(local_name: str) -> str:
    """功能：生成 SpreadsheetML 主命名空间限定名称。

    使用方法：样式 XML 序列化内部调用。
    参数：``local_name`` 为 XML 元素本地名称。
    返回：``{主命名空间}本地名称`` 字符串。
    """
    return f"{{{_MAIN_NS}}}{local_name}"


def _unique(values: Iterable[_ValueType]) -> List[_ValueType]:
    """功能：按首次出现顺序对可哈希值去重。

    使用方法：注册字体、填充、边框和样式时内部调用。
    参数：``values`` 为可迭代的不可变可哈希对象。
    返回：保留首次出现顺序的列表。
    """
    result: List[_ValueType] = []
    seen = set()
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


class StyleRegistry:
    """为一个工作簿中的不可变 Style 分配稳定 XLSX 索引。"""

    def __init__(self, sheets: Sequence[object]) -> None:
        """功能：扫描工作表并注册全部显式样式及其组件。

        使用方法：XlsxWriter 写包前创建一次并传给各工作表。
        参数：``sheets`` 为带 ``_styles`` 字典的工作表序列。
        返回：无；默认样式固定为索引 0。
        """
        explicit = [
            style
            for sheet in sheets
            for _coordinate, style in sorted(sheet._styles.items())
        ]
        self.styles = _unique([DEFAULT_STYLE, *explicit])
        self.fonts = _unique(style.font for style in self.styles)
        self.borders = _unique(style.border for style in self.styles)
        self.custom_fills = _unique(
            style.fill for style in self.styles if style.fill != Fill()
        )
        self.style_ids = {style: index for index, style in enumerate(self.styles)}
        self.font_ids = {font: index for index, font in enumerate(self.fonts)}
        self.border_ids = {border: index for index, border in enumerate(self.borders)}
        self.fill_ids = {Fill(): 0}
        self.fill_ids.update(
            {fill: index + 2 for index, fill in enumerate(self.custom_fills)}
        )
        custom_formats = _unique(
            style.number_format
            for style in self.styles
            if style.number_format not in _BUILTIN_NUMBER_FORMATS
        )
        self.custom_number_formats = {
            number_format: index + 164
            for index, number_format in enumerate(custom_formats)
        }
        self.dxfs = [item for sheet in sheets for item in getattr(sheet, "_conditionals", ())
                     if item.fill or item.font]
        self.dxf_ids = {id(item): index for index, item in enumerate(self.dxfs)}

    def style_id(self, style: Style) -> int:
        """功能：取得 Style 对应的 0-based XLSX cellXfs 索引。

        使用方法：工作表写出单元格 ``s`` 属性时调用。
        参数：``style`` 为已在当前注册器中扫描到的 :class:`Style`。
        返回：非负整数样式索引。
        异常：样式未注册时抛出 ``KeyError``。
        """
        return self.style_ids[style]

    def _number_format_id(self, number_format: str) -> int:
        """功能：取得内置或自定义数字格式 ID。

        使用方法：生成 cellXfs 时内部调用。
        参数：``number_format`` 为非空 Excel 格式代码。
        返回：内置 ID 或从 164 开始的自定义 ID。
        """
        if number_format in _BUILTIN_NUMBER_FORMATS:
            return _BUILTIN_NUMBER_FORMATS[number_format]
        return self.custom_number_formats[number_format]

    def xml(self) -> bytes:
        """功能：生成完整 XLSX styles.xml。

        使用方法：XlsxWriter 写入 ``xl/styles.xml`` 时调用。
        参数：无；内容来自注册时扫描的样式。
        返回：包含字体、填充、边框、数字格式和 cellXfs 的 UTF-8 XML 字节串。
        """
        root = ET.Element(_tag("styleSheet"))
        self._append_number_formats(root)
        self._append_fonts(root)
        self._append_fills(root)
        self._append_borders(root)
        self._append_cell_styles(root)
        self._append_dxfs(root)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _append_dxfs(self, root: ET.Element) -> None:
        """功能：写出条件格式使用的差异样式集合。

        使用方法：由 :meth:`xml` 在生成完整样式表时内部调用。
        参数：``root`` 为 ``styleSheet`` 根元素。
        返回：``None``；没有差异样式时不生成 ``dxfs`` 元素。
        """
        if not self.dxfs:
            return
        element = ET.SubElement(root, _tag("dxfs"), {"count": str(len(self.dxfs))})
        for item in self.dxfs:
            dxf = ET.SubElement(element, _tag("dxf"))
            if item.font:
                font = ET.SubElement(dxf, _tag("font"))
                ET.SubElement(font, _tag("color"), {"rgb": item.font})
            if item.fill:
                fill = ET.SubElement(dxf, _tag("fill"))
                pattern = ET.SubElement(fill, _tag("patternFill"), {"patternType": "solid"})
                ET.SubElement(pattern, _tag("fgColor"), {"rgb": item.fill})

    def _append_number_formats(self, root: ET.Element) -> None:
        """功能：向样式表追加自定义数字格式集合。

        使用方法：由 :meth:`xml` 内部调用。
        参数：``root`` 为 styleSheet 根元素。
        返回：``None``；没有自定义格式时不生成 numFmts。
        """
        if not self.custom_number_formats:
            return
        element = ET.SubElement(
            root, _tag("numFmts"), {"count": str(len(self.custom_number_formats))}
        )
        for code, format_id in self.custom_number_formats.items():
            ET.SubElement(
                element,
                _tag("numFmt"),
                {"numFmtId": str(format_id), "formatCode": code},
            )

    def _append_fonts(self, root: ET.Element) -> None:
        """功能：向样式表追加字体集合。

        使用方法：由 :meth:`xml` 内部调用。
        参数：``root`` 为 styleSheet 根元素。
        返回：``None``。
        """
        fonts = ET.SubElement(root, _tag("fonts"), {"count": str(len(self.fonts))})
        for font in self.fonts:
            element = ET.SubElement(fonts, _tag("font"))
            ET.SubElement(element, _tag("name"), {"val": font.name})
            ET.SubElement(element, _tag("sz"), {"val": str(font.size)})
            if font.bold:
                ET.SubElement(element, _tag("b"))
            if font.italic:
                ET.SubElement(element, _tag("i"))
            if font.underline:
                ET.SubElement(element, _tag("u"))
            if font.color:
                ET.SubElement(element, _tag("color"), {"rgb": font.color})

    def _append_fills(self, root: ET.Element) -> None:
        """功能：向样式表追加 Excel 必需默认填充和自定义实心填充。

        使用方法：由 :meth:`xml` 内部调用。
        参数：``root`` 为 styleSheet 根元素。
        返回：``None``。
        """
        count = 2 + len(self.custom_fills)
        fills = ET.SubElement(root, _tag("fills"), {"count": str(count)})
        first = ET.SubElement(fills, _tag("fill"))
        ET.SubElement(first, _tag("patternFill"), {"patternType": "none"})
        second = ET.SubElement(fills, _tag("fill"))
        ET.SubElement(second, _tag("patternFill"), {"patternType": "gray125"})
        for fill in self.custom_fills:
            element = ET.SubElement(fills, _tag("fill"))
            pattern = ET.SubElement(
                element, _tag("patternFill"), {"patternType": "solid"}
            )
            ET.SubElement(pattern, _tag("fgColor"), {"rgb": fill.color or "00000000"})
            ET.SubElement(pattern, _tag("bgColor"), {"indexed": "64"})

    def _append_borders(self, root: ET.Element) -> None:
        """功能：向样式表追加四边边框集合。

        使用方法：由 :meth:`xml` 内部调用。
        参数：``root`` 为 styleSheet 根元素。
        返回：``None``。
        """
        borders = ET.SubElement(
            root, _tag("borders"), {"count": str(len(self.borders))}
        )
        for border in self.borders:
            element = ET.SubElement(borders, _tag("border"))
            for name in ("left", "right", "top", "bottom"):
                side = getattr(border, name)
                attributes = {"style": side.style} if side.style else {}
                side_element = ET.SubElement(element, _tag(name), attributes)
                if side.color:
                    ET.SubElement(side_element, _tag("color"), {"rgb": side.color})
            ET.SubElement(element, _tag("diagonal"))

    def _append_cell_styles(self, root: ET.Element) -> None:
        """功能：向样式表追加基础样式、单元格格式和 Normal 样式声明。

        使用方法：由 :meth:`xml` 内部调用。
        参数：``root`` 为 styleSheet 根元素。
        返回：``None``。
        """
        base = ET.SubElement(root, _tag("cellStyleXfs"), {"count": "1"})
        ET.SubElement(
            base,
            _tag("xf"),
            {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0"},
        )
        formats = ET.SubElement(
            root, _tag("cellXfs"), {"count": str(len(self.styles))}
        )
        for style in self.styles:
            number_format_id = self._number_format_id(style.number_format)
            attributes = {
                "numFmtId": str(number_format_id),
                "fontId": str(self.font_ids[style.font]),
                "fillId": str(self.fill_ids[style.fill]),
                "borderId": str(self.border_ids[style.border]),
                "xfId": "0",
            }
            if style.font != Font():
                attributes["applyFont"] = "1"
            if style.fill != Fill():
                attributes["applyFill"] = "1"
            if style.border != Border():
                attributes["applyBorder"] = "1"
            if style.number_format != "General":
                attributes["applyNumberFormat"] = "1"
            if style.alignment != Alignment():
                attributes["applyAlignment"] = "1"
            cell_format = ET.SubElement(formats, _tag("xf"), attributes)
            if style.alignment != Alignment():
                alignment_attributes = {}
                if style.alignment.horizontal:
                    alignment_attributes["horizontal"] = style.alignment.horizontal
                if style.alignment.vertical:
                    alignment_attributes["vertical"] = style.alignment.vertical
                if style.alignment.wrap_text:
                    alignment_attributes["wrapText"] = "1"
                ET.SubElement(cell_format, _tag("alignment"), alignment_attributes)
        cell_styles = ET.SubElement(root, _tag("cellStyles"), {"count": "1"})
        ET.SubElement(
            cell_styles,
            _tag("cellStyle"),
            {"name": "Normal", "xfId": "0", "builtinId": "0"},
        )
