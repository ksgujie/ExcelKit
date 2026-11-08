"""XLSX 写出器：基于标准库 zipfile/xml 生成 OpenXML 格式，无外部依赖。

采用内联字符串（inlineStr）避免共享字符串表，代码更精简；
样式写入基础字体、填充、边框、对齐、数字格式。
"""

from __future__ import annotations

import re
import zipfile
from datetime import date, datetime
from xml.etree import ElementTree as ET

from excelkit.errors import InvalidFileError

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_MARKUP = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _q(tag: str) -> str:
    return f"{{{_NS_MAIN}}}{tag}"


def _escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _cell_ref(row: int, column: int) -> str:
    """0-based 行列 -> A1 地址。"""
    letters = []
    n = column + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters)) + str(row + 1)


def _format_value(val) -> str:
    """把 Python 值转为 XML 内联字符串文本。"""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%dT%H:%M:%S")
    if isinstance(val, date):
        return val.isoformat()
    return str(val)


def _style_id(style, styles_cache: dict, style_objects: dict) -> int:
    """计算样式在 styles.xml 中的 cellXfs 索引，并记录 style 对象。"""
    if style is None:
        return 0  # 0 表示默认样式
    key = _style_key(style)
    if key in styles_cache:
        return styles_cache[key]
    idx = len(styles_cache) + 1
    styles_cache[key] = idx
    style_objects[idx] = style
    return idx


def _style_key(style) -> tuple:
    font = style.font
    fill = style.fill
    border = style.border
    alignment = style.alignment
    return (
        (
            (font.name, font.size, font.bold, font.italic, font.color, font.underline)
            if font else None
        ),
        ((fill.fg_color, fill.pattern_type) if fill else None),
        (
            (
                (border.top.style if border.top else None, border.top.color if border.top else None),
                (border.bottom.style if border.bottom else None, border.bottom.color if border.bottom else None),
                (border.left.style if border.left else None, border.left.color if border.left else None),
                (border.right.style if border.right else None, border.right.color if border.right else None),
            )
            if border else None
        ),
        (
            (alignment.horizontal, alignment.vertical, alignment.wrap_text, alignment.indent)
            if alignment else None
        ),
        style.number_format,
    )


def _build_styles(styles_cache: dict, style_objects: dict) -> bytes:
    """生成 styles.xml 内容。

    styles_cache: key=_style_key(tuple) -> idx
    style_objects: idx -> Style 对象
    """
    style_items = sorted(style_objects.items(), key=lambda kv: kv[0])  # idx -> style

    root = ET.Element(_q("styleSheet"))
    # fonts: index 0 = default
    fonts = ET.SubElement(root, _q("fonts"))
    fonts.set("count", str(len(style_items) + 1))
    ET.SubElement(fonts, _q("font"))  # index 0: default
    for idx, style in style_items:
        font_el = ET.SubElement(fonts, _q("font"))
        if style.font:
            f = style.font
            if f.name:
                name_el = ET.SubElement(font_el, _q("name"))
                name_el.set("val", f.name)
            if f.size:
                sz = ET.SubElement(font_el, _q("sz"))
                sz.set("val", str(f.size))
            if f.bold:
                ET.SubElement(font_el, _q("b"))
            if f.italic:
                ET.SubElement(font_el, _q("i"))
            if f.underline:
                ET.SubElement(font_el, _q("u"))
            if f.color:
                color_el = ET.SubElement(font_el, _q("color"))
                color_el.set("rgb", f.color)
        else:
            ET.SubElement(font_el, _q("name")).set("val", "Calibri")
            ET.SubElement(font_el, _q("sz")).set("val", "11")
    # fills: index 0 = none, index 1 = default none
    fills = ET.SubElement(root, _q("fills"))
    fills.set("count", str(len(style_items) + 2))
    ET.SubElement(fills, _q("fill"))  # index 0: none
    fill_default = ET.SubElement(fills, _q("fill"))  # index 1: default
    pf_def = ET.SubElement(fill_default, _q("patternFill"))
    pf_def.set("patternType", "none")
    for idx, style in style_items:
        fill_el = ET.SubElement(fills, _q("fill"))
        pattern_fill = ET.SubElement(fill_el, _q("patternFill"))
        pattern_fill.set("patternType", style.fill.pattern_type if style.fill else "solid")
        if style.fill and style.fill.fg_color:
            color_el = ET.SubElement(pattern_fill, _q("fgColor"))
            color_el.set("rgb", style.fill.fg_color)
    # borders: index 0 = none
    borders = ET.SubElement(root, _q("borders"))
    borders.set("count", str(len(style_items) + 1))
    ET.SubElement(borders, _q("border"))  # index 0: none
    for idx, style in style_items:
        border_el = ET.SubElement(borders, _q("border"))
        if style.border:
            b = style.border
            for side_name, side in (("top", b.top), ("bottom", b.bottom), ("left", b.left), ("right", b.right)):
                if side and side.style:
                    side_el = ET.SubElement(border_el, _q(side_name))
                    style_el = ET.SubElement(side_el, _q("style"))
                    style_el.set("val", side.style)
                    if side.color:
                        color_el = ET.SubElement(side_el, _q("color"))
                        color_el.set("rgb", side.color)
    # cellStyleXfs, cellXfs
    cell_style_xfs = ET.SubElement(root, _q("cellStyleXfs"))
    cell_style_xfs.set("count", "1")
    ET.SubElement(cell_style_xfs, _q("xf")).set("xfId", "0")

    cell_xfs = ET.SubElement(root, _q("cellXfs"))
    cell_xfs.set("count", str(len(style_items) + 1))
    # index 0: default xf
    ET.SubElement(cell_xfs, _q("xf"))

    num_fmts = None
    for idx, style in style_items:  # idx 从 1 开始，对应 cellXfs 索引
        xf = ET.SubElement(cell_xfs, _q("xf"))
        apply_font = style.font is not None
        apply_fill = style.fill is not None
        apply_border = style.border is not None
        apply_align = style.alignment is not None
        apply_numfmt = style.number_format is not None
        xf.set("applyFont", "1" if apply_font else "0")
        xf.set("applyFill", "1" if apply_fill else "0")
        xf.set("applyBorder", "1" if apply_border else "0")
        xf.set("applyAlignment", "1" if apply_align else "0")
        xf.set("applyNumberFormat", "1" if apply_numfmt else "0")
        # 索引对齐：fonts[0]=default，style 字体在 fonts[idx]
        # fills[0]=none,[1]=default none，style 填充在 fills[idx+1]
        # borders[0]=none，style 边框在 borders[idx]
        xf.set("fontId", str(idx))
        xf.set("fillId", str(idx + 1))
        xf.set("borderId", str(idx))
        if apply_align:
            al = ET.SubElement(xf, _q("alignment"))
            if style.alignment.horizontal:
                al.set("horizontal", style.alignment.horizontal)
            if style.alignment.vertical:
                al.set("vertical", style.alignment.vertical)
            if style.alignment.wrap_text:
                al.set("wrapText", "1")
        if apply_numfmt:
            if num_fmts is None:
                num_fmts = ET.SubElement(root, _q("numFmts"))
                num_fmts.set("count", "0")
            count = int(num_fmts.get("count")) + 1
            num_fmts.set("count", str(count))
            nf = ET.SubElement(num_fmts, _q("numFmt"))
            nf.set("numFmtId", str(164 + idx))
            nf.set("formatCode", style.number_format)
            xf.set("numFmtId", str(164 + idx))

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _build_sheet_xml(ws, shared_strings: list, styles_cache: dict, style_objects: dict) -> bytes:
    """生成单张工作表的 sheetN.xml 内容。"""
    root = ET.Element(_q("worksheet"))
    root.set("xmlns", _NS_MAIN)

    sheet_format = ET.SubElement(root, _q("sheetFormatPr"))
    sheet_format.set("defaultRowHeight", "15")

    # 合并单元格
    if ws._merged:
        merge_cells = ET.SubElement(root, _q("mergeCells"))
        merge_cells.set("count", str(len(ws._merged)))
        for min_r, min_c, max_r, max_c in ws._merged:
            mc = ET.SubElement(merge_cells, _q("mergeCell"))
            mc.set("ref", f"{_cell_ref(min_r, min_c)}:{_cell_ref(max_r, max_c)}")

    # sheetData
    sheet_data = ET.SubElement(root, _q("sheetData"))
    max_r = ws._max_row
    max_c = ws._max_column
    if max_r >= 0 and max_c >= 0:
        for r in range(max_r + 1):
            row_el = ET.SubElement(sheet_data, _q("row"))
            row_el.set("r", str(r + 1))
            has_cell = False
            for c in range(max_c + 1):
                cell = ws._cells.get((r, c))
                if cell is None:
                    continue
                val = cell.get("value")
                formula = cell.get("formula")
                style = cell.get("style")
                c_el = ET.SubElement(row_el, _q("c"))
                c_el.set("r", _cell_ref(r, c))
                if style is not None:
                    c_el.set("s", str(_style_id(style, styles_cache, style_objects)))
                if formula is not None:
                    f_el = ET.SubElement(c_el, _q("f"))
                    f_el.text = formula.lstrip("=")
                    v_el = ET.SubElement(c_el, _q("v"))
                    v_el.text = _format_value(val) if val is not None else ""
                    has_cell = True
                elif val is not None:
                    if isinstance(val, bool):
                        c_el.set("t", "b")
                        v_el = ET.SubElement(c_el, _q("v"))
                        v_el.text = "1" if val else "0"
                    elif isinstance(val, (int, float)):
                        c_el.set("t", "n")
                        v_el = ET.SubElement(c_el, _q("v"))
                        v_el.text = str(val)
                    else:
                        c_el.set("t", "inlineStr")
                        is_el = ET.SubElement(c_el, _q("inlineStr"))
                        t_el = ET.SubElement(is_el, _q("t"))
                        t_el.set("xml:space", "preserve")
                        t_el.text = _format_value(val)
                    has_cell = True
                else:
                    # 只有样式的空单元格
                    c_el.set("t", "inlineStr")
                    is_el = ET.SubElement(c_el, _q("inlineStr"))
                    t_el = ET.SubElement(is_el, _q("t"))
                    t_el.set("xml:space", "preserve")
                    t_el.text = ""
                    has_cell = True
            if not has_cell and r <= max_r:
                # 空行可不写，dimension 靠 max 行列体现
                pass

    # dimension
    if max_r >= 0 and max_c >= 0:
        dim = ET.SubElement(root, _q("dimension"))
        dim.set("ref", f"A1:{_cell_ref(max_r, max_c)}")

    # page setup 简化
    page = ws._page
    if page.orientation == "landscape" or page.print_area or page.fit_width:
        pg = ET.SubElement(root, _q("pageSetup"))
        pg.set("orientation", page.orientation)
        pg.set("fitToWidth", str(page.fit_width or 1))
        pg.set("fitToHeight", str(page.fit_height or 0))
        pg.set("paperSize", "9")  # A4

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def save_xlsx(workbook, filename: str, *, formulas: bool = False) -> None:
    """将工作簿写出为 XLSX 文件。"""
    styles_cache: dict = {}  # key=_style_key(tuple) -> idx
    style_objects: dict = {}  # idx -> Style
    sheets_xml = []
    shared_strings: list = []  # 精简版不使用共享字符串

    for ws in workbook.sheets:
        sheet_bytes = _build_sheet_xml(ws, shared_strings, styles_cache, style_objects)
        sheets_xml.append((ws.name, sheet_bytes))

    # styles.xml
    styles_bytes = _build_styles(styles_cache, style_objects)

    # workbook.xml
    wb_root = ET.Element(_q("workbook"))
    wb_root.set("xmlns", _NS_MAIN)
    sheets_el = ET.SubElement(wb_root, _q("sheets"))
    for i, (name, _) in enumerate(sheets_xml):
        sheet_el = ET.SubElement(sheets_el, _q("sheet"))
        sheet_el.set("name", name)
        sheet_el.set("sheetId", str(i + 1))
        sheet_el.set(
            f"{{{_NS_REL}}}id", f"rId{i + 1}"
        )
    workbook_xml = ET.tostring(wb_root, encoding="utf-8", xml_declaration=True)

    # 关系
    rels_root = ET.Element("Relationships")
    rels_root.set("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
    for i, (name, _) in enumerate(sheets_xml):
        rel = ET.SubElement(rels_root, "Relationship")
        rel.set("Id", f"rId{i + 1}")
        rel.set("Type", f"{_NS_REL}/worksheet")
        rel.set("Target", f"worksheets/sheet{i + 1}.xml")
    rels_xml = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

    # content types
    ct_root = ET.Element("Types")
    ct_root.set("xmlns", _NS_CT)
    d1 = ET.SubElement(ct_root, "Default")
    d1.set("Extension", "rels")
    d1.set("ContentType", "application/vnd.openxmlformats-package.relationships+xml")
    d2 = ET.SubElement(ct_root, "Default")
    d2.set("Extension", "xml")
    d2.set("ContentType", "application/xml")
    o1 = ET.SubElement(ct_root, "Override")
    o1.set("PartName", "/xl/workbook.xml")
    o1.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    o2 = ET.SubElement(ct_root, "Override")
    o2.set("PartName", "/xl/styles.xml")
    o2.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
    for i in range(len(sheets_xml)):
        o = ET.SubElement(ct_root, "Override")
        o.set("PartName", f"/xl/worksheets/sheet{i + 1}.xml")
        o.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    content_types = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

    # _rels/.rels
    rels_dot = ET.Element("Relationships")
    rels_dot.set("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
    rel = ET.SubElement(rels_dot, "Relationship")
    rel.set("Id", "rId1")
    rel.set("Type", f"{_NS_REL}/officeDocument")
    rel.set("Target", "xl/workbook.xml")
    rels_dot_xml = ET.tostring(rels_dot, encoding="utf-8", xml_declaration=True)

    # 写入 zip
    try:
        with zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("_rels/.rels", rels_dot_xml)
            zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
            zf.writestr("xl/workbook.xml", workbook_xml)
            zf.writestr("xl/styles.xml", styles_bytes)
            for i, (name, sheet_bytes) in enumerate(sheets_xml):
                zf.writestr(f"xl/worksheets/sheet{i + 1}.xml", sheet_bytes)
    except Exception as exc:
        raise InvalidFileError(f"无法写出 XLSX 文件：{exc}") from exc


__all__ = ["save_xlsx"]
