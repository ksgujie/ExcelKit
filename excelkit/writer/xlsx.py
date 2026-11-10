"""XLSX 写出器：基于标准库 zipfile/xml 生成 OpenXML 格式，无外部依赖。

采用内联字符串（inlineStr）避免共享字符串表；
样式写入基础字体、填充、边框、对齐、数字格式；
支持 PNG/JPEG 图片（浮动 oneCellAnchor / 嵌入 twoCellAnchor editAs="twoCell"）。
"""

from __future__ import annotations

import re
import zipfile
from datetime import date, datetime
from xml.etree import ElementTree as ET

from excelkit.address import index_to_cell, range_to_index
from excelkit.core.image import EMU_PER_PIXEL
from excelkit.errors import InvalidFileError

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# 关系类型
_REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
_REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
_REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

# 内容类型
_CT_WB = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
_CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
_CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
_CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
_CT_XML = "application/xml"



def _register_xlsx_namespaces(*, drawing: bool = False) -> None:
    """集中注册 OOXML 序列化所需前缀。

    注意：xml.etree.ElementTree.register_namespace 是进程级全局状态，stdlib 无作用域隔离 API。
    集中在此注册可避免散落重复调用；宿主程序若依赖自定义前缀，可能受影响（ET 固有限制）。
    """
    ET.register_namespace("", _NS_MAIN)
    ET.register_namespace("r", _NS_R)
    if drawing:
        ET.register_namespace("xdr", _NS_XDR)
        ET.register_namespace("a", _NS_A)

def _q(tag: str) -> str:
    return f"{{{_NS_MAIN}}}{tag}"


def _cell_ref(row: int, column: int) -> str:
    """0-based 行列 -> A1 地址。"""
    letters = []
    n = column + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters)) + str(row + 1)


def _format_value(val) -> str:
    """把 Python 值转为 XML 文本。"""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "1" if val else "0"
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%dT%H:%M:%S")
    if isinstance(val, date):
        return val.isoformat()
    return str(val)


def _style_id(style, styles_cache: dict, style_objects: dict) -> int:
    """计算样式在 styles.xml 中的 cellXfs 索引，并记录 style 对象。"""
    if style is None:
        return 0
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
            if font
            else None
        ),
        ((fill.fg_color, fill.pattern_type) if fill else None),
        (
            (
                (border.top.style if border.top else None, border.top.color if border.top else None),
                (border.bottom.style if border.bottom else None, border.bottom.color if border.bottom else None),
                (border.left.style if border.left else None, border.left.color if border.left else None),
                (border.right.style if border.right else None, border.right.color if border.right else None),
            )
            if border
            else None
        ),
        (
            (alignment.horizontal, alignment.vertical, alignment.wrap_text, alignment.indent)
            if alignment
            else None
        ),
        style.number_format,
    )


def _build_styles(styles_cache: dict, style_objects: dict) -> bytes:
    """生成 styles.xml 内容（包含 Excel 兼容的默认样式定义）。"""
    style_items = sorted(style_objects.items(), key=lambda kv: kv[0])

    root = ET.Element(_q("styleSheet"))

    # fonts: index 0 = default Calibri 11
    fonts = ET.SubElement(root, _q("fonts"))
    fonts.set("count", str(len(style_items) + 1))
    font0 = ET.SubElement(fonts, _q("font"))
    ET.SubElement(font0, _q("sz")).set("val", "11")
    ET.SubElement(font0, _q("color")).set("theme", "1")
    ET.SubElement(font0, _q("name")).set("val", "Calibri")
    ET.SubElement(font0, _q("family")).set("val", "2")
    ET.SubElement(font0, _q("scheme")).set("val", "minor")
    for idx, style in style_items:
        font_el = ET.SubElement(fonts, _q("font"))
        if style.font:
            f = style.font
            ET.SubElement(font_el, _q("sz")).set("val", str(f.size if f.size else 11))
            if f.color:
                ET.SubElement(font_el, _q("color")).set("rgb", f.color)
            else:
                ET.SubElement(font_el, _q("color")).set("theme", "1")
            ET.SubElement(font_el, _q("name")).set("val", f.name if f.name else "Calibri")
            ET.SubElement(font_el, _q("family")).set("val", "2")
            if f.bold:
                ET.SubElement(font_el, _q("b"))
            if f.italic:
                ET.SubElement(font_el, _q("i"))
            if f.underline:
                ET.SubElement(font_el, _q("u"))
        else:
            ET.SubElement(font_el, _q("sz")).set("val", "11")
            ET.SubElement(font_el, _q("color")).set("theme", "1")
            ET.SubElement(font_el, _q("name")).set("val", "Calibri")
            ET.SubElement(font_el, _q("family")).set("val", "2")

    # fills: index 0 = none, index 1 = gray125 (Excel 默认要求)
    fills = ET.SubElement(root, _q("fills"))
    fills.set("count", str(len(style_items) + 2))
    pf0 = ET.SubElement(ET.SubElement(fills, _q("fill")), _q("patternFill"))
    pf0.set("patternType", "none")
    pf1 = ET.SubElement(ET.SubElement(fills, _q("fill")), _q("patternFill"))
    pf1.set("patternType", "gray125")
    for idx, style in style_items:
        fill_el = ET.SubElement(fills, _q("fill"))
        pf = ET.SubElement(fill_el, _q("patternFill"))
        pf.set("patternType", style.fill.pattern_type if style.fill else "solid")
        if style.fill and style.fill.fg_color:
            ET.SubElement(pf, _q("fgColor")).set("rgb", style.fill.fg_color)

    # borders: index 0 = empty border
    borders = ET.SubElement(root, _q("borders"))
    borders.set("count", str(len(style_items) + 1))

    def _empty_border(parent):
        b = ET.SubElement(parent, _q("border"))
        ET.SubElement(b, _q("left"))
        ET.SubElement(b, _q("right"))
        ET.SubElement(b, _q("top"))
        ET.SubElement(b, _q("bottom"))
        ET.SubElement(b, _q("diagonal"))
        return b

    _empty_border(borders)
    for idx, style in style_items:
        if style.border:
            b_el = ET.SubElement(borders, _q("border"))
            b = style.border
            for side_name, side in (("left", b.left), ("right", b.right), ("top", b.top), ("bottom", b.bottom)):
                side_el = ET.SubElement(b_el, _q(side_name))
                if side and side.style:
                    side_el.set("style", side.style)
                    if side.color:
                        ET.SubElement(side_el, _q("color")).set("rgb", side.color)
            ET.SubElement(b_el, _q("diagonal"))
        else:
            _empty_border(borders)

    # cellStyleXfs
    cell_style_xfs = ET.SubElement(root, _q("cellStyleXfs"))
    cell_style_xfs.set("count", "1")
    csxf = ET.SubElement(cell_style_xfs, _q("xf"))
    csxf.set("numFmtId", "0")
    csxf.set("fontId", "0")
    csxf.set("fillId", "0")
    csxf.set("borderId", "0")
    csxf.set("xfId", "0")

    # cellXfs
    cell_xfs = ET.SubElement(root, _q("cellXfs"))
    cell_xfs.set("count", str(len(style_items) + 1))
    xf0 = ET.SubElement(cell_xfs, _q("xf"))
    xf0.set("numFmtId", "0")
    xf0.set("fontId", "0")
    xf0.set("fillId", "0")
    xf0.set("borderId", "0")
    xf0.set("xfId", "0")

    num_fmts = None
    for idx, style in style_items:
        xf = ET.SubElement(cell_xfs, _q("xf"))
        apply_font = style.font is not None
        apply_fill = style.fill is not None
        apply_border = style.border is not None
        apply_align = style.alignment is not None
        apply_numfmt = style.number_format is not None
        xf.set("numFmtId", "0")
        xf.set("fontId", str(idx))
        xf.set("fillId", str(idx + 1))
        xf.set("borderId", str(idx))
        xf.set("xfId", "0")
        if apply_font:
            xf.set("applyFont", "1")
        if apply_fill:
            xf.set("applyFill", "1")
        if apply_border:
            xf.set("applyBorder", "1")
        if apply_align:
            xf.set("applyAlignment", "1")
            al = ET.SubElement(xf, _q("alignment"))
            if style.alignment.horizontal:
                al.set("horizontal", style.alignment.horizontal)
            if style.alignment.vertical:
                al.set("vertical", style.alignment.vertical)
            if style.alignment.wrap_text:
                al.set("wrapText", "1")
            if getattr(style.alignment, "indent", None):
                al.set("indent", str(style.alignment.indent))
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
            xf.set("applyNumberFormat", "1")

    # cellStyles (after cellXfs)
    cell_styles = ET.SubElement(root, _q("cellStyles"))
    cell_styles.set("count", "1")
    cell_style = ET.SubElement(cell_styles, _q("cellStyle"))
    cell_style.set("name", "Normal")
    cell_style.set("xfId", "0")
    cell_style.set("builtinId", "0")

    # dxfs, tableStyles for better compatibility
    dxfs = ET.SubElement(root, _q("dxfs"))
    dxfs.set("count", "0")
    table_styles = ET.SubElement(root, _q("tableStyles"))
    table_styles.set("count", "0")
    table_styles.set("defaultTableStyle", "TableStyleMedium2")
    table_styles.set("defaultPivotStyle", "PivotStyleLight16")

    # numFmts must appear before fonts
    if num_fmts is not None:
        root.remove(num_fmts)
        root.insert(0, num_fmts)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _row_height_pt(ws, row: int) -> float:
    """获取行高（磅），默认 15。"""
    dim = ws._row_dims.get(row)
    if dim and dim.height is not None:
        return float(dim.height)
    return 15.0


def _col_width_chars(ws, col: int) -> float:
    """获取列宽（字符单位），默认 8.43。"""
    dim = ws._col_dims.get(col)
    if dim and dim.width is not None:
        return float(dim.width)
    return 8.43


def _col_width_px(chars: float) -> int:
    """列宽字符近似转像素（常用近似）。"""
    return max(0, int(chars * 7 + 5))


def _row_height_px(pt: float) -> int:
    """行高磅转像素（96 DPI）。"""
    return max(0, int(pt * 96 / 72))


def _anchor_range_pixels(ws, min_r: int, min_c: int, max_r: int, max_c: int) -> tuple[int, int]:
    """计算锚定区域的总像素宽高（用于 embed 模式）。"""
    w = sum(_col_width_px(_col_width_chars(ws, c)) for c in range(min_c, max_c + 1))
    h = sum(_row_height_px(_row_height_pt(ws, r)) for r in range(min_r, max_r + 1))
    return (max(1, w), max(1, h))


def _build_drawing_xml(ws, sheet_index: int, media_map: dict[int, str]) -> tuple[bytes, bytes]:
    """生成 drawingN.xml 和 drawingN.xml.rels。

    media_map: 图片对象 id(img) -> zip 内 media 路径（如 ../media/image1.png，相对 drawings）
    返回 (drawing_xml, rels_xml)。
    """
    # xdr 根（命名空间由 _register_xlsx_namespaces 统一注册）
    root = ET.Element(f"{{{_NS_XDR}}}wsDr")

    # rels
    rels_root = ET.Element("Relationships")
    rels_root.set("xmlns", _NS_PKG_REL)
    next_rid = 1
    img_rid: dict[int, str] = {}

    for i, img in enumerate(ws._images):
        iid = id(img)
        if iid not in img_rid:
            rid = f"rId{next_rid}"
            next_rid += 1
            img_rid[iid] = rid
            rel = ET.SubElement(rels_root, "Relationship")
            rel.set("Id", rid)
            rel.set("Type", _REL_IMAGE)
            # drawing 在 xl/drawings/，media 在 xl/media/
            rel.set("Target", f"../media/{media_map[iid]}")
        rid = img_rid[iid]

        if img.mode == "embed":
            anchor_el = ET.SubElement(root, f"{{{_NS_XDR}}}twoCellAnchor")
            anchor_el.set("editAs", "twoCell")
            min_r, min_c, max_r, max_c = range_to_index(img.anchor)
            # from: 左上角
            frm = ET.SubElement(anchor_el, f"{{{_NS_XDR}}}from")
            ET.SubElement(frm, f"{{{_NS_XDR}}}col").text = str(min_c)
            ET.SubElement(frm, f"{{{_NS_XDR}}}colOff").text = "0"
            ET.SubElement(frm, f"{{{_NS_XDR}}}row").text = str(min_r)
            ET.SubElement(frm, f"{{{_NS_XDR}}}rowOff").text = "0"
            # to: 右下角（下一单元格的 0 偏移，即覆盖到 max 单元格末尾）
            to = ET.SubElement(anchor_el, f"{{{_NS_XDR}}}to")
            ET.SubElement(to, f"{{{_NS_XDR}}}col").text = str(max_c + 1)
            ET.SubElement(to, f"{{{_NS_XDR}}}colOff").text = "0"
            ET.SubElement(to, f"{{{_NS_XDR}}}row").text = str(max_r + 1)
            ET.SubElement(to, f"{{{_NS_XDR}}}rowOff").text = "0"
            pic = _build_pic(img, rid, i + 1)
            anchor_el.append(pic)
            ET.SubElement(anchor_el, f"{{{_NS_XDR}}}clientData")
        else:
            anchor_el = ET.SubElement(root, f"{{{_NS_XDR}}}oneCellAnchor")
            # anchor 可能是区域形式？float 通常用单格
            if ":" in img.anchor:
                min_r, min_c, _, _ = range_to_index(img.anchor)
            else:
                from excelkit.address import cell_to_index

                min_r, min_c = cell_to_index(img.anchor)
            frm = ET.SubElement(anchor_el, f"{{{_NS_XDR}}}from")
            ET.SubElement(frm, f"{{{_NS_XDR}}}col").text = str(min_c)
            ET.SubElement(frm, f"{{{_NS_XDR}}}colOff").text = str(max(0, img.offset[0]) * EMU_PER_PIXEL)
            ET.SubElement(frm, f"{{{_NS_XDR}}}row").text = str(min_r)
            ET.SubElement(frm, f"{{{_NS_XDR}}}rowOff").text = str(max(0, img.offset[1]) * EMU_PER_PIXEL)
            ext = ET.SubElement(anchor_el, f"{{{_NS_XDR}}}ext")
            ext.set("cx", str(max(1, img.display_width) * EMU_PER_PIXEL))
            ext.set("cy", str(max(1, img.display_height) * EMU_PER_PIXEL))
            pic = _build_pic(img, rid, i + 1)
            anchor_el.append(pic)
            ET.SubElement(anchor_el, f"{{{_NS_XDR}}}clientData")

    drawing_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    rels_xml = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
    return drawing_xml, rels_xml


def _build_pic(img, rid: str, pic_id: int) -> ET.Element:
    """构建 xdr:pic 元素。"""
    pic = ET.Element(f"{{{_NS_XDR}}}pic")
    nv = ET.SubElement(pic, f"{{{_NS_XDR}}}nvPicPr")
    c_nv = ET.SubElement(nv, f"{{{_NS_XDR}}}cNvPr")
    c_nv.set("id", str(pic_id))
    c_nv.set("name", img.alt_text or f"Picture {pic_id}")
    if img.alt_text:
        c_nv.set("descr", img.alt_text)
    c_nv_pic = ET.SubElement(nv, f"{{{_NS_XDR}}}cNvPicPr")
    ET.SubElement(c_nv_pic, f"{{{_NS_A}}}picLocks").set("noChangeAspect", "1")

    blip_fill = ET.SubElement(pic, f"{{{_NS_XDR}}}blipFill")
    blip = ET.SubElement(blip_fill, f"{{{_NS_A}}}blip")
    blip.set(f"{{{_NS_R}}}embed", rid)
    stretch = ET.SubElement(blip_fill, f"{{{_NS_A}}}stretch")
    ET.SubElement(stretch, f"{{{_NS_A}}}fillRect")

    sp_pr = ET.SubElement(pic, f"{{{_NS_XDR}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{_NS_A}}}xfrm")
    ET.SubElement(xfrm, f"{{{_NS_A}}}off").set("x", "0")
    ET.SubElement(xfrm, f"{{{_NS_A}}}off").set("y", "0")
    ext = ET.SubElement(xfrm, f"{{{_NS_A}}}ext")
    ext.set("cx", str(max(1, img.display_width) * EMU_PER_PIXEL))
    ext.set("cy", str(max(1, img.display_height) * EMU_PER_PIXEL))
    prst = ET.SubElement(sp_pr, f"{{{_NS_A}}}prstGeom")
    prst.set("prst", "rect")
    ET.SubElement(prst, f"{{{_NS_A}}}avLst")
    return pic


def _build_sheet_xml(ws, styles_cache: dict, style_objects: dict, has_drawing: bool, is_active: bool = False) -> bytes:
    """生成单张工作表的 sheetN.xml 内容。"""
    root = ET.Element(_q("worksheet"))

    # 1. sheetPr（必须靠前）
    if ws._color:
        sheet_pr = ET.SubElement(root, _q("sheetPr"))
        tab_color = ET.SubElement(sheet_pr, _q("tabColor"))
        tab_color.set("rgb", ws._color if len(ws._color) == 8 else "FF" + ws._color)

    # 2. dimension（建议在 sheetFormatPr 前；覆盖单元格/合并/图片锚点范围）
    max_r = ws._max_row
    max_c = ws._max_column
    for mr0, mc0, mr1, mc1 in getattr(ws, "_merged", ()):
        if mr1 > max_r: max_r = mr1
        if mc1 > max_c: max_c = mc1
    for img in getattr(ws, "_images", ()):
        try:
            if img.mode == "embed" or ":" in img.anchor:
                ar1, ac1, ar2, ac2 = range_to_index(img.anchor)
                if ar2 > max_r: max_r = ar2
                if ac2 > max_c: max_c = ac2
            else:
                from excelkit.address import cell_to_index as _ci_img
                ar, ac = _ci_img(img.anchor)
                if ar > max_r: max_r = ar
                if ac > max_c: max_c = ac
        except Exception:
            pass
    if max_r >= 0 and max_c >= 0:
        dim = ET.SubElement(root, _q("dimension"))
        dim.set("ref", f"A1:{_cell_ref(max_r, max_c)}")


    # 3. sheetViews（冻结窗格 + 默认视图，提升 WPS/Excel 兼容性）
    freeze = getattr(ws, "_freeze_panes", None)
    views = ET.SubElement(root, _q("sheetViews"))
    sv = ET.SubElement(views, _q("sheetView"))
    sv.set("workbookViewId", "0")
    if is_active:
        sv.set("tabSelected", "1")
    if freeze:
        try:
            from excelkit.address import cell_to_index as _ci
            fr, fc = _ci(freeze)
            if fr > 0 or fc > 0:
                pane = ET.SubElement(sv, _q("pane"))
                if fc > 0:
                    pane.set("xSplit", str(fc))
                if fr > 0:
                    pane.set("ySplit", str(fr))
                pane.set("topLeftCell", freeze)
                if fr > 0 and fc > 0:
                    pane.set("activePane", "bottomRight")
                elif fr > 0:
                    pane.set("activePane", "bottomLeft")
                else:
                    pane.set("activePane", "topRight")
                pane.set("state", "frozen")
        except Exception:
            pass
    sel = ET.SubElement(sv, _q("selection"))
    sel.set("activeCell", "A1")
    sel.set("sqref", "A1")
    # 4. sheetFormatPr
    sheet_format = ET.SubElement(root, _q("sheetFormatPr"))
    sheet_format.set("defaultRowHeight", "15")

    # 5. sheetData
    sheet_data = ET.SubElement(root, _q("sheetData"))
    if max_r >= 0 and max_c >= 0:
        for r in range(max_r + 1):
            row_el = ET.SubElement(sheet_data, _q("row"))
            row_el.set("r", str(r + 1))
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
                    v_el.text = "" if val is None else _format_value(val)
                elif val is not None:
                    if isinstance(val, bool):
                        c_el.set("t", "b")
                        ET.SubElement(c_el, _q("v")).text = "1" if val else "0"
                    elif isinstance(val, (int, float)) and not isinstance(val, bool):
                        c_el.set("t", "n")
                        ET.SubElement(c_el, _q("v")).text = str(val)
                    else:
                        c_el.set("t", "inlineStr")
                        is_el = ET.SubElement(c_el, _q("is"))
                        t_el = ET.SubElement(is_el, _q("t"))
                        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        t_el.text = _format_value(val)
                else:
                    c_el.set("t", "inlineStr")
                    is_el = ET.SubElement(c_el, _q("is"))
                    t_el = ET.SubElement(is_el, _q("t"))
                    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_el.text = ""

    # 6. autoFilter（在 sheetData 之后、mergeCells 之前）
    af = getattr(ws, "_auto_filter_range", None)
    if af:
        ET.SubElement(root, _q("autoFilter")).set("ref", af)

    # 7. mergeCells
    if ws._merged:
        merge_cells = ET.SubElement(root, _q("mergeCells"))
        merge_cells.set("count", str(len(ws._merged)))
        for min_r, min_c, mr, mc in ws._merged:
            mc_el = ET.SubElement(merge_cells, _q("mergeCell"))
            mc_el.set("ref", f"{_cell_ref(min_r, min_c)}:{_cell_ref(mr, mc)}")

    # 8. pageMargins + pageSetup
    page = ws._page
    # 简化：始终可省略 pageMargins；有设置时写 pageSetup
    if page.orientation == "landscape" or page.print_area or page.fit_width:
        pg = ET.SubElement(root, _q("pageSetup"))
        pg.set("orientation", page.orientation)
        pg.set("fitToWidth", str(page.fit_width or 1))
        pg.set("fitToHeight", str(page.fit_height or 0))
        pg.set("paperSize", "9")

    # 9. drawing（靠后）
    if has_drawing:
        drawing = ET.SubElement(root, _q("drawing"))
        drawing.set(f"{{{_NS_R}}}id", "rId1")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def save_xlsx(workbook, filename: str, *, formulas: bool = False) -> None:
    """将工作簿写出为 XLSX 文件。"""
    styles_cache: dict = {}
    style_objects: dict = {}
    _register_xlsx_namespaces(drawing=any(getattr(ws, "_images", None) for ws in workbook.sheets))

    # 收集全局 media（去重可选，这里按对象顺序编号）
    media_entries: list[tuple[str, bytes, str]] = []  # (filename, data, mime)
    sheet_drawings: dict[int, tuple[bytes, bytes]] = {}  # sheet_index -> (drawing_xml, rels_xml)
    sheet_has_drawing: dict[int, bool] = {}
    media_counter = 0
    for si, ws in enumerate(workbook.sheets):
        if not ws._images:
            sheet_has_drawing[si] = False
            continue
        sheet_has_drawing[si] = True
        media_map: dict[int, str] = {}
        for img in ws._images:
            iid = id(img)
            if iid not in media_map:
                media_counter += 1
                fname = f"image{media_counter}.{img.ext}"
                media_map[iid] = fname
                media_entries.append((fname, img.data, img.mime))
        sheet_drawings[si] = _build_drawing_xml(ws, si, media_map)

    # 生成 sheet xml
    sheets_xml = []
    for si, ws in enumerate(workbook.sheets):
        sheets_xml.append((ws.name, _build_sheet_xml(ws, styles_cache, style_objects, sheet_has_drawing.get(si, False), is_active=(si == getattr(workbook, "_active_index", 0)))))

    styles_bytes = _build_styles(styles_cache, style_objects)

    # workbook.xml
    wb_root = ET.Element(_q("workbook"))
    book_views = ET.SubElement(wb_root, _q("bookViews"))
    ET.SubElement(book_views, _q("workbookView")).set("activeTab", str(min(getattr(workbook, "_active_index", 0), max(0, len(sheets_xml) - 1))))
    sheets_el = ET.SubElement(wb_root, _q("sheets"))
    for i, (name, _) in enumerate(sheets_xml):
        sheet_el = ET.SubElement(sheets_el, _q("sheet"))
        sheet_el.set("name", name)
        sheet_el.set("sheetId", str(i + 1))
        sheet_el.set(f"{{{_NS_R}}}id", f"rId{i + 1}")
    calc_pr = ET.SubElement(wb_root, _q("calcPr"))
    calc_pr.set("calcId", "171027")
    workbook_xml = ET.tostring(wb_root, encoding="utf-8", xml_declaration=True)

    # xl/_rels/workbook.xml.rels
    rels_root = ET.Element("Relationships")
    rels_root.set("xmlns", _NS_PKG_REL)
    for i, (name, _) in enumerate(sheets_xml):
        rel = ET.SubElement(rels_root, "Relationship")
        rel.set("Id", f"rId{i + 1}")
        rel.set("Type", _REL_WORKSHEET)
        rel.set("Target", f"worksheets/sheet{i + 1}.xml")
    styles_rid = f"rId{len(sheets_xml) + 1}"
    rel = ET.SubElement(rels_root, "Relationship")
    rel.set("Id", styles_rid)
    rel.set("Type", _REL_STYLES)
    rel.set("Target", "styles.xml")
    rels_xml = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

    # [Content_Types].xml
    ct_root = ET.Element("Types")
    ct_root.set("xmlns", _NS_CT)
    for ext, ct in [("rels", _CT_RELS), ("xml", _CT_XML)]:
        d = ET.SubElement(ct_root, "Default")
        d.set("Extension", ext)
        d.set("ContentType", ct)
    if any(m[2] == "image/png" for m in media_entries):
        d = ET.SubElement(ct_root, "Default")
        d.set("Extension", "png")
        d.set("ContentType", "image/png")
    if any(m[2] == "image/jpeg" for m in media_entries):
        for ext in ("jpeg", "jpg"):
            d = ET.SubElement(ct_root, "Default")
            d.set("Extension", ext)
            d.set("ContentType", "image/jpeg")
    for part, ct in [("/xl/workbook.xml", _CT_WB), ("/xl/styles.xml", _CT_STYLES)]:
        o = ET.SubElement(ct_root, "Override")
        o.set("PartName", part)
        o.set("ContentType", ct)
    for i in range(len(sheets_xml)):
        o = ET.SubElement(ct_root, "Override")
        o.set("PartName", f"/xl/worksheets/sheet{i + 1}.xml")
        o.set("ContentType", _CT_WS)
    for si in sheet_drawings:
        o = ET.SubElement(ct_root, "Override")
        o.set("PartName", f"/xl/drawings/drawing{si + 1}.xml")
        o.set("ContentType", _CT_DRAWING)
    content_types = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

    # _rels/.rels
    rels_dot = ET.Element("Relationships")
    rels_dot.set("xmlns", _NS_PKG_REL)
    rel = ET.SubElement(rels_dot, "Relationship")
    rel.set("Id", "rId1")
    rel.set("Type", _REL_OFFICE_DOC)
    rel.set("Target", "xl/workbook.xml")
    rels_dot_xml = ET.tostring(rels_dot, encoding="utf-8", xml_declaration=True)

    try:
        with zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("_rels/.rels", rels_dot_xml)
            zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
            zf.writestr("xl/workbook.xml", workbook_xml)
            zf.writestr("xl/styles.xml", styles_bytes)
            for i, (name, sheet_bytes) in enumerate(sheets_xml):
                zf.writestr(f"xl/worksheets/sheet{i + 1}.xml", sheet_bytes)
                if sheet_has_drawing.get(i):
                    drawing_xml, drawing_rels = sheet_drawings[i]
                    zf.writestr(f"xl/drawings/drawing{i + 1}.xml", drawing_xml)
                    zf.writestr(f"xl/drawings/_rels/drawing{i + 1}.xml.rels", drawing_rels)
                    sheet_rel = ET.Element("Relationships")
                    sheet_rel.set("xmlns", _NS_PKG_REL)
                    r = ET.SubElement(sheet_rel, "Relationship")
                    r.set("Id", "rId1")
                    r.set("Type", _REL_DRAWING)
                    r.set("Target", f"../drawings/drawing{i + 1}.xml")
                    zf.writestr(f"xl/worksheets/_rels/sheet{i + 1}.xml.rels", ET.tostring(sheet_rel, encoding="utf-8", xml_declaration=True))
            for fname, data, mime in media_entries:
                zf.writestr(f"xl/media/{fname}", data)
    except Exception as exc:
        raise InvalidFileError(f"无法写出 XLSX 文件：{exc}") from exc


__all__ = ["save_xlsx"]
