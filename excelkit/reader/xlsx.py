"""XLSX 读取器：基于标准库 zipfile/xml 解析 OpenXML 格式，无外部依赖。

支持读取：单元格值/公式、合并单元格、工作表标签颜色、PNG/JPEG 图片（浮动/嵌入）。
"""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from excelkit.address import cell_to_index, index_to_range, range_to_index
from excelkit.core.image import EMU_PER_PIXEL, Image, _detect_image
from excelkit.errors import InvalidAddressError, InvalidFileError

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _read_xml(zip_file: ZipFile, name: str):
    with zip_file.open(name) as f:
        return ET.parse(f)


def _read_rels(zip_file: ZipFile, rels_path: str) -> dict[str, str]:
    """读取 .rels 文件，返回 rId -> Target 映射。"""
    try:
        tree = _read_xml(zip_file, rels_path)
    except KeyError:
        return {}
    rels = {}
    for rel in tree.getroot():
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rels[rid] = target
    return rels


def _resolve_part(base_part: str, target: str) -> str:
    """根据 Target 解析 zip 内部件路径。"""
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_part)
    joined = posixpath.normpath(posixpath.join(base_dir, target))
    return joined.lstrip("/")


def _read_shared_strings(zip_file: ZipFile) -> list[str]:
    try:
        tree = _read_xml(zip_file, "xl/sharedStrings.xml")
    except KeyError:
        return []
    strings = []
    for si in tree.getroot().iter(f"{_NS}si"):
        parts = []
        for t in si.iter(f"{_NS}t"):
            parts.append(t.text or "")
        strings.append("".join(parts))
    return strings


def _cell_value(raw: str | None, cell_type: str | None, shared: list[str]) -> object:
    if raw is None:
        return None
    if cell_type == "s":
        try:
            idx = int(raw)
            if 0 <= idx < len(shared):
                return shared[idx]
        except ValueError:
            pass
        return raw
    if cell_type == "b":
        return raw in ("1", "true", "TRUE")
    if cell_type in ("str", "inlineStr"):
        return raw
    # n, d, e, or default
    try:
        if "." in raw or "e" in raw.lower() or "E" in raw:
            return float(raw)
        return int(raw)
    except (ValueError, TypeError):
        return raw


def _parse_marker(marker) -> tuple[int, int, int, int]:
    """解析 xdr:from/to 标记，返回 (col, colOff, row, rowOff)。"""
    def _int(tag):
        el = marker.find(f"{_XDR}{tag}")
        if el is None or el.text is None:
            return 0
        try:
            return int(el.text)
        except ValueError:
            return 0
    return (_int("col"), _int("colOff"), _int("row"), _int("rowOff"))


def _load_sheet_images(zip_file: ZipFile, sheet_part: str, sheet_root) -> list[Image]:
    """从工作表解析图片。"""
    drawing_el = sheet_root.find(f"{_NS}drawing")
    if drawing_el is None:
        return []
    rid = drawing_el.get(f"{_NS_R}id") or drawing_el.get(f"{_R}id")
    if not rid:
        return []
    sheet_rels_path = posixpath.join("xl/worksheets/_rels", posixpath.basename(sheet_part) + ".rels")
    sheet_rels = _read_rels(zip_file, sheet_rels_path)
    drawing_target = sheet_rels.get(rid)
    if not drawing_target:
        return []
    drawing_part = _resolve_part(sheet_part, drawing_target)
    try:
        drawing_root = _read_xml(zip_file, drawing_part).getroot()
    except KeyError:
        return []
    drawing_rels_path = posixpath.join(posixpath.dirname(drawing_part), "_rels", posixpath.basename(drawing_part) + ".rels")
    drawing_rels = _read_rels(zip_file, drawing_rels_path)

    images: list[Image] = []
    for anchor in list(drawing_root):
        tag = anchor.tag
        mode = "float"
        anchor_ref = "A1"
        off = (0, 0)
        w_px = None
        h_px = None
        if tag == f"{_XDR}oneCellAnchor":
            frm = anchor.find(f"{_XDR}from")
            ext = anchor.find(f"{_XDR}ext")
            if frm is not None:
                col, col_off, row, row_off = _parse_marker(frm)
                from excelkit.address import index_to_cell as _ca

                anchor_ref = _ca(row, col)
                off = (max(0, round(col_off / EMU_PER_PIXEL)), max(0, round(row_off / EMU_PER_PIXEL)))
            if ext is not None:
                cx = ext.get("cx")
                cy = ext.get("cy")
                if cx:
                    w_px = max(1, round(int(cx) / EMU_PER_PIXEL))
                if cy:
                    h_px = max(1, round(int(cy) / EMU_PER_PIXEL))
            mode = "float"
        elif tag == f"{_XDR}twoCellAnchor":
            edit_as = (anchor.get("editAs") or "").strip()
            frm = anchor.find(f"{_XDR}from")
            to = anchor.find(f"{_XDR}to")
            if frm is not None and to is not None:
                c1, _, r1, _ = _parse_marker(frm)
                c2, co2, r2, ro2 = _parse_marker(to)
                # to 常指向下一个边界；转换为包含的单元格范围
                max_r = r2 - 1 if (co2 == 0 and ro2 == 0 and r2 > r1) else r2
                max_c = c2 - 1 if (co2 == 0 and ro2 == 0 and c2 > c1) else c2
                max_r = max(r1, max_r)
                max_c = max(c1, max_c)
                anchor_ref = index_to_range(r1, c1, max_r, max_c)
            mode = "embed" if edit_as == "twoCell" else "float"
        else:
            continue

        pic = anchor.find(f"{_XDR}pic")
        if pic is None:
            continue
        blip = pic.find(f"{_XDR}blipFill/{_A}blip")
        if blip is None:
            continue
        embed = blip.get(f"{_R}embed") or blip.get(f"{_NS_R}embed")
        if not embed:
            continue
        target = drawing_rels.get(embed)
        if not target:
            continue
        media_part = _resolve_part(drawing_part, target)
        try:
            data = zip_file.read(media_part)
        except KeyError:
            continue
        try:
            mime, ext, nat_w, nat_h = _detect_image(data)
        except (ValueError, InvalidFileError, KeyError, IndexError):
            continue
        alt = None
        c_nv = pic.find(f"{_XDR}nvPicPr/{_XDR}cNvPr")
        if c_nv is not None:
            alt = c_nv.get("descr") or c_nv.get("name")
            if alt and alt.startswith("Picture "):
                alt = c_nv.get("descr")

        if mode == "float":
            if w_px is None:
                w_px = nat_w
            if h_px is None:
                h_px = nat_h
        else:
            w_px, h_px = None, None
        images.append(
            Image(
                anchor=anchor_ref,
                data=data,
                mime=mime,
                ext=ext,
                natural_width=nat_w,
                natural_height=nat_h,
                mode=mode,
                width_px=w_px,
                height_px=h_px,
                offset=off,
                alt_text=alt,
            )
        )
    return images


def load_xlsx(filename: str) -> "Workbook":
    """读取 XLSX 文件并返回 Workbook。"""
    try:
        zip_file = ZipFile(filename)
    except Exception as exc:
        raise InvalidFileError(f"无法打开 XLSX 文件：{exc}") from exc

    try:
        wb_tree = _read_xml(zip_file, "xl/workbook.xml")
        wb_root = wb_tree.getroot()
        sheets_el = wb_root.find(f"{_NS}sheets")
        sheet_nodes = list(sheets_el) if sheets_el is not None else []

        shared = _read_shared_strings(zip_file)
        rels = _read_rels(zip_file, "xl/_rels/workbook.xml.rels")

        sheet_order = []
        for sheet in sheet_nodes:
            name = sheet.get("name") or f"Sheet{len(sheet_order) + 1}"
            rid = sheet.get(f"{_NS_R}id") or sheet.get(f"{_R}id")
            target = rels.get(rid, f"worksheets/sheet{len(sheet_order) + 1}.xml")
            sheet_part = _resolve_part("xl/workbook.xml", target)
            sheet_order.append((name, sheet_part))

        from excelkit.core.workbook import Workbook

        # 创建空工作簿容器（不通过 __init__ 生成默认 Sheet1，避免加载时的实例/状态干扰）
        workbook = Workbook.__new__(Workbook)
        workbook._sheets = []
        workbook._name_index = {}
        workbook._active_index = 0

        for name, sheet_part in sheet_order:
            ws = workbook.add_sheet(name)
            try:
                sheet_tree = _read_xml(zip_file, sheet_part)
            except KeyError:
                continue
            sheet_root = sheet_tree.getroot()

            # 标签颜色
            tab_color_el = sheet_root.find(f"{_NS}sheetPr/{_NS}tabColor")
            if tab_color_el is not None:
                rgb = tab_color_el.get("rgb")
                if rgb:
                    try:
                        ws.tab_color = rgb
                    except (ValueError, InvalidAddressError):
                        pass

            # 冻结窗格
            pane = sheet_root.find(f"{_NS}sheetViews/{_NS}sheetView/{_NS}pane")
            if pane is not None:
                try:
                    top_left = pane.get("topLeftCell")
                    x_split = int(pane.get("xSplit") or "0")
                    y_split = int(pane.get("ySplit") or "0")
                    if top_left and (x_split > 0 or y_split > 0):
                        ws.freeze_panes = top_left
                except (ValueError, InvalidAddressError, TypeError):
                    pass

            # 自动筛选
            af = sheet_root.find(f"{_NS}autoFilter")
            if af is not None:
                ref = af.get("ref")
                if ref:
                    try:
                        ws.auto_filter.range = ref
                    except (ValueError, InvalidAddressError):
                        pass

            # 单元格
            for row in sheet_root.iter(f"{_NS}row"):
                r_attr = row.get("r")
                if r_attr is None:
                    continue
                row_idx = int(r_attr) - 1
                for cell in row:
                    if cell.tag != f"{_NS}c":
                        continue
                    cell_ref = cell.get("r")
                    if cell_ref is None:
                        continue
                    m = re.match(r"([A-Za-z]+)", cell_ref)
                    if not m:
                        continue
                    col_str = m.group(1).upper()
                    col_idx = 0
                    for ch in col_str:
                        col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
                    col_idx -= 1
                    cell_type = cell.get("t")
                    raw = None
                    v = cell.find(f"{_NS}v")
                    if v is not None:
                        raw = v.text
                    if raw is None:
                        # 标准 inline string 使用 <is>；兼容旧版本写出的 <inlineStr>
                        is_el = cell.find(f"{_NS}is")
                        if is_el is None:
                            is_el = cell.find(f"{_NS}inlineStr")
                        if is_el is not None:
                            t_el = is_el.find(f"{_NS}t")
                            if t_el is not None:
                                raw = t_el.text or ""
                    formula = None
                    f = cell.find(f"{_NS}f")
                    if f is not None:
                        formula = "=" + (f.text or "")
                    val = _cell_value(raw, cell_type, shared)
                    ws._set_cell_value(row_idx, col_idx, val, formula=formula)

            # 合并（捕获数据异常，暴露代码错误）
            for mc in sheet_root.iter(f"{_NS}mergeCell"):
                ref = mc.get("ref")
                if ref and ":" in ref:
                    try:
                        r1, c1, r2, c2 = range_to_index(ref)
                        ws._merge_cells(r1, c1, r2, c2)
                    except (ValueError, KeyError, IndexError, InvalidAddressError):
                        pass

            # 图片（捕获数据/文件异常，暴露代码错误）
            try:
                ws._images = _load_sheet_images(zip_file, sheet_part, sheet_root)
            except (ValueError, KeyError, IndexError, InvalidAddressError, InvalidFileError, ET.ParseError, OSError):
                ws._images = []

        return workbook
    finally:
        zip_file.close()


__all__ = ["load_xlsx"]
