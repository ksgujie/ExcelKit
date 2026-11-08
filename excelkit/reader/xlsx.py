"""XLSX 读取器：基于标准库 zipfile/xml 解析 OpenXML 格式，无外部依赖。"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from excelkit.core.cell import Cell, CellValue
from excelkit.core.worksheet import Worksheet
from excelkit.errors import InvalidFileError

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _parse_bool(text: str | None, default: bool = False) -> bool:
    if text is None:
        return default
    return text.strip().lower() in ("1", "true", "yes")


def _read_shared_strings(zip_file: ZipFile) -> list[str]:
    try:
        with zip_file.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
    except KeyError:
        return []
    strings = []
    for si in tree.getroot().iter(f"{_NS}si"):
        text_parts = []
        for t in si.iter(f"{_NS}t"):
            text_parts.append(t.text or "")
        strings.append("".join(text_parts))
    return strings


def _cell_value(raw: str | None, cell_type: str | None, shared: list[str]) -> object:
    if raw is None:
        return None
    if cell_type == "s":
        idx = int(raw)
        if 0 <= idx < len(shared):
            return shared[idx]
        return raw
    if cell_type in ("b", "str"):
        return raw.lower() == "1" if cell_type == "b" else raw
    if cell_type in ("n", "d"):
        try:
            if "." in raw or "e" in raw.lower():
                return float(raw)
            return int(raw)
        except ValueError:
            return raw
    return raw


def load_xlsx(filename: str) -> Workbook:
    """读取 XLSX 文件并返回 Workbook。"""
    try:
        zip_file = ZipFile(filename)
    except Exception as exc:
        raise InvalidFileError(f"无法打开 XLSX 文件：{exc}") from exc

    try:
        # 解析工作簿结构
        with zip_file.open("xl/workbook.xml") as f:
            wb_tree = ET.parse(f)
        wb_root = wb_tree.getroot()

        sheets_el = wb_root.find(f"{_NS}sheets")
        sheet_nodes = list(sheets_el) if sheets_el is not None else []

        # 读取共享字符串
        shared = _read_shared_strings(zip_file)

        # 关系映射：sheet -> 目标文件
        try:
            with zip_file.open("xl/_rels/workbook.xml.rels") as f:
                rels_tree = ET.parse(f)
            rels = {}
            for rel in rels_tree.getroot():
                rid = rel.get("Id")
                target = rel.get("Target")
                if target and target.startswith("worksheets/"):
                    rels[rid] = target
        except KeyError:
            rels = {}

        # r:id -> sheet 名称与索引
        sheet_order = []
        for sheet in sheet_nodes:
            name = sheet.get("name")
            rid = sheet.get(f"{{http://schemas.openxmlformats.org/officeDocument/2006/relationships}}id")
            target = rels.get(rid, f"xl/worksheets/sheet{len(sheet_order)+1}.xml")
            sheet_order.append((name or f"Sheet{len(sheet_order)+1}", target))

        from excelkit.core.workbook import Workbook

        workbook = Workbook()
        # 移除默认 Sheet1，按实际顺序创建
        workbook._sheets.clear()
        workbook._name_index.clear()

        for name, target in sheet_order:
            ws = workbook.add_sheet(name)
            sheet_xml = target
            if not sheet_xml.startswith("xl/"):
                sheet_xml = "xl/" + sheet_xml.lstrip("/")
            try:
                with zip_file.open(sheet_xml) as f:
                    sheet_tree = ET.parse(f)
            except KeyError:
                continue
            sheet_root = sheet_tree.getroot()

            dim = sheet_root.find(f"{_NS}dimension")
            if dim is not None:
                ref = dim.get("ref")
                if ref:
                    # 触及边界（不直接设值，只记录）
                    pass

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
                    col_str = re.match(r"([A-Za-z]+)", cell_ref).group(1)
                    col_idx = 0
                    for ch in col_str.upper():
                        col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
                    col_idx -= 1
                    cell_type = cell.get("t")
                    raw = None
                    v = cell.find(f"{_NS}v")
                    if v is not None:
                        raw = v.text
                    # inlineStr 格式：值在 <is><t> 中，而非 <v>
                    if raw is None:
                        is_el = cell.find(f"{_NS}inlineStr")
                        if is_el is not None:
                            t_el = is_el.find(f"{_NS}t")
                            if t_el is not None:
                                raw = t_el.text
                    formula = None
                    f = cell.find(f"{_NS}f")
                    if f is not None:
                        formula = f.text or ""
                    val = _cell_value(raw, cell_type, shared)
                    ws._set_cell_value(row_idx, col_idx, val, formula=formula)

            # 合并单元格
            for mc in sheet_root.iter(f"{_NS}mergeCell"):
                ref = mc.get("ref")
                if ref:
                    parts = ref.split(":")
                    if len(parts) == 2:
                        from excelkit.address import cell_index
                        r1, c1 = cell_index(parts[0])
                        r2, c2 = cell_index(parts[1])
                        ws._merge_cells(r1, c1, r2, c2)

        return workbook
    finally:
        zip_file.close()


__all__ = ["load_xlsx"]
