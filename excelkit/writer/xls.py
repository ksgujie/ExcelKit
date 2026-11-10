"""XLS 写出器：依赖 xlwt；图片写入不支持（xlwt 限制）。

自动筛选说明
------------
XLS（BIFF8）的自动筛选不是 xlwt 的公开能力，需要直接写入两类二进制记录：

* 工作表子流：``AUTOFILTERINFO``（0x009D），记录筛选区域的列数；
* 工作簿全局流：``SUPBOOK``（0x01AE）+ ``EXTERNSHEET``（0x0017）+
  ``NAME``（0x0018，内置名 0x0D ``_FilterDatabase``），记录精确的筛选区域。

只写 ``AUTOFILTERINFO`` 时 Excel 会按“当前使用区域”自行推断范围，容易出错；
因此这里同时写入 ``_FilterDatabase`` 定义名，保证范围与
``ws.auto_filter.range`` 一致。注入失败时回退为“不带筛选保存”，绝不损坏文件。
"""

import struct
import warnings

from excelkit.address import XLS_MAX_COLUMNS, XLS_MAX_ROWS, range_to_index
from excelkit.errors import UnsupportedFeatureError

# ---- BIFF8 记录号 ----
_SID_EOF = 0x000A
_SID_BOUNDSHEET = 0x0085
_SID_BOF = 0x0809
_SID_SST = 0x00FC
_SID_DIMENSIONS = 0x0200
_SID_AUTOFILTERINFO = 0x009D
_SID_NAME = 0x0018
_SID_SUPBOOK = 0x01AE
_SID_EXTERNSHEET = 0x0017

# NAME 记录的 grbit：fBuiltin(0x20) | fHidden(0x01)
_NAME_GRBIT_FILTER_DATABASE = 0x0021
# 内置定义名 _FilterDatabase 的编号
_BUILTIN_FILTER_DATABASE = 0x0D
# ptgArea3d：三维区域引用操作码
_PTG_AREA3D = 0x3B


def save_xls(workbook, filename: str) -> None:
    """将工作簿写出为 XLS 文件。若工作簿包含图片则抛出 UnsupportedFeatureError。"""
    # 检查图片：xlwt 不支持图片写入，避免静默丢失
    for ws in workbook.sheets:
        if getattr(ws, "_images", None):
            raise UnsupportedFeatureError(
                "XLS 格式的图片写入暂不支持（底层 xlwt 限制）。请改用 .xlsx 保存；"
                "或移除图片后再保存为 .xls。"
            )
    try:
        import xlwt  # noqa: F401
    except ImportError as exc:
        raise ImportError(
            "写出 XLS 文件需要安装 xlwt：pip install xlwt"
        ) from exc

    import xlwt

    book = xlwt.Workbook()
    for ws in workbook.sheets:
        sheet = book.add_sheet(ws.name)
        for r in range(ws.max_row + 1):
            for c in range(ws.max_column + 1):
                val = ws._get_cell_value(r, c)
                if val is None:
                    continue
                sheet.write(r, c, _xl_write_value(val))

    filters = _collect_filters(workbook)
    if filters:
        try:
            data = _build_biff_data(book, filters)
        except Exception as exc:  # noqa: BLE001 - 任何注入失败都回退，避免产出损坏文件
            warnings.warn(
                f"XLS 自动筛选写入失败，已改为保存不带筛选的文件：{exc}",
                RuntimeWarning,
                stacklevel=2,
            )
            data = book.get_biff_data()
    else:
        data = book.get_biff_data()

    # 直接用 xlwt 的复合文档写入器保存已经修补好的 BIFF 流
    from xlwt.CompoundDoc import XlsDoc

    XlsDoc().save(filename, data)


def _xl_write_value(val):
    from datetime import date, datetime

    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return val
    if isinstance(val, bool):
        return val
    return val


def _collect_filters(workbook):
    """收集每个工作表的自动筛选区域。

    返回 ``{工作表索引: (min_row, min_col, max_row, max_col)}``，行列均为 0-based。
    区域无效或超出 XLS 上限时跳过并给出警告。
    """
    filters = {}
    for index, ws in enumerate(workbook.sheets):
        address = getattr(ws, "_auto_filter_range", None)
        if not address:
            continue
        try:
            min_row, min_col, max_row, max_col = range_to_index(address)
        except Exception as exc:  # noqa: BLE001 - 地址非法时不阻断保存
            warnings.warn(
                f"工作表 {ws.name!r} 的自动筛选区域无效，已跳过：{exc}",
                RuntimeWarning,
                stacklevel=3,
            )
            continue
        if max_row >= XLS_MAX_ROWS or max_col >= XLS_MAX_COLUMNS:
            warnings.warn(
                f"工作表 {ws.name!r} 的自动筛选区域超出 XLS 上限"
                f"（{XLS_MAX_ROWS} 行 / {XLS_MAX_COLUMNS} 列），已跳过。",
                RuntimeWarning,
                stacklevel=3,
            )
            continue
        filters[index] = (min_row, min_col, max_row, max_col)
    return filters


def _record(sid: int, body: bytes) -> bytes:
    """按 BIFF8 格式拼装一条记录：记录号(2) + 长度(2) + 数据区。"""
    return struct.pack("<HH", sid, len(body)) + body


def _iter_records(data: bytes, start: int = 0, end: int = None):
    """按 BIFF8 记录格式遍历 ``data[start:end]``。

    逐个产出 ``(记录起始偏移, 记录号, 数据区长度)``；遇到长度越界的残缺记录即停止。
    """
    end = len(data) if end is None else end
    pos = start
    while pos + 4 <= end:
        sid, length = struct.unpack_from("<HH", data, pos)
        if pos + 4 + length > end:
            break
        yield pos, sid, length
        pos += 4 + length


def _split_substreams(data: bytes):
    """切分 BIFF 流：``[(起始, 结束), ...]``，第 0 个是全局流，其余是工作表子流。"""
    streams = []
    current_start = None
    for pos, sid, length in _iter_records(data):
        if sid == _SID_BOF and current_start is None:
            current_start = pos
        elif sid == _SID_EOF and current_start is not None:
            streams.append((current_start, pos + 4))
            current_start = None
    return streams


def _find_record(data: bytes, start: int, end: int, target_sid: int):
    """在 ``[start, end)`` 内查找指定记录号，返回记录起始偏移；找不到返回 None。"""
    for pos, sid, _length in _iter_records(data, start, end):
        if sid == target_sid:
            return pos
    return None


def _build_globals_records(sheet_count: int, filters) -> bytes:
    """构造全局流中的 SUPBOOK + EXTERNSHEET + 每个筛选表的 NAME 记录。"""
    parts = []

    # 内部 3D 引用用的 SUPBOOK：工作表数量 + BIFF5 遗留字节 0x01 0x04
    parts.append(_record(_SID_SUPBOOK, struct.pack("<HBB", sheet_count, 0x01, 0x04)))

    # EXTERNSHEET：为每个工作表登记一条 (iSupBook=0, itabFirst, itabLast)
    externsheet = struct.pack("<H", sheet_count)
    for index in range(sheet_count):
        externsheet += struct.pack("<HHH", 0, index, index)
    parts.append(_record(_SID_EXTERNSHEET, externsheet))

    # 每个启用了筛选的工作表写一条内置定义名 _FilterDatabase
    for index in sorted(filters):
        min_row, min_col, max_row, max_col = filters[index]
        # ptgArea3d：ixti(2) + rwFirst(2) + rwLast(2) + colFirst(2) + colLast(2)
        rgce = struct.pack(
            "<BHHHHH", _PTG_AREA3D, index, min_row, max_row, min_col, max_col
        )
        body = bytearray()
        body += struct.pack("<H", _NAME_GRBIT_FILTER_DATABASE)  # grbit
        body += b"\x00"                                        # chKey
        body += b"\x01"                                        # cch：内置名长度 1
        body += struct.pack("<H", len(rgce))                    # cce
        body += struct.pack("<H", 0)                            # ixals
        body += struct.pack("<H", index + 1)                    # itab：1-based 工作表索引
        body += bytes(4)                                        # 菜单/描述/帮助/状态栏长度
        body += bytes((0x00, _BUILTIN_FILTER_DATABASE))         # 非 Unicode + 内置名编号
        body += rgce
        parts.append(_record(_SID_NAME, bytes(body)))

    return b"".join(parts)


def _patch_boundsheets(data: bytes) -> bytes:
    """按最终的流位置回填每条 BOUNDSHEET 记录的 lbPlyPos（工作表 BOF 绝对偏移）。"""
    streams = _split_substreams(data)
    if not streams:
        return data
    sheet_offsets = [start for start, _end in streams[1:]]
    global_start, global_end = streams[0]

    out = bytearray(data)
    index = 0
    for pos, sid, _length in _iter_records(data, global_start, global_end):
        if sid != _SID_BOUNDSHEET:
            continue
        if index >= len(sheet_offsets):
            break
        struct.pack_into("<I", out, pos + 4, sheet_offsets[index])
        index += 1
    return bytes(out)


def _build_biff_data(book, filters) -> bytes:
    """生成 BIFF 流并注入自动筛选记录。

    处理顺序：
    1. 从后往前给每个工作表插入 AUTOFILTERINFO（避免偏移互相影响）；
    2. 在全局流 SST 前插入 SUPBOOK/EXTERNSHEET/NAME；
    3. 重新计算并回填所有 BOUNDSHEET 的工作表偏移。
    """
    data = book.get_biff_data()
    streams = _split_substreams(data)
    if len(streams) < 2:
        raise ValueError("未找到工作表子流")

    # 1. 工作表子流：DIMENSIONS 前插入 AUTOFILTERINFO（列数）
    for sheet_index in sorted(filters, reverse=True):
        start, end = streams[sheet_index + 1]
        insert_at = _find_record(data, start, end, _SID_DIMENSIONS)
        if insert_at is None:
            insert_at = _find_record(data, start, end, _SID_EOF)
        if insert_at is None:
            raise ValueError("工作表缺少 DIMENSIONS/EOF 记录")
        min_row, min_col, max_row, max_col = filters[sheet_index]
        count = max_col - min_col + 1
        data = (
            data[:insert_at]
            + _record(_SID_AUTOFILTERINFO, struct.pack("<H", count))
            + data[insert_at:]
        )

    # 2. 全局流：SST 前插入 SUPBOOK/EXTERNSHEET/NAME
    global_start, global_end = streams[0]
    insert_at = _find_record(data, global_start, global_end, _SID_SST)
    if insert_at is None:
        insert_at = _find_record(data, global_start, global_end, _SID_EOF)
    if insert_at is None:
        raise ValueError("全局流缺少 SST/EOF 记录")
    data = (
        data[:insert_at]
        + _build_globals_records(len(streams) - 1, filters)
        + data[insert_at:]
    )

    # 3. 所有插入完成后统一回填工作表偏移
    return _patch_boundsheets(data)


__all__ = ["save_xls"]
