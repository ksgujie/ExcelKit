"""XLS 读取器：依赖 xlrd，未安装时抛出清晰错误。"""


def load_xls(filename: str):
    """读取 XLS 文件并返回 Workbook。"""
    try:
        import xlrd  # noqa: F401
    except ImportError as exc:
        raise ImportError(
            "读取 XLS 文件需要安装 xlrd：pip install xlrd"
        ) from exc

    import xlrd

    book = xlrd.open_workbook(filename)

    from excelkit.core.workbook import Workbook

    workbook = Workbook()
    workbook._sheets.clear()
    workbook._name_index.clear()

    for sheet_idx, sh in enumerate(book.sheets()):
        name = sh.name
        ws = workbook.add_sheet(name) if sheet_idx > 0 else workbook.active
        if sheet_idx == 0:
            ws.name = name
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                val = _xl_cell_value(cell, book.datemode)
                ws._set_cell_value(r, c, val, formula=None)

    _apply_auto_filters(workbook, book)
    return workbook


def _apply_auto_filters(workbook, xls_book) -> None:
    """从 ``_FilterDatabase`` 内置定义名还原每个工作表的自动筛选区域。"""
    names = getattr(xls_book, "name_map", {}).get("_filterdatabase")
    if not names:
        return
    if not isinstance(names, (list, tuple)):
        names = [names]

    from excelkit.address import index_to_cell, index_to_range

    for name in names:
        try:
            if getattr(name, "scope", None) is None:
                continue
            result = name.result
            coords = result.value[0].coords
        except (AttributeError, IndexError, TypeError):
            continue
        if not coords or len(coords) != 6:
            continue
        sht_lo, sht_hi, row_lo, row_hi, col_lo, col_hi = coords
        # 只支持单个工作表的引用；跨表引用跳过
        if sht_hi - sht_lo != 1:
            continue
        if not (0 <= sht_lo < len(workbook.sheets)):
            continue
        if row_hi <= row_lo or col_hi <= col_lo:
            continue
        ws = workbook.sheets[sht_lo]
        if row_hi - 1 == row_lo and col_hi - 1 == col_lo:
            # 单格区域用 "A1" 表示，与 XLSX 读取结果保持一致
            ws._auto_filter_range = index_to_cell(row_lo, col_lo)
        else:
            ws._auto_filter_range = index_to_range(
                row_lo, col_lo, row_hi - 1, col_hi - 1
            )


def _xl_cell_value(cell, datemode: int):
    """把 xlrd 单元格转为 Python 原生值。"""
    import xlrd

    if cell.ctype == xlrd.XL_CELL_EMPTY:
        return None
    if cell.ctype == xlrd.XL_CELL_TEXT:
        return cell.value
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        val = cell.value
        if float(val).is_integer():
            return int(val)
        return val
    if cell.ctype == xlrd.XL_CELL_DATE:
        from datetime import datetime

        return xlrd.xldate_as_datetime(cell.value, datemode)
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return None
    return cell.value


__all__ = ["load_xls"]
