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
    return workbook


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
