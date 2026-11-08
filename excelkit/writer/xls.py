"""XLS 写出器：依赖 xlwt，未安装时抛出清晰错误。"""


def save_xls(workbook, filename: str) -> None:
    """将工作簿写出为 XLS 文件。"""
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
    book.save(filename)


def _xl_write_value(val):
    from datetime import date, datetime

    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return val
    if isinstance(val, bool):
        return val
    return val


__all__ = ["save_xls"]
