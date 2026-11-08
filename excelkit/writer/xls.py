"""XLS 写出器：依赖 xlwt；图片写入不支持（xlwt 限制）。"""

from excelkit.errors import UnsupportedFeatureError


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
