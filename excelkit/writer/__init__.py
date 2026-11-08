"""ExcelKit 写出器：统一 save_workbook 入口，内部按格式分派。"""

from excelkit.errors import InvalidFileError
from excelkit.writer.delimited import save_delimited
from excelkit.writer.xls import save_xls
from excelkit.writer.xlsx import save_xlsx

__all__ = ["save_delimited", "save_xls", "save_xlsx", "save_workbook"]


def save_workbook(
    workbook, filename: str, *, encoding: str | None = None, formulas: bool = False
) -> None:
    """按文件扩展名自动分派写出器。

    支持 .xlsx、.xlsm、.xltx（按 XLSX 写出）、.xls（按 XLS 写出）、
    .csv、.tsv（仅单工作表，按分隔文本写出）。
    """
    import os

    ext = os.path.splitext(filename)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xltx"):
        save_xlsx(workbook, filename, formulas=formulas)
    elif ext == ".xls":
        save_xls(workbook, filename)
    elif ext in (".csv", ".tsv"):
        save_delimited(workbook, filename, encoding=encoding)
    else:
        raise InvalidFileError(f"不支持的文件扩展名：{ext}")
