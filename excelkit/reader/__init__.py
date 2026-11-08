"""ExcelKit 读取器：统一 load_workbook 入口，内部按格式分派。"""

from excelkit.errors import InvalidFileError
from excelkit.reader.delimited import load_delimited
from excelkit.reader.xls import load_xls
from excelkit.reader.xlsx import load_xlsx

__all__ = ["load_delimited", "load_xls", "load_xlsx", "load_workbook"]


def load_workbook(filename: str, *, encoding: str | None = None):
    """按文件扩展名自动分派读取器。

    支持 .xlsx、.xlsm、.xltx（按 XLSX 解析）、.xls（按 XLS 解析）、
    .csv、.tsv（按分隔文本解析）。
    """
    import os

    ext = os.path.splitext(filename)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xltx"):
        return load_xlsx(filename)
    if ext == ".xls":
        return load_xls(filename)
    if ext in (".csv", ".tsv"):
        return load_delimited(filename, encoding=encoding)
    raise InvalidFileError(f"不支持的文件扩展名：{ext}")
