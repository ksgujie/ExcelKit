"""Workbook.load 使用的唯一内部格式分派入口。"""

from __future__ import annotations

import os
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING, Type, TypeVar

from ..errors import InvalidFileError
from .delimited import _load_delimited
from .xls import _load_xls
from .xlsx import _load_xlsx

if TYPE_CHECKING:
    from ..core.workbook import Workbook

_WorkbookType = TypeVar("_WorkbookType", bound="Workbook")
_OPEN_XML_SUFFIXES = {".xlsx", ".xlsm", ".xltx"}
_DELIMITED_SUFFIXES = {".csv", ".tsv"}


def _load_workbook(
    workbook_class: Type[_WorkbookType], filename: os.PathLike | str, *,
    encoding: str | None = None, delimiter: str | None = None,
    has_header: bool = False,
) -> _WorkbookType:
    """功能：按扩展名和文件结构分派到唯一对应读取实现。

    使用方法：仅由 ``Workbook.load(filename)`` 类方法调用。
    参数：``workbook_class`` 为 Workbook 类或子类；``filename`` 为字符串或
    ``os.PathLike`` 路径。
    返回：从 XLS、XLSX、XLSM、XLTX、CSV 或 TSV 构造的新工作簿。
    异常：类型错误时抛出 ``TypeError``；其他不支持格式抛出
    :class:`InvalidFileError`；文件不存在时透传 ``FileNotFoundError``。
    """
    if not isinstance(filename, (str, os.PathLike)):
        raise TypeError("filename 必须是字符串或 PathLike 对象")
    path = Path(filename)
    suffix = path.suffix.lower()
    if suffix in _DELIMITED_SUFFIXES:
        return _load_delimited(workbook_class, path, encoding=encoding,
                               delimiter=delimiter, has_header=has_header)
    if encoding is not None or delimiter is not None or has_header:
        raise ValueError("encoding、delimiter、has_header 仅适用于 CSV/TSV 文件")
    if suffix == ".xls":
        return _load_xls(workbook_class, path)
    if suffix in _OPEN_XML_SUFFIXES:
        return _load_xlsx(workbook_class, path)
    if not path.exists():
        raise FileNotFoundError(path)
    if zipfile.is_zipfile(path):
        return _load_xlsx(workbook_class, path)
    raise InvalidFileError(f"不支持的工作簿文件格式：{suffix or '无扩展名'}")


__all__ = []
