"""CSV 和 TSV 分隔文本工作簿的内部读取实现。"""

from __future__ import annotations

import csv
import io
import os
from pathlib import Path
from typing import TYPE_CHECKING, Type, TypeVar

from ..errors import InvalidFileError

if TYPE_CHECKING:
    from ..core.workbook import Workbook

_WorkbookType = TypeVar("_WorkbookType", bound="Workbook")


def _decode_text(payload: bytes) -> str:
    """功能：使用常见表格文本编码解码文件内容。

    使用方法：CSV/TSV 主读取函数读取字节后内部调用。
    参数：``payload`` 为文件原始字节；按 UTF-8 BOM、UTF-8、GB18030 顺序尝试。
    返回：解码后的 Unicode 字符串。
    异常：所有支持编码均失败时抛出 :class:`InvalidFileError`。
    """
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise InvalidFileError("文本表格不是支持的 UTF-8 或 GB18030 编码")


def _detect_delimiter(text: str, suffix: str) -> str:
    """功能：确定分隔文本文件使用的字段分隔符。

    使用方法：读取文本内容后由主函数内部调用。
    参数：``text`` 为已解码内容；``suffix`` 为小写扩展名，``.tsv`` 固定使用
    制表符，``.csv`` 在逗号、分号、制表符和竖线中检测。
    返回：单字符分隔符字符串。
    """
    if suffix == ".tsv":
        return "\t"
    try:
        return csv.Sniffer().sniff(text[:8192], delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


def _load_delimited(
    workbook_class: Type[_WorkbookType], filename: os.PathLike | str
) -> _WorkbookType:
    """功能：把 CSV 或 TSV 文件读取为单工作表 Workbook。

    使用方法：仅由 ``Workbook.load(filename)`` 根据扩展名分派调用。
    参数：``workbook_class`` 为 Workbook 类或子类；``filename`` 为 CSV/TSV 路径。
    返回：包含名为 ``Sheet1`` 的新工作簿；字段保留为字符串，带 ``#`` 的日期
    字面量仍会经过统一普通值转换。
    异常：文件不存在时透传 ``FileNotFoundError``；编码或 CSV 结构无法读取时抛出
    :class:`InvalidFileError`。
    """
    path = Path(filename)
    text = _decode_text(path.read_bytes())
    delimiter = _detect_delimiter(text, path.suffix.lower())
    workbook = workbook_class()
    worksheet = workbook.add_sheet("Sheet1")
    try:
        rows = csv.reader(io.StringIO(text, newline=""), delimiter=delimiter)
        for row_index, row_values in enumerate(rows):
            if not row_values:
                worksheet._touch(row_index, 0)
                continue
            for column_index, value in enumerate(row_values):
                worksheet._set_value(row_index, column_index, value)
    except (csv.Error, ValueError) as error:
        raise InvalidFileError(f"无法读取分隔文本文件：{path}") from error
    return workbook
