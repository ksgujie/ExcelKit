"""CSV 与 TSV 分隔文本的内部写出实现。"""

from __future__ import annotations

import csv
import os
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any

from ..errors import InvalidFileError

if TYPE_CHECKING:
    from ..core.worksheet import Worksheet


def _resolve_delimiter(path: Path, delimiter: str | None) -> str:
    """功能：根据文件扩展名或显式参数确定文本列分隔符。

    使用方法：由 :func:`write_delimited` 在打开目标文件前内部调用。
    参数：``path`` 为目标路径；``delimiter`` 为可选的单字符分隔符。省略时 ``.csv``
    使用逗号、``.tsv`` 使用制表符。
    返回：合法的单字符分隔符。
    异常：扩展名不受支持或分隔符类型、长度无效时抛出 ``InvalidFileError`` 或
    ``ValueError``。
    """
    if delimiter is not None:
        if not isinstance(delimiter, str) or len(delimiter) != 1:
            raise ValueError("delimiter 必须是单字符字符串或 None")
        return delimiter
    if path.suffix.lower() == ".csv":
        return ","
    if path.suffix.lower() == ".tsv":
        return "\t"
    raise InvalidFileError("分隔文本只能保存为 .csv 或 .tsv")


def _text_value(value: Any) -> Any:
    """功能：将单元格值转换为 csv 模块可稳定写出的文本字段。

    使用方法：由 :func:`write_delimited` 为每个单元格内部调用。
    参数：``value`` 为普通值或公式缓存值。
    返回：空值返回空字符串；日期时间使用 ISO 文本；其余值保持原样交由 ``csv``
    模块处理。
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return value


def write_delimited(
    worksheet: "Worksheet",
    filename: str | os.PathLike[str],
    *,
    encoding: str = "utf-8-sig",
    delimiter: str | None = None,
    formulas: bool = False,
) -> None:
    """功能：将一张工作表导出为 CSV 或 TSV 文件。

    使用方法：由 ``Worksheet.export()`` 和单工作表 ``Workbook.save()`` 调用。
    参数：``worksheet`` 为源工作表；``filename`` 必须以 ``.csv`` 或 ``.tsv``
    结尾；``encoding`` 为非空 Python 编码名；``delimiter`` 可覆盖扩展名默认值；
    ``formulas`` 为真时导出公式文本，否则导出公式缓存值。
    返回：``None``；文件内容按 0-based 行列索引从 A1 到已触及的最大边界写出。
    异常：参数、格式或编码无效时抛出 ``TypeError``、``ValueError`` 或
    ``InvalidFileError``；路径不可写时透传文件系统异常。
    """
    if not isinstance(filename, (str, os.PathLike)):
        raise TypeError("filename 必须是字符串或 PathLike 对象")
    if not isinstance(encoding, str) or not encoding:
        raise TypeError("encoding 必须是非空字符串")
    if not isinstance(formulas, bool):
        raise TypeError("formulas 必须是 bool")
    path = Path(filename)
    separator = _resolve_delimiter(path, delimiter)
    try:
        with path.open("w", encoding=encoding, newline="") as handle:
            writer = csv.writer(handle, delimiter=separator, lineterminator="\n")
            for row in range(worksheet.max_row + 1):
                values: list[Any] = []
                for column in range(worksheet.max_column + 1):
                    coordinate = (row, column)
                    if coordinate in worksheet._formulas:
                        value = (
                            worksheet._formulas[coordinate]
                            if formulas
                            else worksheet._formula_values.get(coordinate)
                        )
                    else:
                        value = worksheet._values.get(row, column)
                    values.append(_text_value(value))
                writer.writerow(values)
    except LookupError as error:
        raise ValueError(f"不支持的文本编码：{encoding!r}") from error


__all__ = []
