"""CSV/TSV 读取器：基于标准库 csv，无外部依赖。"""

import csv


def load_delimited(filename: str, *, encoding: str | None = None):
    from excelkit.core.workbook import Workbook
    """读取 CSV/TSV 文件，按扩展名或内容判断分隔符。"""
    import os

    ext = os.path.splitext(filename)[1].lower()
    delimiter = "\t" if ext == ".tsv" else ","
    enc = encoding or "utf-8"
    workbook = Workbook()
    ws = workbook.active
    ws.name = "Sheet1"
    with open(filename, "r", newline="", encoding=enc) as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row_vals in reader:
            converted = []
            for v in row_vals:
                converted.append(_convert(v))
            ws.append(converted)
    return workbook


def _convert(text: str):
    if text == "" or text is None:
        return None
    # 日期字面量 #YYYY-M-D
    if text.startswith("#"):
        return text
    try:
        if "." in text or "e" in text.lower():
            return float(text)
        return int(text)
    except ValueError:
        pass
    return text


__all__ = ["load_delimited"]
