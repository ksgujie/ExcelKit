"""CSV/TSV 写出器：基于标准库 csv，仅导出单工作表。"""

import csv


def save_delimited(workbook, filename: str, *, encoding: str | None = None) -> None:
    """将活动工作表导出为 CSV/TSV 文件。"""
    import os

    ext = os.path.splitext(filename)[1].lower()
    delimiter = "\t" if ext == ".tsv" else ","
    enc = encoding or "utf-8"
    ws = workbook.active
    with open(filename, "w", newline="", encoding=enc) as f:
        writer = csv.writer(f, delimiter=delimiter)
        for row_vals in ws.values:
            writer.writerow(["" if v is None else v for v in row_vals])


__all__ = ["save_delimited"]
