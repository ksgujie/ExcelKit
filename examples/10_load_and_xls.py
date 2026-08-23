"""Workbook.load 与 XLS/XLSX 读写示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：生成 XLS 和 XLSX，再通过统一类方法读取并打印数据。

    使用方法：在项目根目录执行 ``python -m examples.10_load_and_xls``。
    参数：无。
    返回：``None``；在当前目录生成两个示例工作簿并打印读取结果。
    """
    source = Workbook()
    source.active.append_rows([
        ["姓名", "时间", "成绩"],
        ["张三", "#2026-8-1 12:33", 95],
    ])
    xlsx_file = Path("10_load.xlsx")
    xls_file = Path("10_load.xls")
    source.save(xlsx_file)
    source.save(xls_file)

    print("XLSX：", Workbook.load(xlsx_file).active.values)
    print("XLS：", Workbook.load(xls_file).active.values)


if __name__ == "__main__":
    main()
