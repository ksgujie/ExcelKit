"""高级 XlsxWriter 直接写出示例。"""

from pathlib import Path

from excelkit import Workbook
from excelkit.writer.xlsx import XlsxWriter


def main() -> None:
    """功能：绕过 Workbook.save，直接使用唯一高级写出器生成 XLSX。

    使用方法：在项目根目录执行 ``python -m examples.07_xlsx_writer``。
    参数：无。
    返回：``None``；在当前目录生成 ``07_xlsx_writer.xlsx``。
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["项目", "值"])
    worksheet.append(["版本", "0.1.2"])

    output = Path("07_xlsx_writer.xlsx")
    XlsxWriter(workbook).write(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
