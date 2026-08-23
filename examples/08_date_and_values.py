"""日期、日期时间字面量与 Worksheet.values 示例。"""

from datetime import date, datetime

from excelkit import Workbook


def main() -> None:
    """功能：演示自动日期转换并读取整张工作表的普通值。

    使用方法：在项目根目录执行 ``python -m examples.08_date_and_values``。
    参数：无。
    返回：``None``；在终端打印转换后的二维数据。
    """
    worksheet = Workbook().active
    worksheet.append(["日期", "日期时间"])
    worksheet.append(["#2026-8-1", "#2026/8/1 12:33"])
    worksheet.append(["#2026/8/2", "#2026-8-2 08:05:30"])

    assert worksheet["A2"].value == date(2026, 8, 1)
    assert worksheet["B2"].value == datetime(2026, 8, 1, 12, 33)
    print(worksheet.values)


if __name__ == "__main__":
    main()
