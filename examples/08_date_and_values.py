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
    saved_day = worksheet.cell("A4").set_value("2026-8-3").as_date()
    worksheet.cell("B4").set_value("125")
    read_number = worksheet["B4"].read().as_int()

    assert worksheet["A2"].value == date(2026, 8, 1)
    assert worksheet["B2"].value == datetime(2026, 8, 1, 12, 33)
    assert saved_day == date(2026, 8, 3)
    assert worksheet["A4"].value == date(2026, 8, 3)
    assert read_number == 125
    assert worksheet["B4"].value == "125"  # read() 转换不会写回。
    print(worksheet.values)


if __name__ == "__main__":
    main()
