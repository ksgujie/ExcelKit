"""单元格字体、填充、边框、对齐和数字格式示例。"""

from pathlib import Path

from excelkit import Workbook
from excelkit.style import Alignment, Border, BorderSide, Fill, Font, Style


def main() -> None:
    """功能：创建带组合样式的 XLSX 工作簿。

    使用方法：在项目根目录执行 ``python -m examples.09_style``。
    参数：无。
    返回：``None``；在当前目录生成 ``09_style.xlsx``。
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["姓名", "成绩"])
    worksheet.append(["张三", 95.5])

    title_style = Style(
        font=Font(name="微软雅黑", size=12, bold=True, color="FFFFFF"),
        fill=Fill("4472C4"),
        border=Border(bottom=BorderSide(Border.THIN, "000000")),
        alignment=Alignment(
            horizontal=Alignment.HORIZONTAL_CENTER,
            vertical=Alignment.VERTICAL_CENTER,
            wrap_text=True,
        ),
    )
    number_style = Style(number_format="0.00")
    for column in range(2):
        worksheet.cell(0, column).style = title_style
    worksheet["B2"].style = number_style

    output = Path("09_style.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
