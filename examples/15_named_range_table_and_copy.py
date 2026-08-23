"""命名区域、Excel数据表、样式复制和区域复制完整示例。"""

from pathlib import Path

from excelkit import Workbook
from excelkit.style import Alignment, Fill, Font, Style


def create_table_workbook() -> Workbook:
    """功能：创建带命名区域、数据表以及复制格式的销售工作簿。

    使用方法：``workbook = create_table_workbook()``。
    参数：无。
    返回：包含全部0.4.0区域组织功能的 :class:`Workbook`。
    """
    workbook = Workbook()
    worksheet = workbook.add_sheet("销售明细")
    worksheet.append_rows([
        ["产品", "数量", "单价", "金额"],
        ["产品 A", 2, 19.5, None],
        ["产品 B", 3, 8, None],
    ])
    worksheet["D2"].formula = "=B2*C2"
    worksheet["D3"].formula = "=B3*C3"

    title_style = Style(
        font=Font(bold=True, color="FFFFFF"),
        fill=Fill("4472C4"),
        alignment=Alignment(horizontal=Alignment.HORIZONTAL_CENTER),
    )
    worksheet["A1"].style = title_style
    for column in range(1, 4):
        worksheet.cell(0, column).copy_style(worksheet["A1"])

    workbook.add_named_range("SalesAmount", worksheet.range("D2:D3"))
    worksheet.add_table(
        "A1:D3",
        name="SalesTable",
        style="TableStyleMedium9",
        show_row_stripes=True,
    )

    # 复制公式和样式到一个同尺寸区域；相对行列引用会自动平移。
    worksheet.range("A1:D3").copy_to(worksheet.range("F1:I3"))
    workbook.calculate()
    return workbook


def main() -> None:
    """功能：保存命名区域与数据表 XLSX 示例文件。

    使用方法：在项目根目录执行
    ``python -m examples.15_named_range_table_and_copy``。
    参数：无。
    返回：``None``；生成 ``15_named_range_table_and_copy.xlsx``。
    """
    output = Path("15_named_range_table_and_copy.xlsx")
    create_table_workbook().save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
