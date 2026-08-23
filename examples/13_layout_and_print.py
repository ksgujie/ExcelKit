"""工作表结构、布局和打印设置完整示例。"""

from pathlib import Path

from excelkit import Workbook
from excelkit.page_setup import HeaderFooter, PageMargins


def create_report() -> Workbook:
    """功能：创建包含合并、尺寸、冻结、筛选和打印设置的销售报表。

    使用方法：``workbook = create_report()``。
    参数：无。
    返回：已经完成页面布局的 :class:`Workbook`。
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.name = "销售报表"
    worksheet.color = "4472C4"

    worksheet.range("A1:F1").merge()
    worksheet["A1"] = "2026 年销售报表"
    worksheet.append(["序号", "产品", "数量", "单价", "金额", "备注"])
    worksheet.append_rows([
        [1, "产品 A", 2, 19.5, 39, "正常"],
        [2, "产品 B", 3, 8, 24, "正常"],
        [3, "产品 C", 1, 120, 120, "重点"],
    ])

    worksheet.row(0).height = 28
    worksheet.column(0).width = 10
    worksheet.column(1).width = 24
    worksheet.column(2).width = 12
    worksheet.column(3).width = 14
    worksheet.column(4).width = 14
    worksheet.column(5).width = 20
    worksheet.freeze_panes = "A3"
    worksheet.auto_filter_range = "A2:F5"
    worksheet.show_gridlines = False

    page = worksheet.page
    page.orientation = "landscape"
    page.paper_size = "A4"
    page.fit(width=1)
    page.print_area = "A1:F5"
    page.repeat_rows = (0, 1)
    page.center_horizontal = True
    page.print_gridlines = False
    page.print_headings = False
    page.margins = PageMargins(
        left=1.5,
        right=1.5,
        top=2.0,
        bottom=2.0,
        header=0.8,
        footer=0.8,
    )
    page.header = HeaderFooter(center="2026 年销售报表", right="&D")
    page.footer = HeaderFooter(
        left="ExcelKit",
        center="第 &P 页，共 &N 页",
        right="&F",
    )
    return workbook


def main() -> None:
    """功能：生成可直接检查布局和打印效果的XLSX与XLS示例文件。

    使用方法：在项目根目录执行 ``python -m examples.13_layout_and_print``。
    参数：无。
    返回：``None``；生成 ``13_layout_and_print.xlsx`` 和同名 ``.xls``。
    """
    workbook = create_report()
    xlsx_file = Path("13_layout_and_print.xlsx")
    xls_file = Path("13_layout_and_print.xls")
    workbook.save(xlsx_file)
    workbook.save(xls_file)
    print(f"XLSX：{xlsx_file.resolve()}")
    print(f"XLS：{xls_file.resolve()}")


if __name__ == "__main__":
    main()
