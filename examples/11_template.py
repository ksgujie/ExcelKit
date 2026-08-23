"""Excel 标量标签和循环行块模板渲染示例。"""

from pathlib import Path

from excelkit import Workbook
from excelkit.style import Alignment, Fill, Font, Style


def create_template(filename: Path) -> None:
    """功能：创建一个可用 Excel 打开和编辑的多工作表销售报表模板。

    使用方法：``create_template(Path('销售模板.xlsx'))``。
    参数：``filename`` 为待生成的 XLSX 模板路径。
    返回：``None``；模板包含明细与汇总表、标量标签、循环、样式和公式。
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.name = "销售明细"
    worksheet["A1"] = "{title}"
    worksheet["A2"] = "客户：{customer.name}"
    worksheet.append(["序号", "产品", "数量", "单价", "金额", "显示金额"])
    worksheet["A4"] = "{loop items}"
    worksheet["A5"] = "{items.@index + 1}"
    worksheet["B5"] = "{items.name}"
    worksheet["C5"] = "{items.quantity}"
    worksheet["D5"] = "{items.price}"
    worksheet["E5"] = "{items.quantity * items.price}"
    worksheet["F5"] = '￥{items.quantity * items.price | format:",.2f"}'
    worksheet["A6"] = "{/loop}"
    worksheet["A7"] = "制表人：{operator}"

    title_style = Style(
        font=Font(name="微软雅黑", size=14, bold=True, color="FFFFFF"),
        fill=Fill("4472C4"),
        alignment=Alignment(horizontal="center", vertical="center"),
    )
    row_style = Style(alignment=Alignment(vertical="center"))
    worksheet["A1"].style = title_style
    for column in range(6):
        worksheet.cell(2, column).style = Style(font=Font(bold=True))
        worksheet.cell(4, column).style = row_style
    worksheet["D5"].style = Style(
        alignment=Alignment(vertical="center"), number_format="#,##0.00"
    )
    worksheet["E5"].style = Style(
        alignment=Alignment(vertical="center"), number_format="#,##0.00"
    )

    summary = workbook.add_sheet("汇总")
    summary["A1"] = "公司：{company}"
    summary["A2"] = "合计：{total}"
    summary["A3"] = "备注：{note}"
    workbook.save(filename)


def main() -> None:
    """功能：创建 Excel 模板，以公共和分工作表数据生成最终报表。

    使用方法：在项目根目录执行 ``python -m examples.11_template``。
    参数：无。
    返回：``None``；生成 ``11_template.xlsx`` 和 ``11_template_result.xlsx``。
    """
    template_file = Path("11_template.xlsx")
    output_file = Path("11_template_result.xlsx")
    create_template(template_file)

    shared_data = {"company": "示例公司"}
    sheet_data = {
        "销售明细": {
            "title": "2026 年 8 月销售明细",
            "customer": {"name": "示例公司"},
            "operator": "张三",
            "items": [
                {"name": "产品 A", "quantity": 2, "price": 19.5},
                {"name": "产品 B", "quantity": 3, "price": 8},
                {"name": "产品 C", "quantity": 1, "price": 120},
            ],
        },
        "汇总": {"total": 183},
    }
    Workbook.load(template_file).render(
        shared_data, sheet_data=sheet_data
    ).save(output_file)
    print(f"模板：{template_file.resolve()}")
    print(f"结果：{output_file.resolve()}")


if __name__ == "__main__":
    main()
