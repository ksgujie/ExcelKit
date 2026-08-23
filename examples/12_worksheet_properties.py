"""工作表重命名和标签颜色示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：演示重命名工作表、标签查询同步和 XLSX 标签颜色保存。

    使用方法：在项目根目录执行 ``python -m examples.12_worksheet_properties``。
    参数：无。
    返回：``None``；生成带蓝色工作表标签的 ``12_worksheet_properties.xlsx``。
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.label = "销售明细"
    worksheet.label_color = "4472C4"
    worksheet.append_rows([
        ["产品", "数量"],
        ["产品 A", 2],
        ["产品 B", 3],
    ])

    assert workbook.sheet("销售明细") is worksheet
    assert worksheet.label_color == "FF4472C4"
    output = Path("12_worksheet_properties.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
