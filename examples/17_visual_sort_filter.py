"""演示图表、图片、批注、排序、筛选与工作表可见性。"""

from __future__ import annotations

import base64
from pathlib import Path

from excelkit import Workbook, Worksheet
from excelkit.chart import ChartType
from excelkit.note import Note
from excelkit.sort import SortKey

_PIXEL_PNG = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/"
    "9kL7UAAAAABJRU5ErkJggg=="
)


def write_demo_image(filename: str | Path) -> Path:
    """功能：创建示例使用的最小 PNG 图片文件。

    使用方法：``image_file = write_demo_image('pixel.png')``。
    参数：``filename`` 为 PNG 目标路径。
    返回：写入后的 ``Path`` 对象。
    异常：父目录不可写时透传文件系统异常。
    """
    target = Path(filename)
    target.write_bytes(base64.b64decode(_PIXEL_PNG))
    return target


def main() -> None:
    """功能：生成包含可视化和数据整理功能的 XLSX 示例文件。

    使用方法：在项目根目录运行 ``python -m examples.17_visual_sort_filter``。
    参数：无。
    返回：``None``；生成 ``17_visual_sort_filter.xlsx`` 与 ``pixel.png``。
    """
    workbook = Workbook()
    worksheet = workbook.add_sheet("销售数据")
    worksheet.append_rows([
        ["月份", "销售额", "状态"],
        ["一月", 120, "通过"],
        ["二月", 180, "通过"],
        ["三月", 90, "待复核"],
    ])

    worksheet.sort("A2:C4", keys=[SortKey(1, descending=True)])
    worksheet.auto_filter.range = "A1:C4"
    worksheet.auto_filter.add(2, ["通过"]).apply()
    worksheet["A2"].note = Note("当前销售额最高", author="销售部")

    chart = worksheet.add_chart(ChartType.COLUMN, anchor="E2")
    chart.title = "月度销售额"
    chart.add_series(values="B2:B4", categories="A2:A4", name="销售额")

    image_file = write_demo_image("pixel.png")
    worksheet.add_image(image_file, anchor="E18").alt_text = "示例像素图"

    internal = workbook.add_sheet("内部说明")
    internal["A1"] = "此工作表以 VERY_HIDDEN 保存。"
    internal.visibility = Worksheet.VERY_HIDDEN
    workbook.save("17_visual_sort_filter.xlsx")


if __name__ == "__main__":
    main()
