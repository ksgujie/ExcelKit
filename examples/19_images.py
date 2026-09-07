"""演示图片单元格/区域锚定、填充策略和图片查询。"""

from __future__ import annotations

import base64

from excelkit import ImageFit, ImagePlacement, Workbook


def main() -> None:
    """功能：创建包含浮动图片和随区域缩放图片的示例工作簿。

    使用方法：在项目根目录运行 ``python examples/19_images.py``。
    参数：无；示例使用内置的 1×1 PNG 数据，不依赖第三方图片库。
    返回：``None``；在当前目录生成 ``19_images.xlsx``。
    """
    payload = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
    )
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.add_image(payload, anchor="A1", name="pixel.png")
    worksheet.add_image(
        payload,
        anchor="B2:F8",
        name="pixel.png",
        placement=ImagePlacement.CELL,
        fit=ImageFit.CONTAIN,
    )
    print("区域图片：", worksheet.image("B2").anchor)
    workbook.save("19_images.xlsx")


if __name__ == "__main__":
    main()
