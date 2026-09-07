"""基础 XLSX 图片对象。"""

from __future__ import annotations

import os
import struct
from pathlib import Path
from typing import TYPE_CHECKING

from .address import cell_index, range_address, range_index

if TYPE_CHECKING:
    from .core.worksheet import Worksheet


def _image_size(payload: bytes) -> tuple[str, int, int]:
    """功能：读取 PNG 或 JPEG 图片格式及像素尺寸。

    使用方法：由 ``Worksheet.add_image()`` 内部调用。
    参数：``payload`` 为 PNG 或 JPEG 的完整二进制内容。
    返回：``(extension, width, height)``，扩展名为 ``png`` 或 ``jpeg``。
    异常：文件不存在、格式不支持或图片尺寸无效时抛出 ``FileNotFoundError`` 或 ``ValueError``。
    """
    if payload.startswith(b"\x89PNG\r\n\x1a\n") and len(payload) >= 24:
        width, height = struct.unpack(">II", payload[16:24])
        if width and height:
            return "png", width, height
    if payload.startswith(b"\xff\xd8"):
        position = 2
        while position + 9 < len(payload):
            if payload[position] != 0xFF:
                position += 1
                continue
            marker = payload[position + 1]
            position += 2
            if marker in {0xD8, 0xD9}:
                continue
            length = int.from_bytes(payload[position:position + 2], "big")
            if length < 2 or position + length > len(payload):
                break
            if marker in set(range(0xC0, 0xC4)) | set(range(0xC5, 0xC8)) | set(range(0xC9, 0xCC)) | set(range(0xCD, 0xD0)):
                height = int.from_bytes(payload[position + 3:position + 5], "big")
                width = int.from_bytes(payload[position + 5:position + 7], "big")
                if width and height:
                    return "jpeg", width, height
            position += length
    raise ValueError("仅支持具有有效尺寸的 PNG 或 JPEG 图片")


class ImagePlacement:
    """提供图片锚定方式常量，便于 IDE 补全。

    使用方法：作为 ``Worksheet.add_image(..., placement=...)`` 的参数，或读取
    ``Image.placement`` 进行比较。``FLOATING`` 固定像素大小，``CELL`` 随锚点区域
    移动并缩放。
    参数：无实例参数；类属性是字符串常量。
    返回：``FLOATING`` 和 ``CELL`` 两个常量值。
    """

    FLOATING = "floating"
    CELL = "cell"


class ImageFit:
    """提供图片填充区域方式常量，便于 IDE 补全。

    使用方法：作为 ``Worksheet.add_image(..., fit=...)`` 的参数，或读取
    ``Image.fit`` 进行比较。三个常量分别代表拉伸、等比完整显示和等比铺满裁剪。
    参数：无实例参数；类属性是字符串常量。
    返回：``STRETCH``、``CONTAIN`` 和 ``COVER`` 三个常量值。
    """

    STRETCH = "stretch"
    CONTAIN = "contain"
    COVER = "cover"


class Image:
    """表示锚定在工作表内的 PNG 或 JPEG 图片。"""

    __slots__ = (
        "_worksheet", "filename", "_payload", "_anchor", "_bounds",
        "_placement", "_fit", "format", "width", "height", "offset_x",
        "offset_y", "alt_text",
    )

    def __init__(
        self,
        worksheet: "Worksheet",
        source: os.PathLike[str] | str | bytes,
        anchor: str,
        *,
        name: str | None = None,
        placement: str = ImagePlacement.FLOATING,
        fit: str = ImageFit.STRETCH,
    ) -> None:
        """功能：创建已验证文件和锚点的图片对象。

        使用方法：由 ``Worksheet.add_image()`` 创建，不建议直接调用。
        参数：``worksheet`` 为所属工作表；``source`` 为 PNG/JPEG 文件路径或
        二进制内容；``anchor`` 为图片左上角 A1 地址；二进制内容必须提供带扩展名
        的 ``name``。
        返回：无。
        异常：图片不存在、格式不支持或尺寸无效时抛出 ``FileNotFoundError`` 或
        ``ValueError``。
        """
        if isinstance(source, bytes):
            if not isinstance(name, str) or not name:
                raise ValueError("二进制图片必须提供非空 name")
            payload = source
            filename = name
        elif isinstance(source, (str, os.PathLike)):
            payload = Path(source).read_bytes()
            filename = str(Path(source))
        else:
            raise TypeError("source 必须是图片路径或 bytes")
        image_format, width, height = _image_size(payload)
        self._worksheet = worksheet
        self.filename = str(filename)
        self._payload = payload
        self._anchor = ""
        self._bounds = (0, 0, 0, 0)
        self._placement = ImagePlacement.FLOATING
        self._fit = ImageFit.STRETCH
        self.placement = placement
        self.fit = fit
        self.anchor = anchor
        self.format = image_format
        self.width = width
        self.height = height
        self.offset_x = 0
        self.offset_y = 0
        self.alt_text = ""

    @property
    def anchor(self) -> str:
        """功能：取得图片的规范化 A1 锚点地址。

        使用方法：``address = image.anchor``；修改时可赋值给同名属性。
        参数：无；这是读写属性。
        返回：单元格地址（例如 ``"B2"``）或包含首尾单元格的矩形区域地址
        （例如 ``"B2:F8"``）。地址中的行列均按 Excel A1 规则表示。
        """
        return self._anchor

    @anchor.setter
    def anchor(self, value: str) -> None:
        """功能：设置图片锚点并规范化其边界。

        使用方法：``image.anchor = "B2:F8"``；浮动图片只能设置单个单元格。
        参数：``value`` 为单格或矩形 A1 地址，区域首行、首列必须不晚于末行、
        末列。地址不区分大小写，保存为大写规范形式。
        返回：``None``。
        异常：地址格式或边界无效时抛出 ``InvalidAddressError``；浮动图片使用
        多单元格区域时抛出 ``ValueError``。
        """
        if not isinstance(value, str):
            raise TypeError("anchor 必须是单元格或矩形区域地址字符串")
        try:
            bounds = range_index(value)
        except ValueError:
            row, column = cell_index(value)
            bounds = (row, column, row, column)
        if (
            hasattr(self, "_placement")
            and self._placement == ImagePlacement.FLOATING
            and bounds[0:2] != bounds[2:4]
        ):
            raise ValueError("浮动图片只能锚定到单个单元格；区域请使用 ImagePlacement.CELL")
        self._bounds = bounds
        self._anchor = range_address(*bounds)

    @property
    def bounds(self) -> tuple[int, int, int, int]:
        """功能：取得图片锚点的 0-based 矩形边界。

        使用方法：``min_row, min_col, max_row, max_col = image.bounds``。
        参数：无。
        返回：``(最小行, 最小列, 最大行, 最大列)``，顺序固定为先行后列，
        所有索引均为 0-based；返回元组不可通过原地修改影响图片。
        """
        return self._bounds

    @property
    def placement(self) -> str:
        """功能：取得图片的锚定方式。

        使用方法：``image.placement == ImagePlacement.CELL``。
        参数：无。
        返回：``ImagePlacement.FLOATING`` 或 ``ImagePlacement.CELL``。
        """
        return self._placement

    @placement.setter
    def placement(self, value: str) -> None:
        """功能：设置图片是像素浮动锚定还是随单元格区域缩放。

        使用方法：``image.placement = ImagePlacement.CELL``。
        参数：``value`` 必须是 ``ImagePlacement.FLOATING`` 或
        ``ImagePlacement.CELL`` 常量。
        返回：``None``。
        异常：类型或常量无效时抛出 ``TypeError`` 或 ``ValueError``；将已有区域
        锚点改为浮动方式时抛出 ``ValueError``。
        """
        if not isinstance(value, str):
            raise TypeError("placement 必须是 ImagePlacement 常量")
        if value not in {ImagePlacement.FLOATING, ImagePlacement.CELL}:
            raise ValueError("placement 只能是 ImagePlacement.FLOATING 或 CELL")
        if value == ImagePlacement.FLOATING and hasattr(self, "_bounds"):
            if self._bounds[0:2] != self._bounds[2:4]:
                raise ValueError("区域锚点不能设置为浮动图片")
        self._placement = value

    @property
    def fit(self) -> str:
        """功能：取得图片填充锚定区域时采用的比例策略。

        使用方法：``image.fit == ImageFit.CONTAIN``。
        参数：无。
        返回：``ImageFit.STRETCH``、``ImageFit.CONTAIN`` 或 ``ImageFit.COVER``。
        """
        return self._fit

    @fit.setter
    def fit(self, value: str) -> None:
        """功能：设置图片填充区域的缩放策略。

        使用方法：``image.fit = ImageFit.COVER``。
        参数：``value`` 必须是 ``ImageFit.STRETCH``、``CONTAIN`` 或 ``COVER``。
        ``STRETCH`` 拉伸铺满，``CONTAIN`` 保持比例完整显示，``COVER`` 保持
        比例铺满并裁剪超出部分。
        返回：``None``。
        异常：类型或常量无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(value, str):
            raise TypeError("fit 必须是 ImageFit 常量")
        if value not in {ImageFit.STRETCH, ImageFit.CONTAIN, ImageFit.COVER}:
            raise ValueError("fit 只能是 ImageFit.STRETCH、CONTAIN 或 COVER")
        self._fit = value

    @property
    def payload(self) -> bytes:
        """功能：取得写入 XLSX 包的原始图片字节。

        使用方法：由写出器内部读取；业务代码通常不必访问。
        参数：无，只读属性。
        返回：PNG 或 JPEG 二进制内容。
        """
        return self._payload

    @property
    def worksheet(self) -> "Worksheet":
        """功能：取得图片所属工作表。

        使用方法：``worksheet = image.worksheet``。
        参数：无。
        返回：创建该图片的 ``Worksheet``。
        """
        return self._worksheet

    def remove(self) -> "Worksheet":
        """功能：从所属工作表移除当前图片。

        使用方法：``worksheet = image.remove()``。
        参数：无。
        返回：所属 ``Worksheet``，方便继续配置工作表。
        异常：图片已被移除时抛出 ``ValueError``。
        """
        self._worksheet._remove_image(self)
        return self._worksheet


__all__ = ["Image", "ImagePlacement", "ImageFit"]
