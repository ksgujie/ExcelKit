"""基础 XLSX 图片对象。"""

from __future__ import annotations

import os
import struct
from pathlib import Path
from typing import TYPE_CHECKING

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


class Image:
    """表示锚定在工作表内的 PNG 或 JPEG 图片。"""

    __slots__ = ("_worksheet", "filename", "_payload", "anchor", "format", "width", "height", "offset_x", "offset_y", "alt_text")

    def __init__(
        self,
        worksheet: "Worksheet",
        source: os.PathLike[str] | str | bytes,
        anchor: str,
        *,
        name: str | None = None,
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
        self.anchor = anchor
        self.format = image_format
        self.width = width
        self.height = height
        self.offset_x = 0
        self.offset_y = 0
        self.alt_text = ""

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


__all__ = ["Image"]
