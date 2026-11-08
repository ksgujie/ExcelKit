"""工作表图片：支持 PNG/JPEG，锚定到单元格（浮动或随单元格缩放嵌入）。"""

from __future__ import annotations

import io
import os
import struct
from dataclasses import dataclass, field
from typing import Any

from excelkit.address import cell_index, range_index
from excelkit.errors import InvalidAddressError, InvalidFileError


# 像素到 EMU 的换算（Excel/OpenXML 常用）：1 像素 = 9525 EMU（96 DPI 下）
EMU_PER_PIXEL = 9525


def _read_source(source: Any) -> bytes:
    """从路径、bytes、bytearray 或 file-like 对象读取图片二进制数据。"""
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if isinstance(source, str) or hasattr(source, "__fspath__"):
        path = os.fspath(source)
        try:
            with open(path, "rb") as f:
                return f.read()
        except OSError as exc:
            raise InvalidFileError(f"无法读取图片文件：{path}：{exc}") from exc
    if hasattr(source, "read"):
        data = source.read()
        if isinstance(data, str):
            raise InvalidFileError("图片数据必须是二进制，请使用二进制模式打开文件")
        return bytes(data)
    raise InvalidFileError(f"不支持的图片来源类型：{type(source)!r}")


def _detect_image(data: bytes) -> tuple[str, str, int, int]:
    """检测图片格式并返回 (mime, ext, width_px, height_px)。仅支持 PNG/JPEG。"""
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        if len(data) < 24:
            raise InvalidFileError("PNG 数据不完整")
        # IHDR：长度4 + 类型4 + 宽4 + 高4
        width, height = struct.unpack(">II", data[16:24])
        if width <= 0 or height <= 0:
            raise InvalidFileError("PNG 尺寸无效")
        return ("image/png", "png", int(width), int(height))
    if data.startswith(b"\xff\xd8"):
        w, h = _jpeg_size(data)
        return ("image/jpeg", "jpeg", w, h)
    raise InvalidFileError("仅支持 PNG/JPEG 图片格式")


def _jpeg_size(data: bytes) -> tuple[int, int]:
    """解析 JPEG 像素尺寸（扫描 SOF 标记）。"""
    i = 2  # 跳过 SOI
    n = len(data)
    while i + 9 < n:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        i += 2
        # 无长度段的标记
        if marker in (0xD8, 0xD9) or (0xD0 <= marker <= 0xD7):
            continue
        if marker == 0x01:
            continue
        if i + 2 > n:
            break
        seg_len = struct.unpack(">H", data[i : i + 2])[0]
        if seg_len < 2:
            break
        # SOF0-SOF3, SOF5-SOF7, SOF9-SOF11, SOF13-SOF15（排除 DHT/JPG/DAC）
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            if i + 7 > n:
                break
            # seg: len(2) precision(1) height(2) width(2)
            height, width = struct.unpack(">HH", data[i + 3 : i + 7])
            if width <= 0 or height <= 0:
                raise InvalidFileError("JPEG 尺寸无效")
            return (int(width), int(height))
        i += seg_len
    raise InvalidFileError("无法解析 JPEG 尺寸")


@dataclass
class Image:
    """工作表中的图片对象。

    属性说明：
    - anchor: 锚定地址，单格 "B2" 或区域 "B2:D6"
    - mode: "float" 浮动（固定像素尺寸，随单元格移动）；"embed" 嵌入（铺满锚定区域，随行高列宽缩放）
    - width_px/height_px: float 模式下的显示尺寸（像素）；embed 模式下由锚定区域决定
    - offset: float 模式下相对锚点的 (x, y) 像素偏移
    """

    anchor: str
    data: bytes
    mime: str
    ext: str
    natural_width: int
    natural_height: int
    mode: str = "float"
    width_px: int | None = None
    height_px: int | None = None
    offset: tuple[int, int] = (0, 0)
    alt_text: str | None = None
    # 内部：写出时分配
    _part_name: str | None = field(default=None, repr=False, compare=False)

    @property
    def display_width(self) -> int:
        """当前显示宽度（像素）。embed 模式下调用方应按锚定区域计算。"""
        if self.width_px is not None:
            return self.width_px
        return self.natural_width

    @property
    def display_height(self) -> int:
        """当前显示高度（像素）。embed 模式下调用方应按锚定区域计算。"""
        if self.height_px is not None:
            return self.height_px
        return self.natural_height


def create_image(
    anchor: str,
    source: Any,
    *,
    mode: str = "float",
    width: int | None = None,
    height: int | None = None,
    keep_aspect: bool = True,
    offset: tuple[int, int] = (0, 0),
    alt_text: str | None = None,
) -> Image:
    """创建 Image 对象并校验参数。

    参数:
        anchor: 锚定单元格或区域，如 "B2"、"B2:D6"
        source: 图片路径、bytes 或二进制 file-like
        mode: "float"（默认，固定尺寸）或 "embed"（铺满锚定区域，随长宽变化）
        width/height: float 模式下的显示尺寸（像素）；embed 模式下传入将报错
        keep_aspect: float 模式下仅指定宽或高时是否保持比例
        offset: float 模式下的 (x, y) 像素偏移
        alt_text: 替代文字
    """
    # 校验锚点
    try:
        if ":" in str(anchor):
            range_index(anchor)
        else:
            cell_index(anchor)
    except InvalidAddressError:
        raise
    if mode not in ("float", "embed"):
        raise ValueError(f"mode 必须是 'float' 或 'embed'：{mode!r}")
    if mode == "embed" and (width is not None or height is not None):
        raise ValueError("embed 模式下图片随锚定单元格/区域缩放，不支持 width/height 参数")
    data = _read_source(source)
    mime, ext, nat_w, nat_h = _detect_image(data)

    w_px: int | None = None
    h_px: int | None = None
    if mode == "float":
        if width is not None and height is not None:
            w_px, h_px = int(width), int(height)
            if keep_aspect and nat_w > 0 and nat_h > 0:
                # 同时指定时，keep_aspect=True 仍按用户给的两者？通常显式给两者意味着用户要这个尺寸
                # 但为避免歧义：两者都给则尊重两者
                pass
        elif width is not None:
            w_px = int(width)
            h_px = max(1, round(nat_h * w_px / nat_w)) if keep_aspect and nat_w > 0 else nat_h
        elif height is not None:
            h_px = int(height)
            w_px = max(1, round(nat_w * h_px / nat_h)) if keep_aspect and nat_h > 0 else nat_w
        else:
            w_px, h_px = nat_w, nat_h
        if w_px is not None and w_px <= 0:
            raise ValueError("图片宽度必须为正整数")
        if h_px is not None and h_px <= 0:
            raise ValueError("图片高度必须为正整数")

    off = (int(offset[0]), int(offset[1])) if offset else (0, 0)
    return Image(
        anchor=str(anchor).strip().replace("$", "").upper() if ":" not in str(anchor) else _normalize_range_anchor(anchor),
        data=data,
        mime=mime,
        ext=ext,
        natural_width=nat_w,
        natural_height=nat_h,
        mode=mode,
        width_px=w_px,
        height_px=h_px,
        offset=off,
        alt_text=alt_text,
    )


def _normalize_range_anchor(anchor: str) -> str:
    """规范化区域锚点为大写 A1:B2 形式。"""
    r1, c1, r2, c2 = range_index(anchor)
    from excelkit.address import range_address

    return range_address(r1, c1, r2, c2)


__all__ = ["Image", "create_image", "EMU_PER_PIXEL"]
