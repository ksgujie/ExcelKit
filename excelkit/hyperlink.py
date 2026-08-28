"""单元格超链接对象。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True)
class Hyperlink:
    """表示一个可写入 XLSX 的外部或工作簿内部超链接。"""

    target: Optional[str] = None
    display: Optional[str] = None
    tooltip: Optional[str] = None
    location: Optional[str] = None

    def __post_init__(self) -> None:
        """功能：校验超链接目标、显示文本和提示文本。

        使用方法：``Hyperlink(target="https://example.com")``；内部工作簿链接可
        使用 ``Hyperlink(location="Sheet2!A1")``。
        参数：``target`` 为外部网址或文件路径；``display`` 为可选显示文本；
        ``tooltip`` 为可选鼠标提示；``location`` 为可选工作簿内部 A1 位置。
        返回：``None``；对象创建成功后不可变。
        异常：目标和内部位置同时为空、文本参数类型错误或为空字符串时抛出
        ``TypeError`` 或 ``ValueError``。
        """
        for name, value in (
            ("target", self.target),
            ("display", self.display),
            ("tooltip", self.tooltip),
            ("location", self.location),
        ):
            if value is not None and not isinstance(value, str):
                raise TypeError(f"{name} 必须是字符串或 None")
            if isinstance(value, str) and not value:
                raise ValueError(f"{name} 不能为空字符串")
        if self.target is None and self.location is None:
            raise ValueError("target 和 location 至少需要提供一个")
        if self.target is not None and self.location is not None:
            raise ValueError("target 和 location 不能同时提供")


__all__ = ["Hyperlink"]
