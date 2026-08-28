"""单元格传统批注（备注）对象。"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Note:
    """表示可保存到 XLSX 的传统单元格批注。"""

    text: str
    author: str = "ExcelKit"

    def __post_init__(self) -> None:
        """功能：验证批注正文与作者名称。

        使用方法：``Note('请核对', author='财务部')``；通常也可直接为 ``cell.note``
        赋字符串。
        参数：``text`` 为非空批注正文；``author`` 为非空作者名称。
        返回：无；对象创建后不可变。
        异常：类型不正确时抛出 ``TypeError``，空字符串时抛出 ``ValueError``。
        """
        for name, value in (("text", self.text), ("author", self.author)):
            if not isinstance(value, str):
                raise TypeError(f"{name} 必须是字符串")
            if not value:
                raise ValueError(f"{name} 不能为空字符串")


__all__ = ["Note"]
