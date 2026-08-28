"""工作表数据有效性定义。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class Validation:
    """表示一个应用到连续区域的数据验证规则。"""

    range: str
    kind: str = "list"
    operator: Optional[str] = None
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    values: Optional[tuple[str, ...]] = None
    allow_blank: bool = True
    show_dropdown: bool = True
    prompt_title: Optional[str] = None
    prompt: Optional[str] = None
    error_title: Optional[str] = None
    error: Optional[str] = None
    error_style: str = "stop"

    def __post_init__(self) -> None:
        """功能：验证数据有效性字段；参数无；失败抛出 TypeError 或 ValueError。"""
        kinds = {"list", "whole", "decimal", "date", "time", "textLength", "custom"}
        if self.kind not in kinds:
            raise ValueError(f"不支持的数据验证类型：{self.kind}")
        if self.values is not None:
            self.values = tuple(str(item) for item in self.values)
        if not isinstance(self.allow_blank, bool) or not isinstance(self.show_dropdown, bool):
            raise TypeError("allow_blank 和 show_dropdown 必须是 bool")
        if self.error_style not in {"stop", "warning", "information"}:
            raise ValueError("error_style 必须是 stop、warning 或 information")


__all__ = ["Validation"]
