"""工作表条件格式规则。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, ClassVar, Optional


class IconSet:
    """Excel 图标集条件格式常量。

    使用方法：``ws.add_icon_set('C2:C100', style=IconSet.THREE_TRAFFIC_LIGHTS)``。
    参数：本类仅提供 Excel 支持的图标集名称常量，不需要实例化。
    返回：无。
    """

    THREE_ARROWS = "3Arrows"
    THREE_TRAFFIC_LIGHTS = "3TrafficLights1"
    THREE_SIGNS = "3Signs"
    FOUR_ARROWS = "4Arrows"
    FIVE_ARROWS = "5Arrows"


_ICON_SETS = {
    IconSet.THREE_ARROWS, IconSet.THREE_TRAFFIC_LIGHTS, IconSet.THREE_SIGNS,
    IconSet.FOUR_ARROWS, IconSet.FIVE_ARROWS,
}


@dataclass
class ConditionalFormat:
    """表示一条应用于区域的条件格式规则。"""

    range: str
    rule: str = "cellIs"
    operator: Optional[str] = None
    formula: Optional[str] = None
    fill: Optional[str] = None
    font: Optional[str] = None
    priority: int = 1
    stop_if_true: bool = False
    options: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        """功能：验证规则字段；参数无；无效值抛出 TypeError 或 ValueError。"""
        if not isinstance(self.priority, int) or self.priority < 1:
            raise ValueError("priority 必须是正整数")
        if not isinstance(self.stop_if_true, bool):
            raise TypeError("stop_if_true 必须是 bool")
        if not isinstance(self.options, dict):
            raise TypeError("options 必须是字典")
        if self.rule == "iconSet" and self.options.get("style") not in _ICON_SETS:
            raise ValueError("图标集 style 必须是 IconSet 的固定值")
        for name in ("fill", "font"):
            value = getattr(self, name)
            if value is not None and (not isinstance(value, str) or len(value) not in (6, 8)):
                raise ValueError(f"{name} 必须是 6 位或 8 位十六进制颜色")
            if value is not None:
                try:
                    int(value, 16)
                except ValueError as error:
                    raise ValueError(f"{name} 必须是十六进制颜色") from error
                setattr(self, name, ("FF" + value if len(value) == 6 else value).upper())


__all__ = ["ConditionalFormat", "IconSet"]
