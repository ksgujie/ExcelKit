"""基础 Excel 数据表（Table）。"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class TotalFunction:
    """汇总函数常量。"""

    NONE = "none"
    SUM = "sum"
    AVERAGE = "average"
    COUNT = "count"
    MAX = "max"
    MIN = "min"


@dataclass
class Table:
    """基础数据表定义。

    参数:
        name: 表名称，整个工作簿内唯一。
        ref: A1 区域地址。
        style: 表样式名称，如 "TableStyleMedium9"。
        has_header: 是否有表头行。
        show_row_stripes: 是否显示行条纹。
        show_column_stripes: 是否显示列条纹。
    """

    name: str
    ref: str
    style: str | None = "TableStyleMedium2"
    has_header: bool = True
    show_row_stripes: bool = True
    show_column_stripes: bool = False
    totals: dict[str, str] = field(default_factory=dict)  # 列名 -> TotalFunction

    def set_total(self, column: str, function: str) -> None:
        """设置指定列的汇总函数。"""
        self.totals[column] = function

    def __repr__(self) -> str:
        return f"Table(name={self.name!r}, ref={self.ref!r})"


__all__ = ["Table", "TotalFunction"]
