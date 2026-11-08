"""工作簿级命名区域。"""


class NamedRange:
    """命名区域，名称大小写不敏感，关联一个 Range。"""

    def __init__(self, name: str, ref: str) -> None:
        self.name = name
        self.ref = ref  # A1 地址字符串
        self._sheet = None  # 关联的工作表对象，由 Workbook.add_named_range 设置

    def __repr__(self) -> str:
        return f"NamedRange(name={self.name!r}, ref={self.ref!r})"


__all__ = ["NamedRange"]
