"""行列尺寸：行高、列宽、隐藏状态。"""


class RowDimension:
    """单行尺寸。"""

    def __init__(self) -> None:
        self.height: float | None = None
        self.hidden: bool = False


class ColumnDimension:
    """单列尺寸。"""

    def __init__(self) -> None:
        self.width: float | None = None
        self.hidden: bool = False


__all__ = ["RowDimension", "ColumnDimension"]
