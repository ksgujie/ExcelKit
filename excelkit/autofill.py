"""Excel 自动填充模式常量。"""


class AutoFillMode:
    """表示 :meth:`Range.auto_fill` 的可提示填充模式。

    使用方法：``source.auto_fill('A1:A10', mode=AutoFillMode.SERIES)``。
    参数：本类仅提供固定字符串常量，不需要实例化。
    返回：无。
    """

    AUTO = "auto"
    COPY = "copy"
    SERIES = "series"
    FORMATS = "formats"


__all__ = ["AutoFillMode"]
