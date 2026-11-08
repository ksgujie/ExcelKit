"""页面布局与打印设置（简化版），作为 Worksheet.page 唯一入口。"""


class PageMargins:
    """页边距，单位厘米。"""

    def __init__(
        self,
        left: float = 1.91,
        right: float = 1.91,
        top: float = 2.54,
        bottom: float = 2.54,
        header: float = 1.27,
        footer: float = 1.27,
    ) -> None:
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.header = header
        self.footer = footer


class HeaderFooter:
    """页眉/页脚文本及控制符。"""

    def __init__(
        self,
        text: str | None = None,
        left: str | None = None,
        center: str | None = None,
        right: str | None = None,
    ) -> None:
        self.text = text
        self.left = left
        self.center = center
        self.right = right


class PageSettings:
    """页面设置。

    属性:
        orientation: "portrait" 或 "landscape"。
        paper_size: 如 "A4"、"Letter"。
        scale: 百分比缩放，None 表示未启用。
        fit_width / fit_height: 适应页数，设置后会清除 scale。
        print_area: A1 地址区域字符串，如 "A1:F100"。
        repeat_rows: 0-based，(首行, 尾行)，包含结束行。
        repeat_columns: 0-based，(首列, 尾列)。
        margins: PageMargins 实例。
        print_gridlines: 打印时是否显示网格线。
        show_gridlines: 屏幕是否显示网格线。
        horizontal_center / vertical_center: 页面居中。
        header / footer: 页眉页脚文本。
    """

    LANDSCAPE = "landscape"
    PORTRAIT = "portrait"

    def __init__(self) -> None:
        self.orientation: str = self.PORTRAIT
        self.paper_size: str = "A4"
        self.scale: int | None = None
        self.fit_width: int | None = None
        self.fit_height: int | None = None
        self.print_area: str | None = None
        self.repeat_rows: tuple[int, int] | None = None
        self.repeat_columns: tuple[int, int] | None = None
        self.margins: PageMargins = PageMargins()
        self.print_gridlines: bool = False
        self.show_gridlines: bool = True
        self.horizontal_center: bool = False
        self.vertical_center: bool = False
        self._header: HeaderFooter | str = ""
        self._footer: HeaderFooter | str = ""

    def fit(self, width: int | None = None, height: int | None = None) -> None:
        """设置适应页数，会关闭百分比缩放。"""
        self.fit_width = width
        self.fit_height = height
        if width is not None or height is not None:
            self.scale = None

    @property
    def header(self) -> HeaderFooter | str:
        return self._header

    @header.setter
    def header(self, value: HeaderFooter | str | None) -> None:
        if value is None:
            self._header = ""
        elif isinstance(value, HeaderFooter):
            self._header = value
        else:
            self._header = HeaderFooter(text=value or "")

    @property
    def footer(self) -> HeaderFooter | str:
        return self._footer

    @footer.setter
    def footer(self, value: HeaderFooter | str | None) -> None:
        if value is None:
            self._footer = ""
        elif isinstance(value, HeaderFooter):
            self._footer = value
        else:
            self._footer = HeaderFooter(text=value or "")


__all__ = ["PageSettings", "PageMargins"]
