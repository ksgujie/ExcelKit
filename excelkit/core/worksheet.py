"""工作表对象及单元格、区域访问接口。"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
import os
import re
import math
import unicodedata
from typing import TYPE_CHECKING, Any, Dict, Optional, Tuple, Union

from ..address import (
    cell_address,
    cell_index,
    column_to_index,
    index_to_column,
    range_index,
    range_address,
    validate_column_index,
    validate_row_index,
)
from ..hyperlink import Hyperlink
from ..storage import ValueStore
from ..style import DEFAULT_STYLE, Style, _color
from .cell import Cell
from .conversion import normalize_value
from .dimension import ColumnDimension, RowDimension
from .page import PageSettings
from .range import Range
from .table import Table
from ..validation import Validation
from ..conditional import ConditionalFormat, IconSet
from ..protection import Protection
from ..filter import AutoFilter
from ..note import Note
from ..sort import SortKey
from ..chart import Chart, ChartType
from ..image import Image, ImageFit, ImagePlacement

if TYPE_CHECKING:
    from .workbook import Workbook


def _normalize_value(value: Any) -> Any:
    """功能：集中规范化写入工作表的普通值。

    使用方法：由 :meth:`Worksheet._set_value` 内部调用，所有单元格、区域和追加
    写入都经过此函数。
    参数：``value`` 为任意 Python 对象；严格匹配 ``#YYYY-M-D``、
    ``#YYYY/M/D`` 或在其后追加 ``H:M``、``H:M:S`` 的字符串会转换。
    返回：无时间部分时返回 :class:`datetime.date`，有时间部分时返回
    :class:`datetime.datetime`；未匹配字面量时返回原值。
    异常：格式匹配但日期或时间取值无效时抛出 ``ValueError``。
    """
    return normalize_value(value)


class Worksheet:
    """表示隶属于某个 :class:`Workbook` 的命名工作表。"""

    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "very_hidden"

    __slots__ = (
        "_workbook",
        "_name",
        "_color",
        "_visibility",
        "_values",
        "_formulas",
        "_formula_values",
        "_formula_errors",
        "_hyperlinks",
        "_notes",
        "_headers",
        "_styles",
        "_merged_ranges",
        "_rows",
        "_columns",
        "_freeze",
        "_filter_range",
        "_filter_conditions",
        "_show_gridlines",
        "_page",
        "_horizontal_page_breaks",
        "_tables",
        "_validations",
        "_conditionals",
        "_charts",
        "_images",
        "_protection",
        "_max_row",
        "_max_column",
    )

    def __init__(self, workbook: "Workbook", name: str) -> None:
        """功能：初始化工作表、普通值存储、公式存储和最大索引。

        使用方法：通常由 ``workbook.add_sheet(name)`` 创建，不直接调用。
        参数：``workbook`` 为所属工作簿；``name`` 为已经验证的工作表名称。
        返回：无；空表的最大行、列索引均初始化为 ``-1``。
        """
        self._workbook = workbook
        self._name = name
        self._color: Optional[str] = None
        self._visibility = self.VISIBLE
        self._values = ValueStore()
        self._formulas: Dict[Tuple[int, int], str] = {}
        self._formula_values: Dict[Tuple[int, int], Any] = {}
        self._formula_errors: Dict[Tuple[int, int], str] = {}
        self._hyperlinks: Dict[Tuple[int, int], Hyperlink] = {}
        self._notes: Dict[Tuple[int, int], Note] = {}
        self._headers: Optional[Tuple[Any, ...]] = None
        self._styles: Dict[Tuple[int, int], Style] = {}
        self._merged_ranges: list[Tuple[int, int, int, int]] = []
        self._rows: Dict[int, RowDimension] = {}
        self._columns: Dict[int, ColumnDimension] = {}
        self._freeze: Optional[str] = None
        self._filter_range: Optional[str] = None
        self._filter_conditions: dict[int, tuple[str, ...]] = {}
        self._show_gridlines = True
        self._page = PageSettings()
        self._horizontal_page_breaks: set[int] = set()
        self._tables: Dict[str, Table] = {}
        self._validations: list[Validation] = []
        self._conditionals: list[ConditionalFormat] = []
        self._charts: list[Chart] = []
        self._images: list[Image] = []
        self._protection = Protection()
        self._max_row = -1
        self._max_column = -1

    @property
    def name(self) -> str:
        """功能：取得工作表名称。

        使用方法：``name = worksheet.name``。
        参数：无。
        返回：工作表底部标签显示的名称字符串。
        """
        return self._name

    @name.setter
    def name(self, name: str) -> None:
        """功能：重命名工作表并同步所属工作簿的名称索引。

        使用方法：``worksheet.name = '新名称'``。
        参数：``name`` 为符合 Excel 规则的新名称字符串，长度为 1～31。
        返回：``None``；工作表对象、顺序、数据和样式均保持不变。
        异常：名称无效时抛出 ``InvalidWorksheetNameError``；与其他工作表名称
        大小写不敏感重复时抛出 ``ValueError``。
        """
        self._workbook._rename_sheet(self, name)

    @property
    def color(self) -> Optional[str]:
        """功能：取得工作表标签颜色。

        使用方法：``color = worksheet.color``。
        参数：无，只读时不需要参数；设置颜色使用同名属性设置器。
        返回：8 位大写 ARGB 字符串；没有设置颜色时返回 ``None``。
        """
        return self._color

    @color.setter
    def color(self, color: Optional[str]) -> None:
        """功能：设置或清除工作表标签颜色。

        使用方法：``worksheet.color = '4472C4'``；赋值 ``None`` 清除颜色。
        参数：``color`` 为 6 位 ``RRGGBB``、8 位 ``AARRGGBB`` 字符串或
        ``None``；6 位颜色自动补为完全不透明 ARGB。
        返回：``None``。
        异常：颜色类型、长度或十六进制字符无效时抛出 ``ValueError``。
        """
        self._color = _color(color)

    @property
    def visibility(self) -> str:
        """功能：取得工作表可见状态。

        使用方法：``state = worksheet.visibility``。
        参数：无。
        返回：``Worksheet.VISIBLE``、``Worksheet.HIDDEN`` 或
        ``Worksheet.VERY_HIDDEN``。
        """
        return self._visibility

    @visibility.setter
    def visibility(self, value: str) -> None:
        """功能：设置工作表可见状态。

        使用方法：``worksheet.visibility = Worksheet.HIDDEN``。
        参数：``value`` 必须为 ``VISIBLE``、``HIDDEN``、``VERY_HIDDEN`` 三个常量之一。
        返回：``None``。
        异常：类型或值无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(value, str):
            raise TypeError("visibility 必须是字符串常量")
        if value not in {self.VISIBLE, self.HIDDEN, self.VERY_HIDDEN}:
            raise ValueError("visibility 必须是 Worksheet.VISIBLE、HIDDEN 或 VERY_HIDDEN")
        self._visibility = value

    @property
    def headers(self) -> Optional[Tuple[Any, ...]]:
        """功能：取得分隔文本读取时识别出的首行表头。

        使用方法：``headers = worksheet.headers``。
        参数：无。
        返回：启用 ``Workbook.load(..., has_header=True)`` 且来自 CSV/TSV 时，
        返回首行字段的只读元组；普通创建或 XLS/XLSX 读取的工作表返回 ``None``。
        返回值只是元数据，首行仍会保留在单元格中。
        """
        return self._headers

    def cell(self, row: Union[str, int], column: Optional[int] = None) -> Cell:
        """功能：按 A1 地址或 0-based 行列索引取得单元格。

        使用方法：固定地址使用 ``worksheet.cell("B3")``；动态坐标使用
        ``worksheet.cell(2, 1)``，两者都指向 ``B3``。
        参数：``row`` 为字符串时表示 A1 地址且必须省略 ``column``；``row`` 为
        整数时表示 0-based 行索引，``column`` 必须是 0-based 列索引。顺序始终
        先行后列，布尔值不作为整数索引。
        返回：指向指定位置的 :class:`Cell`。
        异常：参数组合错误时抛出 ``TypeError``；地址或索引无效时抛出
        ``InvalidAddressError``。
        """
        if isinstance(row, str):
            if column is not None:
                raise TypeError("使用 A1 地址访问单元格时不能再传入 column")
            parsed_row, parsed_column = cell_index(row)
            return Cell(self, parsed_row, parsed_column)
        if isinstance(row, int) and not isinstance(row, bool) and column is not None:
            validate_row_index(row)
            validate_column_index(column)
            return Cell(self, row, column)
        raise TypeError("cell() 需要一个 A1 地址，或 row、column 两个 0-based 整数索引")

    def range(self, address: str) -> Range:
        """功能：按 A1 区域地址取得连续矩形区域。

        使用方法：``area = worksheet.range("A1:C10")``。
        参数：``address`` 为包含冒号的标准矩形 A1 区域字符串。
        返回：对应的 :class:`Range`。
        异常：地址、边界或方向无效时抛出 ``InvalidAddressError``。
        """
        return Range(self, *range_index(address))

    def row(self, index: int) -> RowDimension:
        """功能：按0-based索引取得可设置行高和隐藏状态的行对象。

        使用方法：``worksheet.row(0).height = 28``。
        参数：``index`` 为0～1048575的整数，布尔值不作为索引。
        返回：当前行唯一的 :class:`RowDimension` 对象，重复读取返回同一实例。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(index)
        if index not in self._rows:
            self._rows[index] = RowDimension(index)
        return self._rows[index]

    def column(self, index: int) -> ColumnDimension:
        """功能：按0-based索引取得可设置列宽和隐藏状态的列对象。

        使用方法：``worksheet.column(0).width = 20``。
        参数：``index`` 为0～16383的整数，布尔值不作为索引。
        返回：当前列唯一的 :class:`ColumnDimension` 对象，重复读取返回同一实例。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_column_index(index)
        if index not in self._columns:
            self._columns[index] = ColumnDimension(index)
        return self._columns[index]

    @staticmethod
    def _group_bounds(first: int, last: int, *, rows: bool) -> tuple[int, int]:
        """功能：验证行列分组的包含式 0-based 边界。

        使用方法：由四个公开分组方法内部调用。
        参数：``first``、``last`` 为包含式边界；``rows`` 指示验证行还是列。
        返回：验证后的 ``(first, last)`` 元组。
        异常：索引类型、范围或顺序无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        validator = validate_row_index if rows else validate_column_index
        validator(first)
        validator(last)
        if first > last:
            raise ValueError("first 不能大于 last")
        return first, last

    def _group_axis(
        self, first: int, last: int, *, rows: bool, collapsed: bool
    ) -> "Worksheet":
        """功能：为连续行或列增加一层 Excel 大纲分组。

        使用方法：由 ``group_rows()`` 和 ``group_columns()`` 复用。
        参数：``first``、``last`` 为包含式边界；``rows`` 选择轴；``collapsed``
        控制是否隐藏成员并记录折叠标志。
        返回：当前工作表。
        异常：参数或嵌套层级无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(collapsed, bool):
            raise TypeError("collapsed 必须是 bool")
        first, last = self._group_bounds(first, last, rows=rows)
        getter = self.row if rows else self.column
        dimensions = [getter(index) for index in range(first, last + 1)]
        if any(dimension.outline_level >= 7 for dimension in dimensions):
            raise ValueError("Excel 行列分组最多支持 7 层嵌套")
        for dimension in dimensions:
            dimension.outline_level += 1
            if collapsed:
                dimension.hidden = True
        if collapsed:
            dimensions[-1].collapsed = True
        return self

    def _ungroup_axis(self, first: int, last: int, *, rows: bool) -> "Worksheet":
        """功能：从连续行或列移除一层 Excel 大纲分组。

        使用方法：由 ``ungroup_rows()`` 和 ``ungroup_columns()`` 复用。
        参数：``first``、``last`` 为包含式边界；``rows`` 选择轴。
        返回：当前工作表。
        异常：边界无效或范围内存在未分组维度时抛出 ``TypeError`` 或 ``ValueError``。
        """
        first, last = self._group_bounds(first, last, rows=rows)
        getter = self.row if rows else self.column
        dimensions = [getter(index) for index in range(first, last + 1)]
        if any(dimension.outline_level == 0 for dimension in dimensions):
            raise ValueError("目标范围中存在未分组的行或列")
        was_collapsed = dimensions[-1].collapsed
        for dimension in dimensions:
            dimension.outline_level -= 1
            if was_collapsed and dimension.outline_level == 0:
                dimension.hidden = False
        dimensions[-1].collapsed = was_collapsed and any(
            dimension.outline_level > 0 for dimension in dimensions
        )
        return self

    def group_rows(
        self, first_row: int, last_row: int, *, collapsed: bool = False
    ) -> "Worksheet":
        """功能：将连续行组成可在 Excel 中折叠的大纲组。

        使用方法：``ws.group_rows(1, 10)``；``collapsed=True`` 会以折叠状态打开。
        参数：``first_row``、``last_row`` 为包含式 0-based 行索引；``collapsed``
        控制保存文件打开时是否隐藏分组明细。
        返回：当前工作表。
        异常：索引、顺序、层级或开关无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return self._group_axis(first_row, last_row, rows=True, collapsed=collapsed)

    def ungroup_rows(self, first_row: int, last_row: int) -> "Worksheet":
        """功能：移除连续行的一层大纲分组。

        使用方法：``ws.ungroup_rows(1, 10)``。
        参数：``first_row``、``last_row`` 为包含式 0-based 行索引。
        返回：当前工作表。
        异常：范围中存在未分组行或边界无效时抛出 ``ValueError``。
        """
        return self._ungroup_axis(first_row, last_row, rows=True)

    def group_columns(
        self, first_column: int, last_column: int, *, collapsed: bool = False
    ) -> "Worksheet":
        """功能：将连续列组成可在 Excel 中折叠的大纲组。

        使用方法：``ws.group_columns(1, 4, collapsed=True)``。
        参数：``first_column``、``last_column`` 为包含式 0-based 列索引；
        ``collapsed`` 控制是否以折叠状态保存。
        返回：当前工作表。
        异常：索引、顺序、层级或开关无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return self._group_axis(first_column, last_column, rows=False, collapsed=collapsed)

    def ungroup_columns(self, first_column: int, last_column: int) -> "Worksheet":
        """功能：移除连续列的一层大纲分组。

        使用方法：``ws.ungroup_columns(1, 4)``。
        参数：``first_column``、``last_column`` 为包含式 0-based 列索引。
        返回：当前工作表。
        异常：范围中存在未分组列或边界无效时抛出 ``ValueError``。
        """
        return self._ungroup_axis(first_column, last_column, rows=False)

    @property
    def merged_ranges(self) -> tuple[Range, ...]:
        """功能：取得全部合并区域的只读顺序快照。

        使用方法：``for area in worksheet.merged_ranges: print(area.address)``。
        参数：无；修改合并状态使用 ``Range.merge()`` 或 ``Range.unmerge()``。
        返回：按左上角位置排序的 ``tuple[Range, ...]``。
        """
        return tuple(Range(self, *bounds) for bounds in self._merged_ranges)

    @property
    def freeze_panes(self) -> Optional[str]:
        """功能：读取冻结窗格后的第一个可滚动单元格地址。

        使用方法：``address = worksheet.freeze_panes``。
        参数：无。
        返回：大写A1地址或没有冻结窗格时的 ``None``。
        """
        return self._freeze

    @freeze_panes.setter
    def freeze_panes(self, address: Optional[str]) -> None:
        """功能：设置或清除冻结行列。

        使用方法：``worksheet.freeze_panes = "B2"`` 冻结第一行和第一列；赋值
        ``None`` 或 ``"A1"`` 清除冻结。
        参数：``address`` 为第一个可滚动单元格的A1地址或 ``None``。
        返回：``None``。
        异常：地址类型、格式或边界无效时抛出 ``InvalidAddressError``。
        """
        if address is None:
            self._freeze = None
            return
        row, column = cell_index(address)
        self._freeze = None if (row, column) == (0, 0) else cell_address(row, column)

    @property
    def auto_filter(self) -> AutoFilter:
        """功能：取得自动筛选代理对象。

        使用方法：``ws.auto_filter.range = 'A1:D20'``；
        ``ws.auto_filter.set(1, ['通过']).apply()``。
        参数：无。
        返回：绑定当前工作表的 :class:`AutoFilter` 代理。
        """
        return AutoFilter(self)

    @property
    def show_gridlines(self) -> bool:
        """功能：读取Excel屏幕是否显示工作表网格线。

        使用方法：``visible = worksheet.show_gridlines``。
        参数：无。
        返回：布尔值；此属性不控制打印网格线。
        """
        return self._show_gridlines

    @show_gridlines.setter
    def show_gridlines(self, value: bool) -> None:
        """功能：设置Excel屏幕中的工作表网格线可见性。

        使用方法：``worksheet.show_gridlines = False``。
        参数：``value`` 必须是布尔值。
        返回：``None``。
        异常：类型无效时抛出 ``TypeError``。
        """
        if not isinstance(value, bool):
            raise TypeError("show_gridlines 必须是布尔值")
        self._show_gridlines = value

    @property
    def page(self) -> PageSettings:
        """功能：取得当前工作表唯一的页面和打印设置对象。

        使用方法：``worksheet.page.orientation = "landscape"``。
        参数：无；属性本身只读，不允许整体替换。
        返回：当前 :class:`PageSettings`。
        """
        return self._page

    @property
    def protection(self) -> Protection:
        """功能：取得工作表保护设置对象。

        使用方法：``worksheet.protection.enabled = True``。
        参数：无，只读属性；返回对象的字段可以直接修改。
        返回：同一个可修改 :class:`Protection`；保存 XLSX 时写入工作表保护定义。
        """
        return self._protection

    def add_validation(self, address: str, *, kind: str = "list", values=None,
                       operator: Optional[str] = None, formula1: Optional[str] = None,
                       formula2: Optional[str] = None, allow_blank: bool = True,
                       show_dropdown: bool = True, prompt_title: Optional[str] = None,
                       prompt: Optional[str] = None, error_title: Optional[str] = None,
                       error: Optional[str] = None, error_style: str = "stop") -> Validation:
        """功能：为区域添加 Excel 数据有效性规则。

        使用方法：``ws.add_validation('B2:B100', kind='list', values=['是','否'])``。
        参数：``address`` 为A1区域；``kind`` 可为 list、whole、decimal、date、time、
        textLength、custom；``values`` 为下拉候选；``formula1/formula2`` 为公式；其余
        参数控制运算符、空值、提示和错误信息。返回新建 ``Validation``。
        """
        area = self.range(address)
        validation = Validation(area.address, kind=kind, operator=operator,
            formula1=formula1, formula2=formula2, values=values, allow_blank=allow_blank,
            show_dropdown=show_dropdown, prompt_title=prompt_title, prompt=prompt,
            error_title=error_title, error=error, error_style=error_style)
        self._validations.append(validation)
        return validation

    @property
    def validations(self) -> tuple[Validation, ...]:
        """功能：取得当前工作表全部数据有效性规则的只读快照。

        使用方法：``for rule in worksheet.validations: ...``。
        参数：无，只读属性。
        返回：按添加顺序排列的 ``tuple[Validation, ...]``。
        """
        return tuple(self._validations)

    def remove_validation(self, validation: Validation) -> "Worksheet":
        """功能：删除当前工作表中的数据有效性规则。

        使用方法：``worksheet.remove_validation(validation)``。
        参数：``validation`` 必须是当前工作表已经登记的 :class:`Validation`。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：规则不属于当前工作表时抛出 ``ValueError``。
        """
        if validation not in self._validations:
            raise ValueError("数据有效性规则不属于当前工作表")
        self._validations.remove(validation)
        return self

    def add_conditional_format(self, address: str, *, rule: str = "cellIs",
                               operator: Optional[str] = None, formula: Optional[str] = None,
                               fill: Optional[str] = None, font: Optional[str] = None,
                               priority: Optional[int] = None, stop_if_true: bool = False) -> ConditionalFormat:
        """功能：为区域添加条件格式规则。

        使用方法：``ws.add_conditional_format('B2:B20', operator='greaterThan', formula='90', fill='FFC7CE')``。
        参数：``address`` 为A1区域；``rule`` 为 cellIs、expression 等规则类型；其余
        参数指定比较、公式、填充/字体颜色及优先级。返回新建 ``ConditionalFormat``。
        """
        area = self.range(address)
        item = ConditionalFormat(area.address, rule=rule, operator=operator, formula=formula,
                                 fill=fill, font=font,
                                 priority=priority or len(self._conditionals) + 1,
                                 stop_if_true=stop_if_true)
        self._conditionals.append(item)
        return item

    def add_color_scale(
        self,
        address: str,
        *,
        min_color: str = "F8696B",
        mid_color: str | None = "FFEB84",
        max_color: str = "63BE7B",
    ) -> ConditionalFormat:
        """功能：为区域添加按数值渐变着色的双色或三色条件格式。

        使用方法：``ws.add_color_scale('C2:C100')``。
        参数：``address`` 为 A1 区域；``min_color``、``max_color`` 为最低和最高值
        颜色；``mid_color`` 为中间颜色，设为 ``None`` 时创建双色渐变。
        返回：新建 :class:`ConditionalFormat`。
        异常：区域或颜色无效时抛出 ``ValueError``。
        """
        colors = [min_color, max_color] if mid_color is None else [min_color, mid_color, max_color]
        normalized = [_color(color) for color in colors]
        item = ConditionalFormat(
            self.range(address).address, rule="colorScale",
            priority=len(self._conditionals) + 1, options={"colors": normalized},
        )
        self._conditionals.append(item)
        return item

    def add_data_bar(
        self, address: str, *, color: str = "638EC6", show_value: bool = True
    ) -> ConditionalFormat:
        """功能：为数值区域添加按最小值和最大值自动缩放的数据条。

        使用方法：``ws.add_data_bar('D2:D100', color='5B9BD5')``。
        参数：``address`` 为 A1 区域；``color`` 为数据条 RGB/ARGB 色值；
        ``show_value`` 控制是否保留单元格数值显示。
        返回：新建 :class:`ConditionalFormat`。
        异常：颜色或开关类型无效时抛出 ``ValueError`` 或 ``TypeError``。
        """
        if not isinstance(show_value, bool):
            raise TypeError("show_value 必须是 bool")
        item = ConditionalFormat(
            self.range(address).address, rule="dataBar",
            priority=len(self._conditionals) + 1,
            options={"color": _color(color), "show_value": show_value},
        )
        self._conditionals.append(item)
        return item

    def add_icon_set(
        self, address: str, *, style: str = IconSet.THREE_TRAFFIC_LIGHTS
    ) -> ConditionalFormat:
        """功能：为区域添加按百分位自动分档的 Excel 图标集条件格式。

        使用方法：``ws.add_icon_set('E2:E100', style=IconSet.THREE_TRAFFIC_LIGHTS)``。
        参数：``address`` 为 A1 区域；``style`` 为 ``IconSet`` 固定值。
        返回：新建 :class:`ConditionalFormat`。
        异常：区域或图标集类型无效时抛出 ``ValueError``。
        """
        item = ConditionalFormat(
            self.range(address).address, rule="iconSet",
            priority=len(self._conditionals) + 1, options={"style": style},
        )
        self._conditionals.append(item)
        return item

    @property
    def conditional_formats(self) -> tuple[ConditionalFormat, ...]:
        """功能：取得工作表全部条件格式规则的只读快照。

        使用方法：``for rule in worksheet.conditional_formats: ...``。
        参数：无，只读属性。
        返回：按添加顺序排列的 ``tuple[ConditionalFormat, ...]``。
        """
        return tuple(self._conditionals)

    def remove_conditional_format(self, item: ConditionalFormat) -> "Worksheet":
        """功能：删除当前工作表中的一条条件格式规则。

        使用方法：``worksheet.remove_conditional_format(rule)``。
        参数：``item`` 必须是当前工作表已经登记的 :class:`ConditionalFormat`。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：规则不属于当前工作表时抛出 ``ValueError``。
        """
        if item not in self._conditionals:
            raise ValueError("条件格式规则不属于当前工作表")
        self._conditionals.remove(item)
        return self

    @staticmethod
    def _sort_value(value: Any) -> tuple[int, Any]:
        """功能：把不同 Python 值规范为稳定可比较的排序键。

        使用方法：由 :meth:`sort` 内部调用。
        参数：``value`` 为普通值或公式缓存值。
        返回：包含类型优先级及规范化值的二元组，空值会排在其他值之前。
        """
        if value is None:
            return 0, ""
        if isinstance(value, bool):
            return 1, int(value)
        if isinstance(value, (int, float)):
            return 2, value
        return 3, str(value).casefold()

    def sort(
        self,
        address: str,
        *,
        keys: Iterable[SortKey],
        has_header: bool = False,
    ) -> "Worksheet":
        """功能：按一个或多个区域内 0-based 相对列键排序连续矩形区域。

        使用方法：``ws.sort('A2:D100', keys=[SortKey(1, descending=True)])``。
        参数：``address`` 为排序 A1 区域；``keys`` 为非空 ``SortKey`` 可迭代对象，
        每个 ``column`` 相对于区域左侧且从0开始；``has_header`` 为真时首行不排序。
        返回：当前工作表，支持链式调用。
        异常：键、区域或合并区域不适合排序时抛出 ``TypeError`` 或 ``ValueError``。
        """
        area = self.range(address)
        if not isinstance(has_header, bool):
            raise TypeError("has_header 必须是 bool")
        prepared_keys = tuple(keys)
        if not prepared_keys or any(not isinstance(key, SortKey) for key in prepared_keys):
            raise TypeError("keys 必须是非空 SortKey 可迭代对象")
        width = area.max_column - area.min_column + 1
        if any(key.column >= width for key in prepared_keys):
            raise ValueError("排序键列索引超出区域")
        if any(
            not (merged.max_row < area.min_row or merged.min_row > area.max_row
                 or merged.max_column < area.min_column or merged.min_column > area.max_column)
            for merged in self.merged_ranges
        ):
            raise ValueError("包含合并单元格的区域不能排序")
        first_row = area.min_row + int(has_header)
        source_rows = list(range(first_row, area.max_row + 1))
        for key in reversed(prepared_keys):
            column = area.min_column + key.column
            source_rows.sort(
                key=lambda row: self._sort_value(
                    self._formula_values.get((row, column), self._values.get(row, column))
                ),
                reverse=key.descending,
            )
        mappings = (
            self._values._values, self._formulas, self._formula_values,
            self._formula_errors, self._hyperlinks, self._notes, self._styles,
        )
        for mapping in mappings:
            source_values = {
                (row, column): mapping[(row, column)]
                for row in range(first_row, area.max_row + 1)
                for column in range(area.min_column, area.max_column + 1)
                if (row, column) in mapping
            }
            for row in range(first_row, area.max_row + 1):
                for column in range(area.min_column, area.max_column + 1):
                    mapping.pop((row, column), None)
            for target_row, source_row in zip(range(first_row, area.max_row + 1), source_rows):
                for column in range(area.min_column, area.max_column + 1):
                    value = source_values.get((source_row, column))
                    if (source_row, column) in source_values:
                        mapping[(target_row, column)] = value
        self._workbook._invalidate_formula_caches()
        return self

    @staticmethod
    def _matches_search(
        candidate: Any,
        query: Any,
        *,
        match_case: bool,
        whole: bool,
    ) -> bool:
        """功能：按查找选项判断一个候选值是否匹配查询值。

        使用方法：由 :meth:`find` 与 :meth:`replace` 内部调用。
        参数：``candidate`` 为单元格值或公式；``query`` 为待查找值；``match_case``
        控制字符串大小写；``whole`` 控制字符串是否必须完整相等。
        返回：匹配时为 ``True``；非字符串查询始终采用 Python 相等比较。
        """
        if isinstance(query, str):
            if not isinstance(candidate, str):
                candidate = str(candidate)
            if not match_case:
                candidate, query = candidate.casefold(), query.casefold()
            return candidate == query if whole else query in candidate
        return candidate == query

    def find(
        self,
        query: Any,
        *,
        match_case: bool = False,
        whole: bool = False,
        in_formulas: bool = False,
    ) -> tuple[Cell, ...]:
        """功能：在已使用区域的普通值或公式中查找全部匹配单元格。

        使用方法：``ws.find('张三')`` 进行不区分大小写的包含查找；
        ``ws.find('=SUM', in_formulas=True)`` 在公式文本中查找。
        参数：``query`` 为非 ``None`` 的查找值；字符串可配合 ``match_case`` 和
        ``whole`` 控制匹配方式；``in_formulas`` 为真时只搜索公式文本。
        返回：按行优先顺序排列的 ``Cell`` 元组；未找到时返回空元组。
        异常：空字符串、空值或开关类型无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if query is None:
            raise ValueError("query 不能为 None")
        if isinstance(query, str) and not query:
            raise ValueError("字符串 query 不能为空")
        for name, value in (
            ("match_case", match_case), ("whole", whole), ("in_formulas", in_formulas)
        ):
            if not isinstance(value, bool):
                raise TypeError(f"{name} 必须是 bool")
        matches: list[Cell] = []
        for row in range(self._max_row + 1):
            for column in range(self._max_column + 1):
                coordinate = (row, column)
                candidate = (
                    self._formulas.get(coordinate)
                    if in_formulas else self._values.get(row, column)
                )
                if candidate is not None and self._matches_search(
                    candidate, query, match_case=match_case, whole=whole
                ):
                    matches.append(self.cell(row, column))
        return tuple(matches)

    def replace(
        self,
        query: Any,
        replacement: Any,
        *,
        match_case: bool = False,
        whole: bool = False,
        in_formulas: bool = False,
        limit: int | None = None,
    ) -> int:
        """功能：替换已使用区域内普通值或公式文本的全部匹配内容。

        使用方法：``count = ws.replace('旧名称', '新名称')``；使用
        ``ws.replace('SUM', 'AVERAGE', in_formulas=True)`` 可替换公式片段。
        参数：``query``、匹配开关的规则与 :meth:`find` 相同；``replacement`` 为
        替换值，字符串部分匹配时必须为字符串；``in_formulas`` 为真时替换值必须为
        字符串；``limit`` 为可选的最大替换次数，``0`` 表示不替换。
        返回：实际发生替换的单元格数量。
        异常：参数不合法、公式替换后无效或写入合并区域限制时抛出相应异常；开始
        替换前会先找出全部目标，单个写入失败时之前替换不回滚。
        """
        if limit is not None and (
            isinstance(limit, bool) or not isinstance(limit, int) or limit < 0
        ):
            raise ValueError("limit 必须是非负整数或 None")
        matches = self.find(
            query, match_case=match_case, whole=whole, in_formulas=in_formulas
        )
        if in_formulas and not isinstance(replacement, str):
            raise TypeError("替换公式时 replacement 必须是字符串")
        if isinstance(query, str) and not whole and not isinstance(replacement, str):
            raise TypeError("替换字符串片段时 replacement 必须是字符串")
        replaced = 0
        for cell in matches:
            if limit is not None and replaced >= limit:
                break
            if in_formulas:
                source = cell.formula
                if source is None:
                    continue
                flags = 0 if match_case else re.IGNORECASE
                target = (
                    replacement if whole
                    else re.sub(
                        re.escape(query), lambda _match: replacement,
                        source, flags=flags,
                    )
                )
                cell.formula = target
            else:
                source = cell.value
                if isinstance(query, str) and isinstance(source, str):
                    flags = 0 if match_case else re.IGNORECASE
                    target = (
                        replacement if whole
                        else re.sub(
                            re.escape(query), lambda _match: replacement,
                            source, flags=flags,
                        )
                    )
                else:
                    target = replacement
                cell.value = target
            replaced += 1
        return replaced

    def export(
        self,
        filename: str | os.PathLike[str],
        *,
        encoding: str = "utf-8-sig",
        delimiter: str | None = None,
        formulas: bool = False,
    ) -> "Worksheet":
        """功能：将当前单张工作表导出为 CSV 或 TSV 文件。

        使用方法：``ws.export('sales.csv')``；``ws.export('raw.tsv', formulas=True)``
        会输出公式原文而不是缓存结果。
        参数：``filename`` 必须以 ``.csv`` 或 ``.tsv`` 结尾；``encoding`` 为文件
        编码，默认带 BOM 的 UTF-8；``delimiter`` 可覆盖扩展名默认分隔符；
        ``formulas`` 控制公式单元格输出内容。
        返回：当前 ``Worksheet``，支持链式调用。
        异常：参数、扩展名、编码或路径无效时抛出 ``TypeError``、``ValueError``、
        ``InvalidFileError`` 或文件系统异常。
        """
        from ..writer.delimited import write_delimited

        write_delimited(
            self, filename, encoding=encoding, delimiter=delimiter, formulas=formulas
        )
        return self

    def add_chart(self, chart_type: str, *, anchor: str) -> Chart:
        """功能：在当前工作表添加一个基础 XLSX 图表。

        使用方法：``chart = ws.add_chart(ChartType.COLUMN, anchor='E2')``。
        参数：``chart_type`` 为 ``ChartType.COLUMN``、``BAR``、``LINE``、``PIE``
        常量之一；``anchor`` 为图表左上角单个 A1 单元格地址。
        返回：新建 ``Chart``；调用其 ``add_series()`` 后保存为 XLSX 才会显示数据。
        异常：图表类型或锚点无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(chart_type, str):
            raise TypeError("chart_type 必须是 ChartType 字符串常量")
        if chart_type not in {ChartType.COLUMN, ChartType.BAR, ChartType.LINE, ChartType.PIE}:
            raise ValueError("不支持的图表类型")
        row, column = cell_index(anchor)
        chart = Chart(self, chart_type, cell_address(row, column))
        self._charts.append(chart)
        return chart

    @property
    def charts(self) -> tuple[Chart, ...]:
        """功能：取得当前工作表全部图表的只读顺序快照。

        使用方法：``for chart in worksheet.charts: ...``。
        参数：无。
        返回：按添加顺序排列的 ``tuple[Chart, ...]``，不可直接修改。
        """
        return tuple(self._charts)

    def _remove_chart(self, chart: Chart) -> None:
        """功能：删除当前工作表持有的图表对象。

        使用方法：由 ``Chart.remove()`` 内部调用。
        参数：``chart`` 为当前工作表创建的图表对象。
        返回：``None``。
        异常：图表不属于当前工作表时抛出 ``ValueError``。
        """
        if chart not in self._charts:
            raise ValueError("图表不属于当前工作表")
        self._charts.remove(chart)

    def add_image(
        self,
        source: str | os.PathLike[str] | bytes,
        *,
        anchor: str,
        name: str | None = None,
        placement: str = ImagePlacement.FLOATING,
        fit: str = ImageFit.STRETCH,
    ) -> Image:
        """功能：在当前工作表添加 PNG 或 JPEG 图片。

        使用方法：``image = ws.add_image('logo.png', anchor='A1')``；内存二维码可用
        ``ws.add_image(payload, anchor='A1', name='qrcode.png')``；需要图片随区域
        缩放时使用 ``ws.add_image('logo.png', anchor='B2:F8',
        placement=ImagePlacement.CELL, fit=ImageFit.CONTAIN)``。
        参数：``source`` 为字符串、``PathLike`` 图片路径或 PNG/JPEG ``bytes``；
        ``anchor`` 为单个 A1 地址或包含首尾单元格的矩形区域；二进制图片需要
        ``name`` 提供文件名；``placement`` 为 ``ImagePlacement.FLOATING``（默认，
        固定像素尺寸）或 ``ImagePlacement.CELL``（随锚点区域移动并缩放）；``fit``
        为 ``ImageFit.STRETCH``、``CONTAIN`` 或 ``COVER``，分别表示拉伸、等比完整
        显示和等比铺满裁剪。
        返回：新建 ``Image``；其 ``width``、``height`` 为像素，``offset_x``、
        ``offset_y`` 为像素偏移，均可在保存前修改；图片按添加顺序保存在
        ``worksheet.images`` 中。
        异常：路径、格式或锚点无效时抛出文件系统异常、``TypeError`` 或 ``ValueError``。
        """
        if not isinstance(source, (str, os.PathLike, bytes)):
            raise TypeError("source 必须是字符串、PathLike 图片路径或 bytes")
        try:
            bounds = range_index(anchor)
        except ValueError:
            row, column = cell_index(anchor)
            bounds = (row, column, row, column)
        canonical = range_address(*bounds)
        image = Image(
            self,
            source,
            canonical,
            name=name,
            placement=placement,
            fit=fit,
        )
        self._images.append(image)
        return image

    @property
    def images(self) -> tuple[Image, ...]:
        """功能：取得当前工作表全部图片的只读顺序快照。

        使用方法：``for image in worksheet.images: ...``。
        参数：无。
        返回：按添加顺序排列的 ``tuple[Image, ...]``，不可直接修改。
        """
        return tuple(self._images)

    def image(self, anchor_or_index: str | int) -> Image:
        """功能：按锚点地址或添加顺序取得一张图片。

        使用方法：``image = worksheet.image('B2:F8')``；也可用
        ``image = worksheet.image(0)`` 取得第一张图片。字符串会先规范化为
        A1 单格或矩形区域，并与图片的完整 ``anchor`` 匹配；查询区域锚点的左上角
        单元格（例如图片锚点为 ``B2:F8`` 时查询 ``B2``）也可以命中。
        参数：``anchor_or_index`` 为单格/区域 A1 字符串，或图片添加顺序的
        0-based 非负整数；布尔值不作为索引。
        返回：匹配到的 ``Image`` 对象；同一锚点有多张图片时返回最早添加者。
        异常：地址不存在时抛出 ``KeyError``；整数索引无效时抛出 ``IndexError``；
        参数类型无效时抛出 ``TypeError``。
        """
        if isinstance(anchor_or_index, bool):
            raise TypeError("图片索引必须是0-based整数或 A1 地址")
        if isinstance(anchor_or_index, int):
            if anchor_or_index < 0 or anchor_or_index >= len(self._images):
                raise IndexError("图片索引超出范围")
            return self._images[anchor_or_index]
        if not isinstance(anchor_or_index, str):
            raise TypeError("图片查询参数必须是0-based整数或 A1 地址")
        try:
            try:
                query_bounds = range_index(anchor_or_index)
            except ValueError:
                row, column = cell_index(anchor_or_index)
                query_bounds = (row, column, row, column)
            query = range_address(*query_bounds)
        except (TypeError, ValueError) as error:
            raise KeyError(anchor_or_index) from error
        query_top_left = query_bounds[:2]
        for item in self._images:
            if item.anchor.casefold() == query.casefold() or item.bounds[:2] == query_top_left:
                return item
        raise KeyError(anchor_or_index)

    def _remove_image(self, image: Image) -> None:
        """功能：删除当前工作表持有的图片对象。

        使用方法：由 ``Image.remove()`` 内部调用。
        参数：``image`` 为当前工作表创建的图片对象。
        返回：``None``。
        异常：图片不属于当前工作表时抛出 ``ValueError``。
        """
        if image not in self._images:
            raise ValueError("图片不属于当前工作表")
        self._images.remove(image)

    def add_table(
        self,
        address: str,
        *,
        name: str,
        style: str = "TableStyleMedium2",
        has_header: bool = True,
        show_row_stripes: bool = True,
        show_column_stripes: bool = False,
    ) -> Table:
        """功能：在连续区域上创建工作簿内名称唯一的 Excel 数据表。

        使用方法：``worksheet.add_table("A1:F10", name="SalesTable")``。
        参数：``address`` 为A1矩形区域；``name`` 为工作簿内唯一表名；``style``
        为Excel表样式名称；其余布尔值控制表头和行列条纹。
        返回：新创建的 :class:`Table`。
        异常：地址、名称、类型或区域重叠无效时抛出 ``ValueError`` 或 ``TypeError``。
        """
        area = self.range(address)
        self._workbook._validate_table_name(name)
        if self._workbook._table(name) is not None:
            raise ValueError(f"数据表名称已经存在：{name!r}")
        for existing in self._tables.values():
            other = existing.range
            separated = (
                area.max_row < other.min_row
                or area.min_row > other.max_row
                or area.max_column < other.min_column
                or area.min_column > other.max_column
            )
            if not separated:
                raise ValueError("同一工作表中的数据表区域不能重叠")
        table = Table(
            self,
            name,
            area,
            style=style,
            has_header=has_header,
            show_row_stripes=show_row_stripes,
            show_column_stripes=show_column_stripes,
        )
        self._tables[name.casefold()] = table
        self._touch(area.max_row, area.max_column)
        return table

    def table(self, name: str) -> Table:
        """功能：按大小写不敏感名称取得当前工作表的数据表。

        使用方法：``table = worksheet.table("SalesTable")``。
        参数：``name`` 为数据表名称字符串。
        返回：匹配的 :class:`Table`。
        异常：名称不存在时抛出 ``KeyError``；类型错误时抛出 ``TypeError``。
        """
        if not isinstance(name, str):
            raise TypeError("name 必须是字符串")
        try:
            return self._tables[name.casefold()]
        except KeyError:
            raise KeyError(name) from None

    @property
    def tables(self) -> tuple[Table, ...]:
        """功能：取得当前工作表全部数据表的只读顺序快照。

        使用方法：``for table in worksheet.tables: ...``。
        参数：无，只读属性。
        返回：按创建或读取顺序排列的 ``tuple[Table, ...]``。
        """
        return tuple(self._tables.values())

    def remove_table(self, name: str) -> "Worksheet":
        """功能：按名称删除当前工作表中的一个数据表定义。

        使用方法：``worksheet.remove_table("SalesTable")``。
        参数：``name`` 为大小写不敏感名称字符串。
        返回：当前 :class:`Worksheet`；单元格值和样式不删除。
        异常：名称不存在或类型无效时透传 :meth:`table` 的异常。
        """
        table = self.table(name)
        self._tables.pop(table.name.casefold())
        return self

    @property
    def max_row(self) -> int:
        """功能：取得工作表已经触及的最大 0-based 行索引。

        使用方法：A10 被写入后 ``worksheet.max_row`` 返回 ``9``。
        参数：无。
        返回：最大 0-based 行索引；空工作表返回 ``-1``。
        """
        return self._max_row

    @property
    def max_column(self) -> int:
        """功能：取得工作表已经触及的最大 0-based 列索引。

        使用方法：F1 被写入后 ``worksheet.max_column`` 返回 ``5``。
        参数：无。
        返回：最大 0-based 列索引；空工作表返回 ``-1``。
        """
        return self._max_column

    @property
    def values(self) -> list[list[Any]]:
        """功能：读取工作表当前已经触及范围内的全部有效值。

        使用方法：``data = worksheet.values``。
        参数：无，只读属性；需要写入数据时使用单元格、``append``、
        ``append_rows`` 或 ``Range.set_values``。
        返回：从 A1 到 ``max_row``、``max_column`` 的二维列表；空工作表返回
        ``[]``。公式单元格返回当前公式结果；没有有效结果时返回 ``None``。
        """
        if self._max_row < 0:
            return []
        end_address = cell_address(self._max_row, self._max_column)
        return self.range(f"A1:{end_address}").values

    @property
    def used_range(self) -> Range | None:
        """功能：返回工作表实际使用内容覆盖的最小矩形区域。

        使用方法：``area = worksheet.used_range``；空工作表返回 ``None``。
        参数：无，只读属性。普通值、公式、样式、链接、批注、合并区域、数据表、
        验证、条件格式以及图表和图片锚点都会计入使用范围。
        返回：覆盖所有已使用对象的 :class:`Range`；清空单元格后不会受历史最大索引影响。
        """
        coordinates: set[tuple[int, int]] = set(self._values._values) | set(self._formulas)
        coordinates |= set(self._styles) | set(self._hyperlinks) | set(self._notes)
        for min_row, min_column, max_row, max_column in self._merged_ranges:
            coordinates.update({(min_row, min_column), (max_row, max_column)})
        ranges: list[str] = [table.range.address for table in self._tables.values()]
        ranges.extend(item.range for item in self._validations)
        ranges.extend(item.range for item in self._conditionals)
        if self._filter_range:
            ranges.append(self._filter_range)
        for address in ranges:
            try:
                min_row, min_column, max_row, max_column = range_index(address)
            except (TypeError, ValueError):
                continue
            coordinates.update({(min_row, min_column), (max_row, max_column)})
        for item in (*self._charts, *self._images):
            try:
                bounds = range_index(item.anchor)
            except (TypeError, ValueError):
                try:
                    row, column = cell_index(item.anchor)
                    bounds = (row, column, row, column)
                except (TypeError, ValueError):
                    continue
            coordinates.update({(bounds[0], bounds[1]), (bounds[2], bounds[3])})
        if not coordinates:
            return None
        rows = [item[0] for item in coordinates]
        columns = [item[1] for item in coordinates]
        return Range(self, min(rows), min(columns), max(rows), max(columns))

    def validate(self) -> list[str]:
        """功能：检查工作表结构并返回可读的问题列表。

        使用方法：``problems = worksheet.validate()``；列表为空表示未发现问题。
        参数：无；该方法只读，不会修复或修改工作表。
        返回：中文问题描述字符串列表，包含地址、对象名称或公式依赖等上下文。
        """
        from .calculation import formula_dependencies

        problems: list[str] = []
        for bounds in self._merged_ranges:
            min_row, min_column, max_row, max_column = bounds
            if min_row > max_row or min_column > max_column:
                problems.append(f"合并区域边界无效：{bounds!r}")
        table_items = list(self._tables.values())
        for index, table in enumerate(table_items):
            try:
                table_bounds = range_index(table.range.address)
            except (TypeError, ValueError) as error:
                problems.append(f"数据表 {table.name!r} 地址无效：{error}")
                continue
            if table_bounds[0] > table_bounds[2] or table_bounds[1] > table_bounds[3]:
                problems.append(f"数据表 {table.name!r} 边界无效")
            for other in table_items[index + 1:]:
                try:
                    other_bounds = range_index(other.range.address)
                except (TypeError, ValueError):
                    continue
                if not (
                    table_bounds[2] < other_bounds[0]
                    or other_bounds[2] < table_bounds[0]
                    or table_bounds[3] < other_bounds[1]
                    or other_bounds[3] < table_bounds[1]
                ):
                    problems.append(f"数据表 {table.name!r} 与 {other.name!r} 区域重叠")
        for coordinate, formula in self._formulas.items():
            try:
                formula_dependencies(self._workbook, self, formula)
            except Exception as error:
                problems.append(f"公式 {self.cell(*coordinate).address} 无法解析：{error}")
        for item in (*self._validations, *self._conditionals):
            try:
                range_index(item.range)
            except (TypeError, ValueError) as error:
                problems.append(f"规则区域 {item.range!r} 无效：{error}")
        if self._filter_range:
            try:
                range_index(self._filter_range)
            except (TypeError, ValueError) as error:
                problems.append(f"筛选区域无效：{error}")
        if self._freeze:
            try:
                cell_index(self._freeze)
            except (TypeError, ValueError) as error:
                problems.append(f"冻结窗格地址无效：{error}")
        return problems

    def to_records(self, *, header_row: int = 0) -> list[dict[str, Any]]:
        """功能：把单一主数据表工作表转换为字典记录列表。

        使用方法：字段位于 Excel 第2行时调用 ``worksheet.to_records(header_row=1)``。
        参数：``header_row`` 为字段所在的绝对 0-based 行索引，默认 ``0``；字段行
        上方内容忽略，字段行首尾非空字段之间确定数据列。
        返回：字段行下方到工作表末行的记录列表；完全空白行会跳过，没有数据时返回
        ``[]``。普通值和公式结果均通过 ``Cell.value`` 统一读取。
        异常：工作表为空、字段行不存在、字段不连续、为空、重复或类型无效时抛出
        ``TypeError``、``ValueError`` 或 ``InvalidAddressError``。
        """
        validate_row_index(header_row)
        if self._max_row < header_row or self._max_column < 0:
            raise ValueError("工作表中不存在指定字段行")
        header_values = [
            self.cell(header_row, column).value
            for column in range(self._max_column + 1)
        ]
        populated = [
            index for index, value in enumerate(header_values)
            if value not in (None, "")
        ]
        if not populated:
            raise ValueError("字段行中没有字段名")
        first_column, last_column = populated[0], populated[-1]
        names = header_values[first_column:last_column + 1]
        if any(value in (None, "") for value in names):
            raise ValueError("字段行的首尾字段之间不能存在空字段")
        header_area = Range(
            self,
            header_row,
            first_column,
            header_row,
            last_column,
        )
        # 先统一验证字段名，再处理没有数据行的工作表。
        header_area.to_records()
        if self._max_row == header_row:
            return []
        area = Range(
            self,
            header_row + 1,
            first_column,
            self._max_row,
            last_column,
        )
        records = area.to_records(headers=names)
        result: list[dict[str, Any]] = []
        for offset, record in enumerate(records):
            row = header_row + 1 + offset
            has_formula = any(
                (row, column) in self._formulas
                for column in range(first_column, last_column + 1)
            )
            if has_formula or any(value not in (None, "") for value in record.values()):
                result.append(record)
        return result

    @staticmethod
    def _prepare_records(
        records: Iterable[Mapping[str, Any]], headers: bool | Sequence[str]
    ) -> tuple[list[Mapping[str, Any]], tuple[str, ...], bool]:
        """功能：完整验证字典记录并确定稳定的字段顺序。

        使用方法：由 ``write_records`` 和 ``write_table`` 在写入前共同调用。
        参数：``records`` 为映射记录迭代对象；``headers`` 为 ``True``、``False``
        或显式字段名称序列。
        返回：规范化的记录列表、字段名称元组及是否写入表头。
        异常：记录、字段名或表头参数不合法时抛出 ``TypeError`` 或 ``ValueError``。
        """
        try:
            prepared = list(records)
        except TypeError as error:
            raise TypeError("records 必须是映射记录的可迭代对象") from error
        if any(not isinstance(record, Mapping) for record in prepared):
            raise TypeError("records 中的每条记录必须是映射对象")
        if isinstance(headers, bool):
            include_headers = headers
            names: list[str] = []
            for record in prepared:
                for key in record:
                    if not isinstance(key, str) or not key:
                        raise ValueError("记录字段名必须是非空字符串")
                    if key not in names:
                        names.append(key)
        else:
            if isinstance(headers, (str, bytes)):
                raise TypeError("headers 必须是 bool 或字段名序列")
            try:
                names = list(headers)
            except TypeError as error:
                raise TypeError("headers 必须是 bool 或字段名序列") from error
            include_headers = True
            if not names or any(not isinstance(name, str) or not name for name in names):
                raise ValueError("headers 必须是非空字符串的非空序列")
            if len(set(names)) != len(names):
                raise ValueError("headers 不能包含重复字段名")
        if not names:
            raise ValueError("没有可写入的字段；空 records 时请提供 headers 字段序列")
        unknown = {
            key for record in prepared for key in record if key not in names
        }
        if unknown:
            raise ValueError(f"记录包含 headers 中不存在的字段：{sorted(unknown)!r}")
        return prepared, tuple(names), include_headers

    def write_records(
        self,
        row: int,
        column: int,
        records: Iterable[Mapping[str, Any]],
        *,
        headers: bool | Sequence[str] = True,
    ) -> Range:
        """功能：从指定 0-based 行列开始写入字典记录，字段名决定列顺序。

        使用方法：``ws.write_records(0, 0, orders)``；空记录需要显式传入
        ``headers=['订单号', '金额']``。设置 ``headers=False`` 可只写数据行。
        参数：``row``、``column`` 为起始 0-based 索引，顺序先行后列；``records``
        为字典或其他映射对象序列；``headers`` 可为 ``True``（按首次出现的字段顺序
        写表头）、``False``（不写表头）或字段名序列（写指定顺序表头）。
        返回：覆盖写入区域；只有表头时返回一行区域。
        异常：索引、字段、目标合并单元格或数据不合法时抛出 ``TypeError``、
        ``ValueError`` 或 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        prepared, names, include_headers = self._prepare_records(records, headers)
        row_values: list[list[Any]] = [list(names)] if include_headers else []
        row_values.extend([[record.get(name) for name in names] for record in prepared])
        if not row_values:
            raise ValueError("headers=False 时 records 不能为空")
        end_row = row + len(row_values) - 1
        end_column = column + len(names) - 1
        validate_row_index(end_row)
        validate_column_index(end_column)
        area = Range(self, row, column, end_row, end_column)
        area.set_values(row_values)
        return area

    def _next_table_name(self) -> str:
        """功能：生成当前工作簿中尚未使用的默认数据表名称。

        使用方法：由 :meth:`write_table` 在调用者省略名称时内部调用。
        参数：无；检查当前工作簿全部工作表中的 Table 名称。
        返回：形如 ``Table1``、``Table2`` 的唯一名称字符串。
        """
        index = 1
        while self._workbook._table(f"Table{index}") is not None:
            index += 1
        return f"Table{index}"

    def write_table(
        self,
        row: int,
        column: int,
        records: Iterable[Mapping[str, Any]],
        *,
        headers: Sequence[str] | None = None,
        name: str | None = None,
        style: str = "TableStyleMedium2",
        freeze_header: bool = False,
        auto_fit: bool = False,
    ) -> Table:
        """功能：一次写入字典记录并创建带筛选按钮的 Excel 数据表。

        使用方法：``table = ws.write_table(0, 0, orders, name='Orders',
        freeze_header=True, auto_fit=True)``。
        参数：``row``、``column`` 为 0-based 起始位置；``records`` 为映射记录；
        ``headers`` 可指定空数据时的字段顺序；``name`` 省略时自动生成 ``Table1``；
        ``style`` 为 Excel 表格样式；``freeze_header`` 冻结表头上方行；``auto_fit``
        自动调整写入列宽。
        返回：新建 :class:`Table`。
        异常：记录、区域、名称、样式或布尔参数无效时抛出相应异常。
        """
        if headers is not None and (isinstance(headers, (str, bytes)) or not isinstance(headers, Sequence)):
            raise TypeError("headers 必须是字段名序列或 None")
        if not isinstance(freeze_header, bool) or not isinstance(auto_fit, bool):
            raise TypeError("freeze_header 和 auto_fit 必须是 bool")
        area = self.write_records(row, column, records, headers=headers if headers is not None else True)
        table = self.add_table(area.address, name=name or self._next_table_name(), style=style)
        if freeze_header:
            validate_row_index(row + 1)
            self.freeze_panes = cell_address(row + 1, 0)
        if auto_fit:
            self.auto_fit_columns(area.min_column, area.max_column)
        return table

    @staticmethod
    def _display_width(value: Any) -> int:
        """功能：按中日韩全角字符宽度估算单元格显示宽度。

        使用方法：由自动列宽和自动行高计算内部调用。
        参数：``value`` 为要估算显示宽度的任意单元格内容。
        返回：多行内容中最宽一行的近似字符宽度整数。
        """
        text = str(value)
        return max(
            sum(2 if unicodedata.east_asian_width(char) in {"W", "F"} else 1 for char in line)
            for line in text.splitlines() or [""]
        )

    def auto_fit_columns(
        self,
        first_column: int,
        last_column: int,
        *,
        min_width: float = 0,
        max_width: float = 40,
    ) -> "Worksheet":
        """功能：根据已使用单元格内容自动设置连续列的宽度。

        使用方法：``ws.auto_fit_columns(0, 5, max_width=40)``。
        参数：``first_column``、``last_column`` 为包含式 0-based 列索引；
        ``min_width`` 为最小宽度，``max_width`` 为最大宽度，均为正数或零。
        返回：当前工作表。
        异常：索引、宽度或顺序无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        validate_column_index(first_column)
        validate_column_index(last_column)
        if first_column > last_column:
            raise ValueError("first_column 不能大于 last_column")
        for name, value in (("min_width", min_width), ("max_width", max_width)):
            if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0:
                raise ValueError(f"{name} 必须是大于等于0的数字")
        if min_width > max_width or max_width > 255:
            raise ValueError("宽度必须满足 0 <= min_width <= max_width <= 255")
        for column in range(first_column, last_column + 1):
            width = min_width
            for row in range(self._max_row + 1):
                value = self._formulas.get((row, column), self._values.get(row, column))
                if value is not None:
                    width = max(width, self._display_width(value) + 2)
            self.column(column).width = max(0.1, min(float(max_width), float(width)))
        return self

    def auto_fit_rows(
        self,
        first_row: int,
        last_row: int,
        *,
        min_height: float = 0,
        max_height: float = 120,
    ) -> "Worksheet":
        """功能：根据换行和当前列宽估算连续行的合适行高。

        使用方法：``ws.auto_fit_rows(0, 100)``。
        参数：``first_row``、``last_row`` 为包含式 0-based 行索引；``min_height``
        和 ``max_height`` 为磅值边界，0 表示不设下限。
        返回：当前工作表。
        异常：索引、边界或顺序无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        validate_row_index(first_row)
        validate_row_index(last_row)
        if first_row > last_row:
            raise ValueError("first_row 不能大于 last_row")
        for name, value in (("min_height", min_height), ("max_height", max_height)):
            if isinstance(value, bool) or not isinstance(value, (int, float)) or value < 0:
                raise ValueError(f"{name} 必须是大于等于0的数字")
        if min_height > max_height or max_height > 409:
            raise ValueError("高度必须满足 0 <= min_height <= max_height <= 409")
        for row in range(first_row, last_row + 1):
            lines = 1
            for column in range(self._max_column + 1):
                value = self._formulas.get((row, column), self._values.get(row, column))
                if value is None:
                    continue
                available = self.column(column).width or 8.43
                lines = max(lines, math.ceil(self._display_width(value) / available))
            self.row(row).height = max(0.1, min(float(max_height), max(float(min_height), lines * 15.0)))
        return self

    @property
    def horizontal_page_breaks(self) -> tuple[int, ...]:
        """功能：取得全部手动水平分页符所在的 0-based 行索引快照。

        使用方法：``breaks = worksheet.horizontal_page_breaks``。
        参数：无，只读属性；增删使用对应方法。
        返回：按升序排列的 0-based 行索引元组。
        """
        return tuple(sorted(self._horizontal_page_breaks))

    def add_horizontal_page_break(self, row: int) -> "Worksheet":
        """功能：在指定 0-based 行之前插入一个手动水平打印分页符。

        使用方法：``ws.add_horizontal_page_break(49)`` 让第 49 行从新页开始。
        参数：``row`` 为 1 以上的 0-based 行索引，``0`` 不能作为分页符位置。
        返回：当前工作表。
        异常：索引无效或为 ``0`` 时抛出 ``ValueError``。
        """
        validate_row_index(row)
        if row == 0:
            raise ValueError("分页符不能设置在第 0 行之前")
        self._horizontal_page_breaks.add(row)
        return self

    def remove_horizontal_page_break(self, row: int) -> "Worksheet":
        """功能：移除指定行之前的手动水平分页符。

        使用方法：``ws.remove_horizontal_page_break(49)``。
        参数：``row`` 为已设置分页符的 0-based 行索引。
        返回：当前工作表。
        异常：位置未设置分页符或索引无效时抛出 ``ValueError``。
        """
        validate_row_index(row)
        if row not in self._horizontal_page_breaks:
            raise ValueError("该位置不存在手动水平分页符")
        self._horizontal_page_breaks.remove(row)
        return self

    @staticmethod
    def _prepare_append_row(values: Iterable[Any], method_name: str) -> list[Any]:
        """功能：完整读取并规范化一行待追加数据。

        使用方法：由 ``append()`` 和 ``append_rows()`` 在写入前共同调用。
        参数：``values`` 为一维可迭代数据；``method_name`` 用于生成明确错误消息。
        返回：已经完成日期字面量转换的普通值列表。
        异常：整行是字符串、字节或不可迭代对象时抛出 ``TypeError``；值转换失败
        时透传相应异常，工作表尚未发生改变。
        """
        if isinstance(values, (str, bytes)):
            raise TypeError(f"{method_name} 需要一维行数据，不能直接传入字符串或字节对象")
        try:
            row_values = list(values)
        except TypeError as error:
            raise TypeError(f"{method_name} 需要一维可迭代对象") from error
        return [_normalize_value(value) for value in row_values]

    def append(self, values: Iterable[Any]) -> "Worksheet":
        """功能：在当前最大行索引之后追加一行普通值。

        使用方法：``worksheet.append(["姓名", "成绩"])``；空表从 A1 开始。
        参数：``values`` 为一维可迭代对象，元素按 0-based 列索引从左到右写入；
        字符串和字节对象不能作为整行数据，空可迭代对象不会触及新行。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：参数不可迭代或是字符串、字节对象时抛出 ``TypeError``；数据超过
        Excel 行列上限时抛出 ``InvalidAddressError``；值转换失败或目标位于合并
        区域非左上角时保持整行写入前状态。
        """
        row_values = self._prepare_append_row(values, "append()")
        if not row_values:
            return self

        target_row = self._max_row + 1
        validate_row_index(target_row)
        validate_column_index(len(row_values) - 1)
        for column in range(len(row_values)):
            anchor = self._merged_anchor(target_row, column)
            if anchor is not None and anchor != (target_row, column):
                raise ValueError("不能向合并区域的非左上角单元格追加值")
        for column, value in enumerate(row_values):
            self._set_value(target_row, column, value)
        return self

    def append_rows(self, rows: Iterable[Iterable[Any]]) -> "Worksheet":
        """功能：按给定顺序连续追加多行普通值。

        使用方法：``worksheet.append_rows([["张三", 90], ["李四", 88]])``。
        参数：``rows`` 为二维可迭代对象；每个非空元素会作为一行传给
        :meth:`append`，空行与单行 ``append([])`` 一样不会推进位置。
        返回：当前 :class:`Worksheet`，支持链式调用。
        异常：外层或任一行不可迭代时抛出 ``TypeError``；超过 Excel 上限时抛出
        ``InvalidAddressError``。全部行和值会在首次写入前完成验证，任何失败都不会
        留下前置行或部分单元格。
        """
        if isinstance(rows, (str, bytes)):
            raise TypeError("append_rows() 需要二维数据，不能直接传入字符串或字节对象")
        try:
            input_rows = list(rows)
        except TypeError as error:
            raise TypeError("append_rows() 需要二维可迭代对象") from error
        prepared_rows = [
            self._prepare_append_row(values, "append_rows()")
            for values in input_rows
        ]
        prepared_rows = [values for values in prepared_rows if values]
        if not prepared_rows:
            return self

        first_row = self._max_row + 1
        validate_row_index(first_row + len(prepared_rows) - 1)
        for row_offset, row_values in enumerate(prepared_rows):
            validate_column_index(len(row_values) - 1)
            target_row = first_row + row_offset
            for column in range(len(row_values)):
                anchor = self._merged_anchor(target_row, column)
                if anchor is not None and anchor != (target_row, column):
                    raise ValueError("不能向合并区域的非左上角单元格追加值")
        for row_values in prepared_rows:
            self.append(row_values)
        return self

    def _touch(self, row: int, column: int) -> None:
        """功能：更新工作表已经触及的最大 0-based 行列索引。

        使用方法：由普通值和公式写入方法内部调用，业务代码不应直接依赖。
        参数：``row``、``column`` 为有效 0-based 整数索引，顺序为先行后列。
        返回：``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        self._max_row = max(self._max_row, row)
        self._max_column = max(self._max_column, column)

    def _merged_anchor(self, row: int, column: int) -> Optional[Tuple[int, int]]:
        """功能：查询坐标所属合并区域的左上角锚点。

        使用方法：普通值和公式写入前内部调用。
        参数：``row``、``column`` 为0-based行列索引，顺序先行后列。
        返回：属于合并区域时返回锚点元组，否则返回 ``None``。
        """
        for min_row, min_column, max_row, max_column in self._merged_ranges:
            if min_row <= row <= max_row and min_column <= column <= max_column:
                return min_row, min_column
        return None

    def _merge_range(
        self, min_row: int, min_column: int, max_row: int, max_column: int
    ) -> None:
        """功能：原子登记一个不重叠的多单元格合并区域。

        使用方法：仅由 ``Range.merge()`` 调用。
        参数：四项为0-based最小行、最小列、最大行、最大列，顺序先行后列。
        返回：``None``；完全相同的区域重复合并视为幂等操作。
        异常：单格区域、重叠区域或非锚点单元格存在值或公式时抛出 ``ValueError``。
        """
        bounds = (min_row, min_column, max_row, max_column)
        if min_row == max_row and min_column == max_column:
            raise ValueError("合并区域必须至少包含两个单元格")
        for existing in self._merged_ranges:
            if existing == bounds:
                return
            a, b, c, d = existing
            overlaps = not (
                max_row < a or min_row > c or max_column < b or min_column > d
            )
            if overlaps:
                raise ValueError("合并区域不能与已有合并区域重叠")
        for row in range(min_row, max_row + 1):
            for column in range(min_column, max_column + 1):
                if (row, column) == (min_row, min_column):
                    continue
                if (
                    self._values.get(row, column) is not None
                    or (row, column) in self._formulas
                ):
                    raise ValueError("合并前除左上角外的单元格必须为空")
        self._merged_ranges.append(bounds)
        self._merged_ranges.sort()
        self._touch(max_row, max_column)

    def _unmerge_range(
        self, min_row: int, min_column: int, max_row: int, max_column: int
    ) -> None:
        """功能：删除与给定边界完全相同的合并区域记录。

        使用方法：仅由 ``Range.unmerge()`` 调用。
        参数：四项为0-based区域边界，顺序先行后列。
        返回：``None``；单元格内容和样式不改变。
        异常：区域没有被完整合并时抛出 ``ValueError``。
        """
        bounds = (min_row, min_column, max_row, max_column)
        if bounds not in self._merged_ranges:
            raise ValueError("当前区域不是一个完整的合并区域")
        self._merged_ranges.remove(bounds)

    def _set_value(self, row: int, column: int, value: Any) -> None:
        """功能：设置普通值并清除同一位置的公式。

        使用方法：由 :attr:`Cell.value`、区域写入和追加方法内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列；
        ``value`` 为任意 Python 对象，``None`` 表示删除普通值。
        返回：``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        anchor = self._merged_anchor(row, column)
        if anchor is not None and anchor != (row, column):
            raise ValueError("只能向合并区域的左上角单元格写入值")
        normalized_value = _normalize_value(value)
        self._workbook._invalidate_formula_caches()
        self._touch(row, column)
        self._formulas.pop((row, column), None)
        self._formula_values.pop((row, column), None)
        self._formula_errors.pop((row, column), None)
        self._values.set(row, column, normalized_value)

    def _get_formula(self, row: int, column: int) -> Optional[str]:
        """功能：读取指定位置的标准化公式。

        使用方法：由 :attr:`Cell.formula` 读取器内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列。
        返回：带前导 ``=`` 的公式字符串；没有公式时返回 ``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        return self._formulas.get((row, column))

    def _get_hyperlink(self, row: int, column: int) -> Optional[Hyperlink]:
        """功能：读取指定位置的超链接对象。

        使用方法：由 ``Cell.hyperlink`` 读取器内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列。
        返回：对应 ``Hyperlink`` 或 ``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        return self._hyperlinks.get((row, column))

    def _set_hyperlink(
        self, row: int, column: int, value: Optional[Hyperlink | str]
    ) -> None:
        """功能：校验并设置指定位置的超链接，或清除现有链接。

        使用方法：由 ``Cell.hyperlink`` 设置器内部调用。
        参数：``row``、``column`` 为 0-based 索引，顺序为先行后列；``value`` 为
        ``Hyperlink``、网址字符串或 ``None``。
        返回：``None``；普通值、公式和样式保持不变。
        异常：类型、链接内容或合并位置无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        anchor = self._merged_anchor(row, column)
        if anchor is not None and anchor != (row, column):
            raise ValueError("只能向合并区域的左上角单元格设置超链接")
        if value is None:
            self._hyperlinks.pop((row, column), None)
            return
        if isinstance(value, str):
            value = Hyperlink(target=value)
        elif not isinstance(value, Hyperlink):
            raise TypeError("hyperlink 必须是字符串、Hyperlink 或 None")
        self._hyperlinks[(row, column)] = value
        self._touch(row, column)

    def _get_note(self, row: int, column: int) -> Optional[Note]:
        """功能：读取指定位置的传统批注对象。

        使用方法：由 ``Cell.note`` 读取器内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列。
        返回：对应 ``Note`` 或 ``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        return self._notes.get((row, column))

    def _set_note(self, row: int, column: int, value: Optional[Note | str]) -> None:
        """功能：校验并设置或清除指定位置的传统批注。

        使用方法：由 ``Cell.note`` 设置器内部调用。
        参数：``value`` 可为 ``Note``、字符串正文或 ``None``；字符串使用默认作者
        ``ExcelKit``。``row``、``column`` 为 0-based 坐标。
        返回：``None``。
        异常：类型、批注内容或合并区域位置无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        anchor = self._merged_anchor(row, column)
        if anchor is not None and anchor != (row, column):
            raise ValueError("只能向合并区域的左上角单元格设置批注")
        if value is None:
            self._notes.pop((row, column), None)
        elif isinstance(value, Note):
            self._notes[(row, column)] = value
            self._touch(row, column)
        elif isinstance(value, str):
            self._notes[(row, column)] = Note(value)
            self._touch(row, column)
        else:
            raise TypeError("note 必须是 Note、字符串或 None")

    @property
    def hyperlinks(self) -> tuple[tuple[str, Hyperlink], ...]:
        """功能：取得当前工作表全部超链接的只读快照。

        使用方法：``for address, link in worksheet.hyperlinks: ...``。
        参数：无，只读属性。
        返回：按行优先顺序排列的 ``(A1地址, Hyperlink)`` 元组；没有链接时返回空元组。
        """
        return tuple(
            (cell_address(row, column), link)
            for (row, column), link in sorted(self._hyperlinks.items())
        )

    @staticmethod
    def _validate_edit_count(count: int) -> None:
        """功能：验证行列插入或删除数量。

        使用方法：由 ``insert_*`` 和 ``delete_*`` 内部调用。
        参数：``count`` 为正整数数量，布尔值不作为整数处理。
        返回：``None``。
        异常：数量不是正整数时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if isinstance(count, bool) or not isinstance(count, int):
            raise TypeError("count 必须是正整数")
        if count < 1:
            raise ValueError("count 必须是正整数")

    def _validate_edit_index(self, index: int, *, deleting: bool, rows: bool) -> None:
        """功能：验证行列编辑位置是否符合当前工作表边界。

        使用方法：由行列编辑公共方法内部调用。
        参数：``index`` 为0-based位置；``deleting`` 表示是否为删除操作；``rows``
        表示按行还是按列验证边界。
        返回：``None``。
        异常：类型错误、负数或删除空白范围时抛出 ``TypeError`` 或 ``ValueError``。
        """
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError("index 必须是0-based整数")
        axis_max = self._max_row if rows else self._max_column
        maximum = axis_max if deleting else axis_max + 1
        if index < 0 or index > maximum:
            raise ValueError("index 超出当前工作表允许范围")

    @staticmethod
    def _map_index(
        value: int, index: int, count: int, *, deleting: bool
    ) -> Optional[int]:
        """功能：把一个行或列索引映射到插入或删除后的索引。

        使用方法：行列内部状态重建时调用。
        参数：``value`` 为原0-based索引；``index`` 为编辑起点；``count`` 为数量；
        ``deleting`` 表示删除而非插入。
        返回：新索引；被删除范围内的索引返回 ``None``。
        """
        if not deleting:
            return value + count if value >= index else value
        if index <= value < index + count:
            return None
        return value - count if value >= index + count else value

    @classmethod
    def _map_bounds(
        cls,
        bounds: Tuple[int, int, int, int],
        index: int,
        count: int,
        *,
        rows: bool,
        deleting: bool,
    ) -> Optional[Tuple[int, int, int, int]]:
        """功能：映射矩形边界并处理插入扩张或删除收缩。

        使用方法：合并区域、Table 和命名区域同步时内部调用。
        参数：``bounds`` 为最小行、最小列、最大行、最大列；``index``、``count``
        为编辑位置和数量；``rows`` 表示按行编辑；``deleting`` 表示删除。
        返回：调整后的边界；全部范围被删除时返回 ``None``。
        """
        min_row, min_column, max_row, max_column = bounds
        if rows:
            start, end = min_row, max_row
        else:
            start, end = min_column, max_column
        if not deleting:
            if start >= index:
                start += count
                end += count
            elif end >= index:
                end += count
        else:
            deleted_end = index + count - 1
            if end < index:
                pass
            elif start > deleted_end:
                start -= count
                end -= count
            else:
                remaining: list[tuple[int, int]] = []
                if start < index:
                    remaining.append((start, index - 1))
                if end > deleted_end:
                    remaining.append((index, end - count))
                if not remaining:
                    return None
                start = min(item[0] for item in remaining)
                end = max(item[1] for item in remaining)
        if rows:
            return start, min_column, end, max_column
        return min_row, start, max_row, end

    @staticmethod
    def _shift_formula_references(
        formula: str,
        index: int,
        count: int,
        *,
        rows: bool,
        deleting: bool,
        current_sheet: str,
        formula_sheet: str,
    ) -> str:
        """功能：同步行列编辑对公式 A1 引用造成的位移。

        使用方法：工作表插入或删除行列时由工作簿内部调用。
        参数：``formula`` 为原公式；其余参数描述编辑轴、位置、数量、当前工作表
        和公式所在工作表。绝对引用也会随被编辑区域移动，符合 Excel 插入行为。
        返回：调整后的公式；引用被删除时使用 ``#REF!``。
        """
        pattern = re.compile(
            r"(?<![\w.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?"
            r"(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)(?![\w.])"
        )
        current_key = current_sheet.casefold()

        def replace(match: re.Match[str]) -> str:
            """功能：替换公式中的一个单格引用。

            使用方法：由正则替换器自动调用。
            参数：``match`` 为引用正则匹配对象。
            返回：原引用或调整后的 A1 引用文本。
            """
            quoted, plain, column_absolute, column, row_absolute, row_text = match.groups()
            referenced_sheet = (quoted or plain or formula_sheet).replace("''", "'")
            if referenced_sheet.casefold() != current_key:
                return match.group(0)
            number = int(row_text) if rows else column_to_index(column)
            start = index + 1 if rows else index
            end = index + count if rows else index + count - 1
            if deleting and start <= number <= end:
                return "#REF!"
            if (not deleting and number >= start) or (deleting and number > end):
                number += count if not deleting else -count
            if rows:
                if not 1 <= number <= 1048576:
                    return "#REF!"
                if quoted is not None:
                    prefix = f"'{quoted}'!"
                elif plain is not None:
                    prefix = f"{plain}!"
                else:
                    prefix = ""
                return prefix + column_absolute + column + row_absolute + str(number)
            if not 0 <= number < 16384:
                return "#REF!"
            if quoted is not None:
                prefix = f"'{quoted}'!"
            elif plain is not None:
                prefix = f"{plain}!"
            else:
                prefix = ""
            return prefix + column_absolute + index_to_column(number) + row_absolute + row_text

        pieces = re.split(r'("(?:[^"]|"")*")', formula)
        return "".join(
            piece if offset % 2 else pattern.sub(replace, piece)
            for offset, piece in enumerate(pieces)
        )

    def _edit_axis(self, index: int, count: int, *, rows: bool, deleting: bool) -> "Worksheet":
        """功能：统一执行行列插入或删除并重建工作表稀疏状态。

        使用方法：由 ``insert_rows``、``delete_rows``、``insert_columns`` 和
        ``delete_columns`` 调用；业务代码应使用四个公开方法。
        参数：``index`` 为0-based起点；``count`` 为数量；``rows`` 选择行或列；
        ``deleting`` 选择删除或插入。
        返回：当前 ``Worksheet``，支持链式调用。
        异常：参数无效或删除范围超出边界时抛出 ``TypeError`` 或 ``ValueError``。
        """
        self._validate_edit_count(count)
        self._validate_edit_index(index, deleting=deleting, rows=rows)
        if deleting:
            maximum = self._max_row if rows else self._max_column
            if index + count - 1 > maximum:
                raise ValueError("删除范围超出当前工作表边界")
        axis_max = self._max_row if rows else self._max_column
        coordinate_maps = (
            self._values._values,
            self._formulas,
            self._formula_values,
            self._formula_errors,
            self._hyperlinks,
            self._notes,
            self._styles,
        )
        for mapping in coordinate_maps:
            rebuilt = {}
            for (row, column), value in mapping.items():
                selected = row if rows else column
                mapped = self._map_index(selected, index, count, deleting=deleting)
                if mapped is None:
                    continue
                coordinate = (mapped, column) if rows else (row, mapped)
                rebuilt[coordinate] = value
            mapping.clear()
            mapping.update(rebuilt)
        if rows:
            dimensions, max_name = self._rows, "_max_row"
        else:
            dimensions, max_name = self._columns, "_max_column"
        rebuilt_dimensions = {}
        for selected, dimension in dimensions.items():
            mapped = self._map_index(selected, index, count, deleting=deleting)
            if mapped is not None:
                dimension._index = mapped
                rebuilt_dimensions[mapped] = dimension
        dimensions.clear()
        dimensions.update(rebuilt_dimensions)
        mapped_ranges = []
        for bounds in self._merged_ranges:
            mapped = self._map_bounds(bounds, index, count, rows=rows, deleting=deleting)
            if mapped is not None and not (mapped[0] == mapped[2] and mapped[1] == mapped[3]):
                mapped_ranges.append(mapped)
        self._merged_ranges = sorted(mapped_ranges)
        for table in self.tables:
            mapped = self._map_bounds(table._bounds, index, count, rows=rows, deleting=deleting)
            if mapped is None:
                self.remove_table(table.name)
            else:
                table._bounds = mapped
        for named_range in self._workbook.named_ranges:
            if named_range.worksheet is self:
                mapped = self._map_bounds(named_range._bounds, index, count, rows=rows, deleting=deleting)
                if mapped is not None:
                    named_range._bounds = mapped
        mapped_validations = []
        for item in self._validations:
            mapped = self._map_bounds(range_index(item.range), index, count, rows=rows, deleting=deleting)
            if mapped is not None:
                item.range = range_address(*mapped)
                mapped_validations.append(item)
        self._validations = mapped_validations
        mapped_conditionals = []
        for item in self._conditionals:
            mapped = self._map_bounds(range_index(item.range), index, count, rows=rows, deleting=deleting)
            if mapped is not None:
                item.range = range_address(*mapped)
                mapped_conditionals.append(item)
        self._conditionals = mapped_conditionals
        mapped_images = []
        for image in self._images:
            mapped = self._map_bounds(
                image.bounds, index, count, rows=rows, deleting=deleting
            )
            if mapped is None:
                continue
            image.anchor = range_address(*mapped)
            mapped_images.append(image)
        self._images = mapped_images
        if self._freeze is not None:
            freeze_row, freeze_column = cell_index(self._freeze)
            mapped = self._map_index(freeze_row if rows else freeze_column, index, count, deleting=deleting)
            if mapped is not None:
                self._freeze = cell_address(mapped, freeze_column) if rows else cell_address(freeze_row, mapped)
        if self._filter_range is not None:
            bounds = range_index(self._filter_range)
            mapped = self._map_bounds(bounds, index, count, rows=rows, deleting=deleting)
            self._filter_range = range_address(*mapped) if mapped is not None else None
        if self._page.print_area is not None:
            bounds = range_index(self._page.print_area)
            mapped = self._map_bounds(bounds, index, count, rows=rows, deleting=deleting)
            self._page.print_area = range_address(*mapped) if mapped is not None else None
        if rows and self._page.repeat_rows is not None:
            start, end = self._page.repeat_rows
            mapped = self._map_bounds((start, 0, end, 0), index, count, rows=True, deleting=deleting)
            self._page.repeat_rows = None if mapped is None else (mapped[0], mapped[2])
        if not rows and self._page.repeat_columns is not None:
            start, end = self._page.repeat_columns
            mapped = self._map_bounds((0, start, 0, end), index, count, rows=False, deleting=deleting)
            self._page.repeat_columns = None if mapped is None else (mapped[1], mapped[3])
        self._workbook._shift_formula_references(
            self, index, count, rows=rows, deleting=deleting
        )
        if deleting:
            setattr(self, max_name, max(-1, axis_max - count))
        else:
            setattr(self, max_name, axis_max + count)
        self._workbook._invalidate_formula_caches()
        return self

    def insert_rows(self, index: int, count: int = 1) -> "Worksheet":
        """功能：在指定0-based行索引前插入一个或多个空行。

        使用方法：``worksheet.insert_rows(2, count=3)``。
        参数：``index`` 为插入位置，允许取到当前最大行索引加1；``count`` 为正整数。
        返回：当前工作表。
        异常：参数无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return self._edit_axis(index, count, rows=True, deleting=False)

    def delete_rows(self, index: int, count: int = 1) -> "Worksheet":
        """功能：删除指定0-based起点开始的连续行。

        使用方法：``worksheet.delete_rows(2, count=3)``。
        参数：``index`` 为删除起点；``count`` 为正整数，删除范围必须已触及。
        返回：当前工作表。
        异常：范围越界或参数无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return self._edit_axis(index, count, rows=True, deleting=True)

    def insert_columns(self, index: int, count: int = 1) -> "Worksheet":
        """功能：在指定0-based列索引前插入一个或多个空列。

        使用方法：``worksheet.insert_columns(1, count=2)``。
        参数：``index`` 为插入位置，允许取到当前最大列索引加1；``count`` 为正整数。
        返回：当前工作表。
        异常：参数无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return self._edit_axis(index, count, rows=False, deleting=False)

    def delete_columns(self, index: int, count: int = 1) -> "Worksheet":
        """功能：删除指定0-based起点开始的连续列。

        使用方法：``worksheet.delete_columns(1, count=2)``。
        参数：``index`` 为删除起点；``count`` 为正整数，删除范围必须已触及。
        返回：当前工作表。
        异常：范围越界或参数无效时抛出 ``TypeError`` 或 ``ValueError``。
        """
        return self._edit_axis(index, count, rows=False, deleting=True)

    def _set_formula(
        self, row: int, column: int, formula: Optional[str], *, invalidate: bool = True
    ) -> None:
        """功能：校验并设置公式，或使用 ``None`` 清除现有公式。

        使用方法：通过 ``worksheet["A1"].formula = "=SUM(B1:B5)"`` 间接调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列；
        ``formula`` 为非空字符串或 ``None``，字符串可以包含或省略开头的 ``=``。
        返回：``None``；字符串统一保存为带前导 ``=`` 的形式，``None`` 仅删除公式。
        异常：非空公式不是字符串、为空或只有 ``=`` 时抛出 ``TypeError``；索引
        无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        if formula is None:
            if invalidate:
                self._workbook._invalidate_formula_caches()
            self._formulas.pop((row, column), None)
            self._formula_values.pop((row, column), None)
            self._formula_errors.pop((row, column), None)
            return
        anchor = self._merged_anchor(row, column)
        if anchor is not None and anchor != (row, column):
            raise ValueError("只能向合并区域的左上角单元格写入公式")
        if not isinstance(formula, str):
            raise TypeError("公式必须是非空字符串")
        expression = formula.strip()
        if expression.startswith("="):
            expression = expression[1:].strip()
        if not expression:
            raise TypeError("公式必须包含表达式")
        if invalidate:
            self._workbook._invalidate_formula_caches()
        self._touch(row, column)
        self._values.set(row, column, None)
        self._formulas[(row, column)] = f"={expression}"
        self._formula_values.pop((row, column), None)
        self._formula_errors.pop((row, column), None)

    def _get_cached_value(self, row: int, column: int) -> Any:
        """功能：读取公式最近一次有效计算结果的内部缓存。

        使用方法：由 :attr:`Cell.value` 内部调用，业务代码不应直接调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列。
        返回：公式结果的 Python 值；没有公式或没有有效结果时返回 ``None``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        if (row, column) not in self._formulas:
            return None
        return self._formula_values.get((row, column))

    def _set_cached_value(self, row: int, column: int, value: Any) -> None:
        """功能：登记公式单元格从 XLSX 读取到的外部缓存结果。

        使用方法：由 XLSX 读取器在设置公式后内部调用；工作簿计算器直接维护同一
        缓存容器，业务代码不应调用本方法。
        参数：``row``、``column`` 为 0-based 整数索引；``value`` 为 XML ``v``
        元素解析出的 Python 值，``None`` 表示没有缓存结果。
        返回：``None``；结果保存在公式专用内部容器中，不会覆盖公式。
        异常：索引无效或目标位置没有公式时抛出 ``InvalidAddressError`` 或
        ``ValueError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        if (row, column) not in self._formulas:
            raise ValueError("只有公式单元格可以设置缓存结果")
        if value is None:
            self._formula_values.pop((row, column), None)
        else:
            self._formula_values[(row, column)] = value
        self._formula_errors.pop((row, column), None)

    def _get_style(self, row: int, column: int) -> Style:
        """功能：读取指定位置的单元格样式。

        使用方法：由 :attr:`Cell.style` 读取器内部调用。
        参数：``row``、``column`` 为 0-based 整数索引，顺序为先行后列。
        返回：已设置的 :class:`Style`；没有自定义样式时返回 ``DEFAULT_STYLE``。
        异常：索引无效时抛出 ``InvalidAddressError``。
        """
        validate_row_index(row)
        validate_column_index(column)
        return self._styles.get((row, column), DEFAULT_STYLE)

    def _set_style(self, row: int, column: int, style: Style) -> None:
        """功能：设置或恢复指定位置的单元格样式。

        使用方法：由 ``cell.style = style`` 内部调用。
        参数：``row``、``column`` 为 0-based 整数索引；``style`` 必须是
        :class:`Style`。赋值 ``Style()`` 会移除显式样式记录。
        返回：``None``；即使单元格没有值，设置样式也会更新最大索引。
        异常：样式类型错误时抛出 ``TypeError``；索引无效时抛出
        ``InvalidAddressError``。
        """
        if not isinstance(style, Style):
            raise TypeError("cell.style 必须是 Style 对象")
        self._touch(row, column)
        key = (row, column)
        if style == DEFAULT_STYLE:
            self._styles.pop(key, None)
        else:
            self._styles[key] = style

    def __getitem__(self, address: str) -> Cell:
        """功能：支持 ``worksheet["A1"]`` 形式的单元格读取。

        使用方法：``cell = worksheet["A1"]``。
        参数：``address`` 为合法 A1 单元格地址字符串。
        返回：对应的 :class:`Cell`。
        异常：地址无效时抛出 ``InvalidAddressError``。
        """
        return self.cell(address)

    def __setitem__(self, address: str, value: Any) -> None:
        """功能：支持 ``worksheet["A1"] = value`` 形式的普通值赋值。

        使用方法：``worksheet["A1"] = "标题"``。
        参数：``address`` 为 A1 地址；``value`` 为普通值，``None`` 表示清除。
        返回：``None``；写入普通值会清除同一位置的公式。
        异常：地址无效时抛出 ``InvalidAddressError``。
        """
        self.cell(address).value = value

    def __repr__(self) -> str:
        """功能：生成用于调试的工作表文本表示。

        使用方法：``repr(worksheet)``。
        参数：无。
        返回：包含工作表名称的字符串。
        """
        return f"<Worksheet {self._name!r}>"
