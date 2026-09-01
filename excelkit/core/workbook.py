"""工作簿对象及工作表集合管理。"""

from __future__ import annotations

import os
import re
from copy import deepcopy
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union

from ..errors import InvalidFileError, InvalidWorksheetNameError
from ..properties import WorkbookProperties
from ..protection import Protection
from .worksheet import Worksheet
from .named_range import NamedRange
from .range import Range

_FORBIDDEN_SHEET_NAME_CHARS = frozenset(":\\/?*[]")
_NAMED_RANGE_PATTERN = re.compile(r"^(?:[^\W\d]|_)[\w.]*$", re.UNICODE)
_TABLE_NAME_PATTERN = re.compile(r"^(?:[^\W\d]|_)[\w.]*$", re.UNICODE)


class Workbook:
    """表示可读取、编辑并写出 XLSX 或 XLS 文件的内存工作簿。"""

    __slots__ = ("_sheets", "_sheets_by_name", "_named_ranges", "_properties", "_protection")

    def __init__(self) -> None:
        """功能：创建不含工作表的空工作簿。

        使用方法：``workbook = Workbook()``。
        参数：无。
        返回：无；初始化后的实例可通过 :attr:`active` 延迟创建 ``Sheet1``。
        """
        self._sheets: list[Worksheet] = []
        self._sheets_by_name: Dict[str, Worksheet] = {}
        self._named_ranges: Dict[str, NamedRange] = {}
        self._properties = WorkbookProperties()
        self._protection = Protection()

    @staticmethod
    def _validate_sheet_name(name: str) -> None:
        """功能：验证名称是否符合 Excel 工作表命名规则。

        使用方法：由 :meth:`add_sheet` 内部调用。
        参数：``name`` 为待验证的字符串，长度必须为 1～31，且不能包含禁用字符。
        返回：验证成功时返回 ``None``。
        异常：名称类型、长度、字符或 XML 表示不合法时抛出
        :class:`InvalidWorksheetNameError`。
        """
        if not isinstance(name, str):
            raise InvalidWorksheetNameError("Worksheet name must be a string")
        if not 1 <= len(name) <= 31:
            raise InvalidWorksheetNameError("Worksheet name must contain 1 to 31 characters")
        invalid = sorted(set(name) & _FORBIDDEN_SHEET_NAME_CHARS)
        if invalid:
            raise InvalidWorksheetNameError(
                f"Worksheet name contains invalid character(s): {''.join(invalid)}"
            )
        if any(
            ord(char) < 32
            or 0xD800 <= ord(char) <= 0xDFFF
            or ord(char) in (0xFFFE, 0xFFFF)
            for char in name
        ):
            raise InvalidWorksheetNameError(
                "Worksheet name contains a character that cannot be stored in XML"
            )

    def add_sheet(self, name: str) -> Worksheet:
        """功能：按给定名称创建工作表并加入工作簿末尾。

        使用方法：``worksheet = workbook.add_sheet("成绩")``。
        参数：``name`` 为字符串形式的工作表名称。
        返回：新创建的 :class:`Worksheet` 对象。
        异常：名称无效时抛出 :class:`InvalidWorksheetNameError`；名称与现有工作表
        重复（不区分大小写）时抛出 ``ValueError``。
        """
        self._validate_sheet_name(name)
        normalized_name = name.casefold()
        if normalized_name in self._sheets_by_name:
            raise ValueError(f"Worksheet already exists: {name!r}")
        worksheet = Worksheet(self, name)
        self._sheets.append(worksheet)
        self._sheets_by_name[normalized_name] = worksheet
        return worksheet

    def _invalidate_formula_caches(self) -> None:
        """功能：在工作簿输入或公式变化后统一清除全部派生计算状态。

        使用方法：由工作表普通值和公式写入入口内部调用。
        参数：无。
        返回：``None``；所有公式保留，仅删除缓存结果和计算错误。
        """
        for worksheet in self._sheets:
            worksheet._formula_values.clear()
            worksheet._formula_errors.clear()

    def _shift_formula_references(
        self,
        target: Worksheet,
        index: int,
        count: int,
        *,
        rows: bool,
        deleting: bool,
    ) -> None:
        """功能：同步某张工作表行列编辑影响到工作簿内全部公式的引用。

        使用方法：由 ``Worksheet`` 行列编辑方法内部调用。
        参数：``target`` 为被编辑工作表；``index``、``count`` 为0-based位置和数量；
        ``rows`` 表示按行编辑；``deleting`` 表示删除而非插入。
        返回：``None``；公式表达式原地更新，缓存由调用方统一失效。
        """
        for worksheet in self._sheets:
            for coordinate, formula in list(worksheet._formulas.items()):
                worksheet._formulas[coordinate] = Worksheet._shift_formula_references(
                    formula,
                    index,
                    count,
                    rows=rows,
                    deleting=deleting,
                    current_sheet=target.name,
                    formula_sheet=worksheet.name,
                )

    @staticmethod
    def _validate_table_name(name: str) -> None:
        """功能：验证Excel数据表的工作簿级名称。

        使用方法：由 ``Worksheet.add_table()`` 内部调用。
        参数：``name`` 必须为1～255字符，以字母或下划线开头且不含空格。
        返回：验证成功时返回 ``None``。
        异常：名称无效时抛出 ``ValueError``。
        """
        if (
            not isinstance(name, str)
            or not 1 <= len(name) <= 255
            or _TABLE_NAME_PATTERN.fullmatch(name) is None
        ):
            raise ValueError("数据表名称必须以字母或下划线开头，且不能包含空格")
        from ..address import cell_index

        try:
            cell_index(name)
        except ValueError:
            return
        raise ValueError("数据表名称不能与A1单元格地址相同")

    def _table(self, name: str) -> object | None:
        """功能：在全部工作表中按大小写不敏感名称查找数据表。

        使用方法：由数据表创建和复制逻辑内部调用。
        参数：``name`` 为数据表名称。
        返回：匹配的 ``Table``；不存在时返回 ``None``。
        """
        key = name.casefold()
        for worksheet in self._sheets:
            if key in worksheet._tables:
                return worksheet._tables[key]
        return None

    def _unique_table_name(self, base: str) -> str:
        """功能：为复制工作表生成不与现有数据表冲突的新名称。

        使用方法：由 ``copy_sheet()`` 内部调用。
        参数：``base`` 为源数据表名称。
        返回：``base_Copy`` 或带递增数字后缀的唯一名称。
        """
        candidate = f"{base}_Copy"
        index = 2
        while self._table(candidate) is not None:
            candidate = f"{base}_Copy{index}"
            index += 1
        return candidate

    def _rename_sheet(self, worksheet: Worksheet, name: str) -> None:
        """功能：验证新名称并原子更新工作表名称索引。

        使用方法：仅由 ``worksheet.name = new_name`` 属性设置器调用。
        参数：``worksheet`` 为当前工作簿中的工作表；``name`` 为新名称字符串。
        返回：``None``；成功后名称查询立即使用新名称，工作表顺序保持不变。
        异常：工作表不属于当前工作簿时抛出 ``ValueError``；名称无效时抛出
        ``InvalidWorksheetNameError``；与其他工作表大小写不敏感重复时抛出
        ``ValueError``。
        """
        if worksheet not in self._sheets:
            raise ValueError("工作表不属于当前工作簿")
        self._validate_sheet_name(name)
        if worksheet.name == name:
            return
        normalized_name = name.casefold()
        existing = self._sheets_by_name.get(normalized_name)
        if existing is not None and existing is not worksheet:
            raise ValueError(f"Worksheet already exists: {name!r}")
        old_name = worksheet.name
        self._sheets_by_name.pop(old_name.casefold())
        worksheet._name = name
        self._sheets_by_name[normalized_name] = worksheet
        self._invalidate_formula_caches()

    def sheet(self, name: Union[str, int]) -> Worksheet:
        """功能：按名称或索引取得已有工作表。

        使用方法：``workbook.sheet("成绩")`` 按名称查询；``workbook.sheet(0)``
        按创建顺序取得第一张工作表。整数索引从 0 开始且不接受负数；字符串标签
        查询不区分大小写。
        参数：``name`` 可以是工作表名称字符串，也可以是整数索引；布尔值不作为索引。
        返回：匹配的 :class:`Worksheet` 对象。
        异常：字符串名称不存在时抛出 ``KeyError``；索引越界时抛出 ``IndexError``；
        参数不是字符串或整数时抛出 ``TypeError``。
        """
        if isinstance(name, str):
            try:
                return self._sheets_by_name[name.casefold()]
            except KeyError:
                raise KeyError(name) from None
        if isinstance(name, int) and not isinstance(name, bool):
            if name < 0:
                raise IndexError("工作表索引不能为负数")
            return self._sheets[name]
        raise TypeError("工作表标识必须是名称字符串或整数索引")

    def remove_sheet(self, name_or_index: Union[str, int]) -> "Workbook":
        """功能：按名称或0-based索引删除一张工作表。

        使用方法：``workbook.remove_sheet("临时表")`` 或
        ``workbook.remove_sheet(0)``。
        参数：``name_or_index`` 的规则与 :meth:`sheet` 完全一致。
        返回：当前 :class:`Workbook`，支持继续 ``save()`` 链式调用。
        异常：名称不存在、索引越界或参数类型无效时透传 ``sheet()`` 的异常。
        """
        worksheet = self.sheet(name_or_index)
        self._sheets.remove(worksheet)
        self._sheets_by_name.pop(worksheet.name.casefold())
        for key, named_range in list(self._named_ranges.items()):
            if named_range.worksheet is worksheet:
                self._named_ranges.pop(key)
        self._invalidate_formula_caches()
        return self

    @staticmethod
    def _validate_named_range_name(name: str) -> None:
        """功能：验证工作簿级命名区域名称是否清晰且符合常用Excel规则。

        使用方法：由 ``add_named_range()`` 内部调用。
        参数：``name`` 必须以字母或下划线开头，后续可含字母、数字、下划线和点，
        长度不超过255且不能看起来像A1单元格地址。
        返回：验证成功时返回 ``None``。
        异常：名称类型或格式无效时抛出 ``ValueError``。
        """
        if not isinstance(name, str) or not 1 <= len(name) <= 255:
            raise ValueError("命名区域名称必须是1～255个字符的字符串")
        if _NAMED_RANGE_PATTERN.fullmatch(name) is None:
            raise ValueError("命名区域名称必须以字母或下划线开头，且不能包含空格")
        from ..address import cell_index

        try:
            cell_index(name)
        except ValueError:
            return
        raise ValueError("命名区域名称不能与A1单元格地址相同")

    def add_named_range(self, name: str, area: Range) -> NamedRange:
        """功能：为当前工作簿中的一块区域登记唯一业务名称。

        使用方法：``workbook.add_named_range("SalesAmount", worksheet.range("E2:E100"))``。
        参数：``name`` 为大小写不敏感的唯一名称；``area`` 必须是当前工作簿中
        工作表创建的 :class:`Range`。
        返回：新创建的 :class:`NamedRange`。
        异常：名称无效或重复时抛出 ``ValueError``；区域类型或归属错误时抛出
        ``TypeError`` 或 ``ValueError``。
        """
        self._validate_named_range_name(name)
        if not isinstance(area, Range):
            raise TypeError("area 必须是 Range")
        if area.worksheet._workbook is not self:
            raise ValueError("命名区域必须属于当前工作簿")
        key = name.casefold()
        if key in self._named_ranges:
            raise ValueError(f"命名区域已经存在：{name!r}")
        named_range = NamedRange(self, name, area)
        self._named_ranges[key] = named_range
        return named_range

    def named_range(self, name: str) -> NamedRange:
        """功能：按大小写不敏感名称取得工作簿级命名区域。

        使用方法：``named = workbook.named_range("SalesAmount")``。
        参数：``name`` 为名称字符串。
        返回：匹配的 :class:`NamedRange`。
        异常：名称不存在时抛出 ``KeyError``；类型错误时抛出 ``TypeError``。
        """
        if not isinstance(name, str):
            raise TypeError("name 必须是字符串")
        try:
            return self._named_ranges[name.casefold()]
        except KeyError:
            raise KeyError(name) from None

    @property
    def named_ranges(self) -> Tuple[NamedRange, ...]:
        """功能：取得全部工作簿级命名区域的只读顺序快照。

        使用方法：``for item in workbook.named_ranges: ...``。
        参数：无，只读属性。
        返回：按创建或读取顺序排列的 ``tuple[NamedRange, ...]``。
        """
        return tuple(self._named_ranges.values())

    def remove_named_range(self, name: str) -> "Workbook":
        """功能：按名称删除一个工作簿级命名区域。

        使用方法：``workbook.remove_named_range("SalesAmount")``。
        参数：``name`` 为大小写不敏感名称字符串。
        返回：当前 :class:`Workbook`，支持链式调用。
        异常：名称不存在时抛出 ``KeyError``；类型错误时抛出 ``TypeError``。
        """
        named_range = self.named_range(name)
        self._named_ranges.pop(named_range.name.casefold())
        return self

    def move_sheet(self, name_or_index: Union[str, int], index: int) -> "Workbook":
        """功能：把一张工作表移动到指定0-based最终位置。

        使用方法：``workbook.move_sheet("统计", 0)``。
        参数：``name_or_index`` 标识待移动工作表；``index`` 为移动后位置，必须是
        当前工作表范围内的非负0-based整数，布尔值不作为索引。
        返回：当前 :class:`Workbook`，工作表对象及内容保持不变。
        异常：目标索引类型无效时抛出 ``TypeError``，越界时抛出 ``IndexError``；
        工作表标识错误时透传 :meth:`sheet` 的异常。
        """
        worksheet = self.sheet(name_or_index)
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError("目标工作表索引必须是0-based整数")
        if not 0 <= index < len(self._sheets):
            raise IndexError("目标工作表索引越界")
        self._sheets.remove(worksheet)
        self._sheets.insert(index, worksheet)
        return self

    def copy_sheet(
        self, name_or_index: Union[str, int], new_name: str
    ) -> Worksheet:
        """功能：完整复制工作表内容、布局和打印设置到工作簿末尾。

        使用方法：``copied = workbook.copy_sheet("模板", "八月报表")``。
        参数：``name_or_index`` 标识源工作表；``new_name`` 为必须显式提供的唯一
        工作表名称。
        返回：新创建的 :class:`Worksheet`。
        异常：源标识、名称或名称重复时透传 :meth:`sheet`、:meth:`add_sheet`
        的相应异常。
        """
        source = self.sheet(name_or_index)
        copied_state = {
            "values": deepcopy(source._values._values),
            "formulas": dict(source._formulas),
            "formula_values": deepcopy(source._formula_values),
            "formula_errors": dict(source._formula_errors),
            "hyperlinks": dict(source._hyperlinks),
            "notes": dict(source._notes),
            "styles": dict(source._styles),
            "merged_ranges": list(source._merged_ranges),
            "rows": deepcopy(source._rows),
            "columns": deepcopy(source._columns),
            "page": deepcopy(source._page),
            "headers": source._headers,
            "validations": deepcopy(source._validations),
            "conditionals": deepcopy(source._conditionals),
            "horizontal_page_breaks": set(source._horizontal_page_breaks),
        }
        target = self.add_sheet(new_name)
        target._color = source._color
        target._visibility = source._visibility
        target._values._values = copied_state["values"]
        target._formulas = copied_state["formulas"]
        target._formula_values = copied_state["formula_values"]
        target._formula_errors = copied_state["formula_errors"]
        target._hyperlinks = copied_state["hyperlinks"]
        target._notes = copied_state["notes"]
        target._headers = copied_state["headers"]
        target._validations = copied_state["validations"]
        target._conditionals = copied_state["conditionals"]
        target._filter_conditions = dict(source._filter_conditions)
        target._protection.enabled = source._protection.enabled
        target._protection.password = source._protection.password
        target._protection.select_locked = source._protection.select_locked
        target._protection.select_unlocked = source._protection.select_unlocked
        target._styles = copied_state["styles"]
        target._merged_ranges = copied_state["merged_ranges"]
        target._rows = copied_state["rows"]
        target._columns = copied_state["columns"]
        target._freeze = source._freeze
        target._filter_range = source._filter_range
        target._show_gridlines = source._show_gridlines
        target._page = copied_state["page"]
        target._horizontal_page_breaks = copied_state["horizontal_page_breaks"]
        target._max_row = source._max_row
        target._max_column = source._max_column
        for table in source.tables:
            copied_table = target.add_table(
                table.range.address,
                name=self._unique_table_name(table.name),
                style=table.style,
                has_header=table.has_header,
                show_row_stripes=table.show_row_stripes,
                show_column_stripes=table.show_column_stripes,
            )
            if table.show_totals:
                copied_table._data_max_row = table._data_max_row
                copied_table._bounds = table._bounds
                copied_table._show_totals = True
            copied_table._totals.update(table.totals)
        for image in source.images:
            copied_image = target.add_image(
                image.payload, anchor=image.anchor, name=image.filename
            )
            copied_image.width = image.width
            copied_image.height = image.height
            copied_image.offset_x = image.offset_x
            copied_image.offset_y = image.offset_y
            copied_image.alt_text = image.alt_text
        for chart in source.charts:
            copied_chart = target.add_chart(chart.type, anchor=chart.anchor)
            copied_chart.title = chart.title
            copied_chart.width = chart.width
            copied_chart.height = chart.height
            copied_chart.legend.position = chart.legend.position
            for series in chart.series:
                copied_chart.add_series(
                    values=series.values,
                    categories=series.categories,
                    name=series.name,
                )
        return target

    @property
    def sheets(self) -> Tuple[Worksheet, ...]:
        """功能：取得全部工作表的只读顺序快照。

        使用方法：``for worksheet in workbook.sheets: ...``。
        参数：无。
        返回：按创建顺序排列的 ``tuple[Worksheet, ...]``。
        """
        return tuple(self._sheets)

    @property
    def properties(self) -> WorkbookProperties:
        """功能：取得当前工作簿的核心文档属性对象。

        使用方法：``workbook.properties.title = "销售报表"``。
        参数：无，只读属性；返回对象的字段可以直接修改。
        返回：``WorkbookProperties``，同一工作簿每次访问返回同一对象。
        """
        return self._properties

    @property
    def protection(self) -> Protection:
        """功能：取得工作簿保护设置对象。

        使用方法：``workbook.protection.enabled = True``。
        参数：无，只读属性；返回对象的字段可以直接修改。
        返回：同一个可修改 :class:`Protection`；保存 XLSX 时写入工作簿保护定义。
        """
        return self._protection

    @property
    def active(self) -> Worksheet:
        """功能：取得第一张工作表，空工作簿会自动创建 ``Sheet1``。

        使用方法：``worksheet = workbook.active``。
        参数：无。
        返回：工作簿中的第一张 :class:`Worksheet`。
        """
        if not self._sheets:
            return self.add_sheet("Sheet1")
        return self._sheets[0]

    @classmethod
    def load(
        cls,
        filename: str | os.PathLike[str],
        *,
        encoding: str | None = None,
        delimiter: str | None = None,
        has_header: bool = False,
    ) -> "Workbook":
        """功能：从常用表格文件构造全新的工作簿。

        使用方法：``workbook = Workbook.load("input.xlsx")``；读取 CSV 时可写成
        ``Workbook.load("data.csv", encoding="gb18030", delimiter=";",
        has_header=True)``。
        参数：``filename`` 为字符串或 ``os.PathLike`` 路径；支持 XLS、XLSX、
        XLSM、XLTX、CSV 和 TSV。``encoding`` 指定 CSV/TSV 编码，省略时自动尝试
        UTF-8、UTF-8 BOM、GB18030；``delimiter`` 指定单字符分隔符，省略时按
        扩展名或内容检测；``has_header`` 为真时将首行同时记录到
        ``worksheet.headers``（首行单元格数据仍保留）。这三个参数对 XLS/XLSX
        传入非默认值会抛出 ``ValueError``。Open XML 宏内容会被忽略且不会执行。
        返回：包含已读取工作表、普通值、公式、日期和受支持样式的新
        :class:`Workbook`；对子类调用时返回该子类实例。
        异常：文件不存在时抛出 ``FileNotFoundError``；格式不受支持或结构损坏时
        抛出 ``InvalidFileError``。
        """
        from ..reader import _load_workbook

        return _load_workbook(
            cls, filename, encoding=encoding, delimiter=delimiter,
            has_header=has_header,
        )

    def save(
        self,
        filename: str | os.PathLike[str],
        *,
        encoding: str = "utf-8-sig",
        delimiter: str | None = None,
        formulas: bool = False,
        validate: bool = False,
    ) -> "Workbook":
        """功能：按文件扩展名将当前工作簿保存为 Excel、CSV 或 TSV 文件。

        使用方法：``workbook.save("成绩.xlsx")``；单工作表也可使用
        ``workbook.save("成绩.csv", formulas=True)``。
        参数：``filename`` 为字符串或实现 ``os.PathLike`` 的目标文件路径；
        ``.xls`` 使用 Excel 97–2003 格式，``.xlsx`` 使用 Open XML 格式；扩展名
        不区分大小写。``.csv``、``.tsv`` 仅可用于恰好一张工作表，``encoding``
        默认为 Excel 兼容的 UTF-8 BOM，``delimiter`` 可覆盖默认分隔符，``formulas``
        为真时导出公式文本。``validate`` 为真时写出前执行 :meth:`validate`，发现
        问题立即抛出 ``ValueError``。三个文本参数不适用于 Excel 文件。
        返回：当前 :class:`Workbook`，用于链式调用。
        异常：扩展名不受支持、多表导出文本或 Excel 文件使用文本参数时抛出
        ``InvalidFileError`` 或 ``ValueError``；路径不可写时透传文件系统异常。
        扩展名验证失败不会延迟创建 ``Sheet1``。
        """
        if not isinstance(filename, (str, os.PathLike)):
            raise TypeError("filename 必须是字符串或 PathLike 对象")
        if not isinstance(validate, bool):
            raise TypeError("validate 必须是 bool")
        suffix = Path(filename).suffix.lower()
        if suffix not in {".xlsx", ".xls", ".csv", ".tsv"}:
            raise InvalidFileError(
                f"不支持的工作簿保存格式：{suffix or '无扩展名'}"
            )
        if not self._sheets:
            self.active
        if validate:
            problems = self.validate()
            if problems:
                raise ValueError("工作簿校验失败：" + "；".join(problems))
        if suffix in {".csv", ".tsv"}:
            if len(self._sheets) != 1:
                raise InvalidFileError("CSV/TSV 导出要求工作簿恰好包含一张工作表")
            from ..writer.delimited import write_delimited

            write_delimited(
                self._sheets[0], filename, encoding=encoding,
                delimiter=delimiter, formulas=formulas,
            )
        else:
            if encoding != "utf-8-sig" or delimiter is not None or formulas:
                raise ValueError("encoding、delimiter、formulas 仅适用于 CSV/TSV 文件")
        if suffix == ".xls":
            from ..writer.xls import XlsWriter

            XlsWriter(self).write(filename)
        elif suffix == ".xlsx":
            from ..writer.xlsx import XlsxWriter

            XlsxWriter(self).write(filename)
        return self

    def validate(self) -> list[str]:
        """功能：检查工作簿及其全部工作表的结构完整性。

        使用方法：``problems = workbook.validate()``；保存时可使用
        ``workbook.save("output.xlsx", validate=True)`` 自动阻止问题文件写出。
        参数：无；方法只读，不会修复工作簿。
        返回：中文问题描述字符串列表；空列表表示当前结构通过检查。
        """
        problems: list[str] = []
        seen_names: set[str] = set()
        for worksheet in self._sheets:
            key = worksheet.name.casefold()
            if key in seen_names:
                problems.append(f"工作表名称重复：{worksheet.name!r}")
            seen_names.add(key)
            problems.extend(f"工作表 {worksheet.name!r}：{item}" for item in worksheet.validate())
        table_names: dict[str, str] = {}
        for worksheet in self._sheets:
            for table in worksheet.tables:
                key = table.name.casefold()
                previous = table_names.get(key)
                if previous is not None:
                    problems.append(f"数据表名称重复：{previous!r} 与 {table.name!r}")
                else:
                    table_names[key] = table.name
        for named_range in self._named_ranges.values():
            if named_range.worksheet not in self._sheets:
                problems.append(f"命名区域 {named_range.name!r} 所属工作表不存在")
                continue
            min_row, min_column, max_row, max_column = named_range._bounds
            if min_row < 0 or min_column < 0 or min_row > max_row or min_column > max_column:
                problems.append(f"命名区域 {named_range.name!r} 边界无效")
        return problems

    def render(
        self,
        data: Optional[Mapping[str, Any]] = None,
        *,
        sheet_data: Optional[
            Mapping[Union[str, int], Mapping[str, Any]]
        ] = None,
        strict: bool = False,
    ) -> "Workbook":
        """功能：使用公共及分工作表数据替换模板标签并展开循环行块。

        使用方法：``workbook.render(data)`` 使用一份公共数据渲染全部工作表；
        ``workbook.render(data, sheet_data={"明细": local})`` 只渲染指定工作表，
        并让工作表数据覆盖同名公共字段。
        参数：``data`` 为所有目标工作表共享的根映射，``None`` 等价于空映射；
        ``sheet_data`` 为以工作表名称或0-based索引为键、独立根映射为值的映射；
        ``strict`` 默认为 ``False``，缺失标签按空值处理，为 ``True`` 时立即报错。
        返回：当前 :class:`Workbook`，支持链式调用。
        异常：参数、工作表标识或独立数据无效时抛出 ``TypeError``、``KeyError``、
        ``IndexError`` 或 ``ValueError``；严格模式缺失数据、循环结构或表达式错误时
        抛出 ``TemplateError``。任一目标失败时所有工作表保持渲染前状态。
        """
        from ..template import render_workbook

        return render_workbook(
            self, data, sheet_data=sheet_data, strict=strict
        )

    def calculate(self, *, strict: bool = False) -> "Workbook":
        """功能：在 Python 中计算当前版本支持的全部工作簿公式。

        使用方法：``workbook.calculate()``；严格模式使用
        ``workbook.calculate(strict=True)``。
        参数：``strict`` 为布尔值；``False`` 时逐格记录错误并继续，``True`` 时
        首个公式错误抛出 ``FormulaCalculationError``。
        返回：当前 :class:`Workbook`，支持继续 ``save()``。
        异常：``strict`` 类型无效时抛出 ``TypeError``；严格模式计算失败时抛出
        ``FormulaCalculationError``。
        """
        if not isinstance(strict, bool):
            raise TypeError("strict 必须是布尔值")
        from .calculation import calculate_workbook

        calculate_workbook(self, strict=strict)
        return self

    def __len__(self) -> int:
        """功能：返回工作簿当前包含的工作表数量。

        使用方法：``count = len(workbook)``。
        参数：无。
        返回：非负整数；读取数量不会触发 ``Sheet1`` 的延迟创建。
        """
        return len(self._sheets)

    def __repr__(self) -> str:
        """功能：生成用于调试的工作簿文本表示。

        使用方法：``repr(workbook)``。
        参数：无。
        返回：包含工作表数量的字符串。
        """
        return f"<Workbook sheets={len(self._sheets)}>"
