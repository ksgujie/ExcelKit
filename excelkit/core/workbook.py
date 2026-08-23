"""工作簿对象及工作表集合管理。"""

from __future__ import annotations

import os
from copy import deepcopy
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union

from ..errors import InvalidFileError, InvalidWorksheetNameError
from .worksheet import Worksheet

_FORBIDDEN_SHEET_NAME_CHARS = frozenset(":\\/?*[]")


class Workbook:
    """表示一个可写出为 XLSX 文件的内存工作簿。"""

    __slots__ = ("_sheets", "_sheets_by_name")

    def __init__(self) -> None:
        """功能：创建不含工作表的空工作簿。

        使用方法：``workbook = Workbook()``。
        参数：无。
        返回：无；初始化后的实例可通过 :attr:`active` 延迟创建 ``Sheet1``。
        """
        self._sheets: list[Worksheet] = []
        self._sheets_by_name: Dict[str, Worksheet] = {}

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

    def _rename_sheet(self, worksheet: Worksheet, name: str) -> None:
        """功能：验证新名称并原子更新工作表名称索引。

        使用方法：仅由 ``worksheet.label = new_label`` 属性设置器调用。
        参数：``worksheet`` 为当前工作簿中的工作表；``name`` 为新名称字符串。
        返回：``None``；成功后名称查询立即使用新名称，工作表顺序保持不变。
        异常：工作表不属于当前工作簿时抛出 ``ValueError``；名称无效时抛出
        ``InvalidWorksheetNameError``；与其他工作表大小写不敏感重复时抛出
        ``ValueError``。
        """
        if worksheet not in self._sheets:
            raise ValueError("工作表不属于当前工作簿")
        self._validate_sheet_name(name)
        if worksheet.label == name:
            return
        normalized_name = name.casefold()
        existing = self._sheets_by_name.get(normalized_name)
        if existing is not None and existing is not worksheet:
            raise ValueError(f"Worksheet already exists: {name!r}")
        old_name = worksheet.label
        self._sheets_by_name.pop(old_name.casefold())
        worksheet._label = name
        self._sheets_by_name[normalized_name] = worksheet

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
        self._sheets_by_name.pop(worksheet.label.casefold())
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
            "styles": dict(source._styles),
            "merged_ranges": list(source._merged_ranges),
            "rows": deepcopy(source._rows),
            "columns": deepcopy(source._columns),
            "page": deepcopy(source._page),
        }
        target = self.add_sheet(new_name)
        target._label_color = source._label_color
        target._values._values = copied_state["values"]
        target._formulas = copied_state["formulas"]
        target._styles = copied_state["styles"]
        target._merged_ranges = copied_state["merged_ranges"]
        target._rows = copied_state["rows"]
        target._columns = copied_state["columns"]
        target._freeze = source._freeze
        target._filter_range = source._filter_range
        target._show_gridlines = source._show_gridlines
        target._page = copied_state["page"]
        target._max_row = source._max_row
        target._max_column = source._max_column
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
    def load(cls, filename: str | os.PathLike[str]) -> "Workbook":
        """功能：从常用表格文件构造全新的工作簿。

        使用方法：``workbook = Workbook.load("input.xlsx")``。
        参数：``filename`` 为字符串或 ``os.PathLike`` 路径；支持 XLS、XLSX、
        XLSM、XLTX、CSV 和 TSV。Open XML 宏内容会被忽略且不会执行。
        返回：包含已读取工作表、普通值、公式、日期和受支持样式的新
        :class:`Workbook`；对子类调用时返回该子类实例。
        异常：文件不存在时抛出 ``FileNotFoundError``；格式不受支持或结构损坏时
        抛出 ``InvalidFileError``。
        """
        from ..reader import _load_workbook

        return _load_workbook(cls, filename)

    def save(self, filename: str | os.PathLike[str]) -> "Workbook":
        """功能：按文件扩展名将当前工作簿保存为 XLSX 或 XLS 文件。

        使用方法：``workbook.save("成绩.xlsx")``，也可传入 ``pathlib.Path``。
        参数：``filename`` 为字符串或实现 ``os.PathLike`` 的目标文件路径；
        ``.xls`` 使用 Excel 97–2003 格式，``.xlsx`` 使用 Open XML 格式；扩展名
        不区分大小写，其他扩展名不会被写出。
        返回：当前 :class:`Workbook`，用于链式调用。
        异常：扩展名不是 ``.xlsx`` 或 ``.xls`` 时抛出 ``InvalidFileError``；路径
        不可写时透传文件系统异常。扩展名验证失败不会延迟创建 ``Sheet1``。
        """
        if not isinstance(filename, (str, os.PathLike)):
            raise TypeError("filename 必须是字符串或 PathLike 对象")
        suffix = Path(filename).suffix.lower()
        if suffix not in {".xlsx", ".xls"}:
            raise InvalidFileError(
                f"不支持的工作簿保存格式：{suffix or '无扩展名'}"
            )
        if not self._sheets:
            self.active
        if suffix == ".xls":
            from ..writer.xls import XlsWriter

            XlsWriter(self).write(filename)
        else:
            from ..writer.xlsx import XlsxWriter

            XlsxWriter(self).write(filename)
        return self

    def render(
        self,
        data: Optional[Mapping[str, Any]] = None,
        *,
        by_sheet: Optional[
            Mapping[Union[str, int], Mapping[str, Any]]
        ] = None,
        strict: bool = False,
    ) -> "Workbook":
        """功能：使用公共及分工作表数据替换模板标签并展开循环行块。

        使用方法：``workbook.render(data)`` 使用一份公共数据渲染全部工作表；
        ``workbook.render(data, by_sheet={"明细": local})`` 只渲染指定工作表，
        并让工作表数据覆盖同名公共字段。
        参数：``data`` 为所有目标工作表共享的根映射，``None`` 等价于空映射；
        ``by_sheet`` 为以工作表名称或0-based索引为键、独立根映射为值的映射；
        ``strict`` 默认为 ``False``，缺失标签按空值处理，为 ``True`` 时立即报错。
        返回：当前 :class:`Workbook`，支持链式调用。
        异常：参数、工作表标识或独立数据无效时抛出 ``TypeError``、``KeyError``、
        ``IndexError`` 或 ``ValueError``；严格模式缺失数据、循环结构或表达式错误时
        抛出 ``TemplateError``。任一目标失败时所有工作表保持渲染前状态。
        """
        from ..template import render_workbook

        return render_workbook(
            self, data, by_sheet=by_sheet, strict=strict
        )

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
