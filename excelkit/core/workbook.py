"""Workbook：工作簿核心类，是用户的主要入口。"""

from __future__ import annotations

from typing import Any

from excelkit.address import cell_address, cell_index
from excelkit.errors import InvalidAddressError, InvalidFileError, InvalidWorksheetNameError
from excelkit.core.worksheet import Worksheet


class Workbook:
    """工作簿，用户通过 Workbook() 创建，或 Workbook.load() 读取文件。"""

    def __init__(self) -> None:
        self._sheets: list[Worksheet] = []
        self._name_index: dict[str, int] = {}  # 小写名称 -> 索引
        self._active_index: int = 0
        self.add_sheet("Sheet1")

    # ---------- 属性 ----------

    @property
    def sheets(self) -> tuple[Worksheet, ...]:
        """按创建顺序返回工作表 tuple。"""
        return tuple(self._sheets)

    @property
    def active(self) -> Worksheet:
        """返回当前活动工作表；空工作簿会创建 Sheet1。"""
        if not self._sheets:
            self.add_sheet("Sheet1")
        return self._sheets[self._active_index]

    # ---------- 工作表管理 ----------

    def add_sheet(self, name: str = "Sheet1") -> Worksheet:
        """创建并返回工作表。"""
        if not name or not isinstance(name, str):
            raise InvalidWorksheetNameError("工作表名称不能为空")
        invalid_chars = set(":\\/?*[]")
        if any(ch in invalid_chars for ch in name):
            raise InvalidWorksheetNameError(f"工作表名称包含非法字符：{name!r}")
        key = name.lower()
        if key in self._name_index:
            raise InvalidWorksheetNameError(f"工作表名称已存在：{name!r}")
        ws = Worksheet(self, name)
        self._sheets.append(ws)
        self._name_index[key] = len(self._sheets) - 1
        self._active_index = len(self._sheets) - 1
        return ws

    def sheet(self, name_or_index: str | int) -> Worksheet:
        """按不区分大小写的标签或非负 0-based 索引返回工作表。"""
        if isinstance(name_or_index, bool) or not isinstance(name_or_index, int):
            key = str(name_or_index).lower()
            if key not in self._name_index:
                raise InvalidWorksheetNameError(f"工作表不存在：{name_or_index!r}")
            return self._sheets[self._name_index[key]]
        if name_or_index < 0 or name_or_index >= len(self._sheets):
            raise InvalidAddressError(f"工作表索引超出范围：{name_or_index}")
        return self._sheets[name_or_index]

    def remove_sheet(self, name_or_index: str | int) -> Workbook:
        """删除工作表并返回当前工作簿。"""
        ws = self.sheet(name_or_index)
        idx = self._sheets.index(ws)
        key = ws.name.lower()
        del self._name_index[key]
        del self._sheets[idx]
        if not self._sheets:
            self.add_sheet("Sheet1")
            self._active_index = 0
        else:
            self._active_index = min(idx, len(self._sheets) - 1)
            # 重建名称索引
            self._name_index = {s.name.lower(): i for i, s in enumerate(self._sheets)}
        return self

    def move_sheet(self, name_or_index: str | int, index: int) -> Workbook:
        """移动工作表到指定 0-based 最终位置。"""
        ws = self.sheet(name_or_index)
        cur = self._sheets.index(ws)
        if index < 0 or index > len(self._sheets) - 1:
            raise InvalidAddressError(f"目标位置超出范围：{index}")
        self._sheets.pop(cur)
        self._sheets.insert(index, ws)
        self._name_index = {s.name.lower(): i for i, s in enumerate(self._sheets)}
        self._active_index = index
        return self

    def copy_sheet(self, name_or_index: str | int, new_name: str) -> Workbook:
        """完整复制工作表（内容、布局、打印设置）。"""
        import copy

        src = self.sheet(name_or_index)
        new_ws = copy.deepcopy(src)
        new_ws._workbook = self
        # 先加入工作簿，再重命名（名称 setter 需 worksheet 已在 sheets 列表中）
        self._sheets.append(new_ws)
        new_ws.name = new_name
        self._active_index = len(self._sheets) - 1
        return self

    def _sync_sheet_name(self, worksheet: Worksheet, old: str, new: str) -> None:
        """工作表重命名时同步名称索引（由 Worksheet.name setter 调用）。"""
        key_old = old.lower()
        if key_old in self._name_index and self._sheets[self._name_index[key_old]] is worksheet:
            del self._name_index[key_old]
        self._name_index[new.lower()] = self._sheets.index(worksheet)

    # ---------- 读取 ----------

    @classmethod
    def load(cls, filename: str, *, encoding: str | None = None) -> Workbook:
        """读取 XLSX、XLS、CSV 或 TSV 文件。"""
        from excelkit.reader import load_workbook

        return load_workbook(filename, encoding=encoding)

    # ---------- 模板渲染 ----------

    def render(
        self,
        data: dict[str, Any] | None = None,
        *,
        sheet_data: dict[str | int, dict[str, Any]] | None = None,
        strict: bool = False,
    ) -> Workbook:
        """使用公共或分工作表数据替换模板标签并展开循环行块。"""
        from excelkit.template import render_workbook

        render_workbook(self, data or {}, sheet_data or {}, strict)
        return self

    # ---------- 公式计算 ----------

    def calculate(self, *, strict: bool = False) -> Workbook:
        """在 Python 中计算当前支持的公式并更新缓存结果。"""
        from excelkit.calculation import calculate_workbook

        calculate_workbook(self, strict)
        return self

    # ---------- 校验 ----------

    def validate(self) -> list[str]:
        """只读检查工作簿结构并返回中文问题列表。"""
        problems: list[str] = []
        seen_names: set[str] = set()
        for ws in self._sheets:
            for p in ws.validate():
                problems.append(f"[{ws.name}] {p}")
            if ws.name in seen_names:
                problems.append(f"重复的工作表名称：{ws.name}")
            seen_names.add(ws.name)
        return problems

    # ---------- 保存 ----------

    def save(
        self,
        filename: str,
        *,
        encoding: str | None = None,
        formulas: bool = False,
        validate: bool = False,
    ) -> None:
        """保存 XLSX、XLS；单工作表还可保存 CSV/TSV。"""
        if validate:
            problems = self.validate()
            if problems:
                from excelkit.errors import ValidationError

                raise ValidationError("保存前校验失败：\n" + "\n".join(f"  - {p}" for p in problems))
        from excelkit.writer import save_workbook

        save_workbook(self, filename, encoding=encoding, formulas=formulas)


__all__ = ["Workbook"]
