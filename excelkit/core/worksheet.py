"""Worksheet：工作表核心操作。"""

from __future__ import annotations

from typing import Any

from excelkit.address import index_to_cell, cell_to_index, index_to_range, range_to_index
from excelkit.errors import InvalidAddressError, InvalidWorksheetNameError
from excelkit.style import Style
from excelkit.core.dimension import ColumnDimension, RowDimension
from excelkit.core.page import PageSettings


class Worksheet:
    """工作表，由 Workbook 创建，不应直接实例化。"""

    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "very_hidden"

    def __init__(self, workbook: "Workbook", name: str = "Sheet1") -> None:
        from excelkit.core.workbook import Workbook  # noqa: F811 - 避免循环导入，延迟导入

        self._workbook: Workbook = workbook
        self._name = name
        self._color: str | None = None
        # 单元格存储：(row, col) -> {"value": ..., "formula": ..., "style": Style|None}
        self._cells: dict[tuple[int, int], dict[str, Any]] = {}
        # 已触及边界
        self._max_row = -1
        self._max_column = -1
        # 行列尺寸
        self._row_dims: dict[int, RowDimension] = {}
        self._col_dims: dict[int, ColumnDimension] = {}
        # 合并区域
        self._merged: list[tuple[int, int, int, int]] = []
        # 冻结窗格
        self._freeze_panes: str | None = None
        # 自动筛选
        self._auto_filter_range: str | None = None
        # 页面设置
        self._page = PageSettings()
        # 图片
        self._images: list[Any] = []
        # 可见性
        self._visibility = self.VISIBLE

    # ---------- 属性 ----------

    @property
    def name(self) -> str:
        """工作表名称。"""
        return self._name

    @name.setter
    def name(self, value: str) -> None:
        if not value or not isinstance(value, str):
            raise InvalidWorksheetNameError("工作表名称不能为空")
        invalid_chars = set(":\\/?*[]")
        if any(ch in invalid_chars for ch in value):
            raise InvalidWorksheetNameError(f"工作表名称包含非法字符：{value!r}")
        old = self._name
        self._name = value
        self._workbook._sync_sheet_name(self, old, value)

    @property
    def tab_color(self) -> str | None:
        """工作表标签颜色（6 位 RGB 或 8 位 ARGB，不含 #）。"""
        return self._color

    @tab_color.setter
    def tab_color(self, value: str | None) -> None:
        if value is None:
            self._color = None
            return
        if not isinstance(value, str):
            raise ValueError(f"标签颜色必须是字符串或 None：{value!r}")
        text = value.strip().lstrip("#").upper()
        if len(text) not in (6, 8) or any(ch not in "0123456789ABCDEF" for ch in text):
            raise ValueError(
                f"标签颜色必须是 6 位 RGB 或 8 位 ARGB 十六进制字符串：{value!r}"
            )
        self._color = text

    @property
    def max_row(self) -> int:
        """已触及的最大 0-based 行索引；空表为 -1。"""
        return self._max_row

    @property
    def max_column(self) -> int:
        """已触及的最大 0-based 列索引；空表为 -1。"""
        return self._max_column

    @property
    def visibility(self) -> str:
        """可见状态：VISIBLE / HIDDEN / VERY_HIDDEN。"""
        return self._visibility

    @visibility.setter
    def visibility(self, value: str) -> None:
        if value not in (self.VISIBLE, self.HIDDEN, self.VERY_HIDDEN):
            raise InvalidWorksheetNameError(f"无效的可见状态：{value!r}")
        self._visibility = value

    @property
    def merged_ranges(self) -> tuple[tuple[int, int, int, int], ...]:
        """返回全部合并区域的只读 tuple，每项为 (min_row, min_col, max_row, max_col)。"""
        return tuple(self._merged)

    @property
    def freeze_panes(self) -> str | None:
        """冻结后的第一个可滚动 A1 单元格，None 清除。"""
        return self._freeze_panes

    @freeze_panes.setter
    def freeze_panes(self, value: str | None) -> None:
        if value is not None:
            cell_to_index(value)  # 校验地址
        self._freeze_panes = value

    @property
    def auto_filter(self) -> "_AutoFilterProxy":
        """自动筛选代理，支持 .range 读写、.set()/apply()、.clear()。"""
        return _AutoFilterProxy(self)

    @property
    def page(self) -> PageSettings:
        """页面布局与打印设置唯一入口。"""
        return self._page

    @property
    def images(self) -> tuple:
        """返回本工作表中的图片对象（只读副本）。"""
        return tuple(self._images)

    def add_image(
        self,
        anchor: str,
        source: Any,
        *,
        mode: str = "float",
        width: int | None = None,
        height: int | None = None,
        keep_aspect: bool = True,
        offset: tuple[int, int] = (0, 0),
        alt_text: str | None = None,
    ):
        """在工作表中插入图片并返回 Image 对象。

        参数:
            anchor: 锚定单元格或区域，如 "B2"、"B2:D6"。支持 `$A$1` 形式。
            source: 图片路径、bytes 或二进制 file-like 对象（支持 PNG/JPEG）。
            mode: "float" 浮动插入，固定像素尺寸，随单元格移动；
                  "embed" 完全嵌入锚定区域，随行高/列宽变化而缩放。
            width/height: float 模式下的显示尺寸（像素）；仅指定其一时按 keep_aspect 计算另一维。
            keep_aspect: float 模式下是否保持宽高比。
            offset: float 模式下相对锚点的 (x, y) 像素偏移。
            alt_text: 图片替代文字。
        """
        from excelkit.core.image import create_image

        img = create_image(
            anchor,
            source,
            mode=mode,
            width=width,
            height=height,
            keep_aspect=keep_aspect,
            offset=offset,
            alt_text=alt_text,
        )
        # embed 模式：若锚点为单个单元格且该单元格是某合并区域左上角，则自动扩展为整个合并区域
        if img.mode == "embed" and ":" not in img.anchor:
            try:
                r, c = cell_to_index(img.anchor)
                for mr1, mc1, mr2, mc2 in self._merged:
                    if mr1 == r and mc1 == c:
                        img.anchor = index_to_range(mr1, mc1, mr2, mc2)
                        break
            except Exception:
                pass
        self._images.append(img)
        return img

    def remove_image(self, image) -> "Worksheet":
        """移除指定图片对象。"""
        try:
            self._images.remove(image)
        except ValueError:
            pass
        return self

    def clear_images(self) -> "Worksheet":
        """清除本工作表中的全部图片。"""
        self._images.clear()
        return self

    # ---------- 单元格访问 ----------

    def __getitem__(self, address: str) -> "Cell":
        """A1 地址访问单元格，例如 ws["A1"]。"""
        return self.cell(address)

    def __setitem__(self, address: str, value: Any) -> None:
        """A1 地址赋值单元格，例如 ws["A1"] = "姓名"。"""
        cell = self.cell(address)
        cell.value = value

    def cell(self, address_or_row: str | int, column: int | None = None) -> "Cell":
        """显式单元格访问。

        ws.cell("A1")  或  ws.cell(0, 0)  0-based 先行后列。
        """
        from excelkit.core.cell import Cell

        return Cell(self, address_or_row, column)

    def range(self, address: str) -> "Range":
        """创建连续矩形区域，例如 ws.range("A1:C10")。"""
        from excelkit.core.range import Range

        return Range(self, address)

    # ---------- 内部读写 ----------

    def _ensure_cell(self, row: int, column: int) -> dict[str, Any]:
        key = (row, column)
        cell = self._cells.get(key)
        if cell is None:
            cell = {"value": None, "formula": None, "style": None}
            self._cells[key] = cell
        return cell

    def _get_cell_value(self, row: int, column: int) -> Any:
        cell = self._cells.get((row, column))
        if cell is None:
            return None
        formula = cell.get("formula")
        if formula is not None:
            # 公式单元格：返回缓存计算结果（若无缓存则为 None）
            return cell.get("value")
        return cell.get("value")

    def _get_cell_formula(self, row: int, column: int) -> str | None:
        cell = self._cells.get((row, column))
        if cell is None:
            return None
        return cell.get("formula")

    def _get_cell_style(self, row: int, column: int) -> Style | None:
        cell = self._cells.get((row, column))
        if cell is None:
            return None
        return cell.get("style")

    def _set_cell_value(
        self, row: int, column: int, value: Any, formula: str | None
    ) -> None:
        cell = self._ensure_cell(row, column)
        cell["value"] = value
        cell["formula"] = formula
        self._max_row = max(self._max_row, row)
        self._max_column = max(self._max_column, column)

    def _set_cell_style(self, row: int, column: int, style: Style | None) -> None:
        cell = self._ensure_cell(row, column)
        cell["style"] = style

    def _clear_cell(
        self, row: int, column: int, values: bool = True, styles: bool = True
    ) -> None:
        cell = self._cells.get((row, column))
        if cell is None:
            return
        if values:
            cell["value"] = None
            cell["formula"] = None
        if styles:
            cell["style"] = None

    def _get_row_values(
        self, row: int, min_column: int, max_column: int
    ) -> list[Any]:
        return [self._get_cell_value(row, c) for c in range(min_column, max_column + 1)]

    # ---------- 合并 ----------

    def _merge_cells(
        self, min_row: int, min_col: int, max_row: int, max_col: int
    ) -> None:
        self._merged.append((min_row, min_col, max_row, max_col))

    def _unmerge_cells(
        self, min_row: int, min_col: int, max_row: int, max_col: int
    ) -> None:
        self._merged = [
            m for m in self._merged if m != (min_row, min_col, max_row, max_col)
        ]

    # ---------- 行列尺寸 ----------

    def row(self, index: int) -> RowDimension:
        """获取或创建行尺寸对象，用于设置行高、隐藏。"""
        dim = self._row_dims.get(index)
        if dim is None:
            dim = RowDimension()
            self._row_dims[index] = dim
        return dim

    def column(self, index: int) -> ColumnDimension:
        """获取或创建列尺寸对象，用于设置列宽、隐藏。"""
        dim = self._col_dims.get(index)
        if dim is None:
            dim = ColumnDimension()
            self._col_dims[index] = dim
        return dim

    # ---------- 追加数据 ----------

    def append(self, values: list[Any]) -> "Worksheet":
        """追加一行；空表从索引 0（即 A1）开始。"""
        start_row = 0 if self._max_row == -1 else self._max_row + 1
        for c_idx, val in enumerate(values):
            self._set_cell_value(start_row, c_idx, val, formula=None)
        return self

    def append_rows(self, rows: list[list[Any]]) -> "Worksheet":
        """原子校验并连续追加二维数据。"""
        if not rows:
            return self
        start_row = 0 if self._max_row == -1 else self._max_row + 1
        # 先全量校验列数一致
        col_count = len(rows[0])
        for i, row_vals in enumerate(rows):
            if len(row_vals) != col_count:
                raise InvalidWorksheetNameError(
                    f"第 {i} 行列数 {len(row_vals)} 与首行列数 {col_count} 不一致"
                )
        # 再写入
        for r_idx, row_vals in enumerate(rows):
            for c_idx, val in enumerate(row_vals):
                self._set_cell_value(start_row + r_idx, c_idx, val, formula=None)
        return self

    # ---------- 值枚举 ----------

    @property
    def values(self) -> list[list[Any]]:
        """返回从 A1 到已触及边界的全部普通值二维列表。"""
        if self._max_row == -1 or self._max_column == -1:
            return []
        result: list[list[Any]] = []
        for r in range(0, self._max_row + 1):
            result.append(self._get_row_values(r, 0, self._max_column))
        return result

    @property
    def used_range(self) -> "Range | None":
        """返回实际使用对象覆盖的最小 Range，空表为 None。"""
        if self._max_row == -1 or self._max_column == -1:
            return None
        return self.range(index_to_range(0, 0, self._max_row, self._max_column))

    # ---------- 校验 ----------

    def validate(self) -> list[str]:
        """只读检查工作表结构并返回中文问题列表。"""
        problems: list[str] = []
        # 检查合并区域是否重叠
        for i, m1 in enumerate(self._merged):
            for m2 in self._merged[i + 1 :]:
                if (
                    m1[0] <= m2[2]
                    and m1[2] >= m2[0]
                    and m1[1] <= m2[3]
                    and m1[3] >= m2[1]
                ):
                    problems.append(
                        f"合并区域重叠：{index_to_range(*m1)} 与 {index_to_range(*m2)}"
                    )
        return problems

    # ---------- 排序（简化） ----------

    def sort(
        self,
        address: str,
        *,
        keys: list[tuple[int, bool]] | None = None,
        has_header: bool = False,
    ) -> "Worksheet":
        """按 SortKey 对连续区域排序（简化版：仅支持按单列升序/降序）。"""
        min_r, min_c, max_r, max_c = range_to_index(address)
        data_start = min_r + 1 if has_header else min_r
        if data_start > max_r:
            return self
        sort_keys = keys or [(0, False)]  # (列偏移, 是否降序)
        rows = [
            [self._get_cell_value(r, c) for c in range(min_c, max_c + 1)]
            for r in range(data_start, max_r + 1)
        ]

        def sort_key(row):
            result = []
            for col_offset, desc in sort_keys:
                val = row[col_offset]
                result.append((val is not None, val if not desc else _Neg(val)))
            return tuple(result)

        rows.sort(key=sort_key)
        for r_idx, row_vals in enumerate(rows):
            for c_idx, val in enumerate(row_vals):
                self._set_cell_value(data_start + r_idx, min_c + c_idx, val, formula=None)
        return self

    # ---------- 查找/替换（简化，保留常用） ----------

    def find(
        self,
        query: Any,
        *,
        whole: bool = False,
        case_sensitive: bool = False,
        formula: bool = False,
        limit: int | None = None,
    ) -> list["Cell"]:
        """在值中查找匹配单元格，返回 Cell 列表。"""
        results: list[Cell] = []
        count = 0
        for (r, c), cell in self._cells.items():
            val = cell.get("value") if not formula else cell.get("formula")
            if val is None:
                continue
            match = _match_value(val, query, whole, case_sensitive)
            if match:
                results.append(Cell(self, r, c))
                count += 1
                if limit is not None and count >= limit:
                    break
        return results

    def replace(
        self,
        query: Any,
        replacement: Any,
        *,
        whole: bool = False,
        case_sensitive: bool = False,
        limit: int | None = None,
    ) -> int:
        """在值中批量替换，返回替换数量。"""
        replaced = 0
        for (r, c), cell in self._cells.items():
            val = cell.get("value")
            if val is None:
                continue
            if not _match_value(val, query, whole, case_sensitive):
                continue
            cell["value"] = replacement
            replaced += 1
            if limit is not None and replaced >= limit:
                break
        return replaced

    # ---------- 导出单表（简化） ----------

    def export(self, filename: str, *, delimiter: str = ",") -> None:
        """将当前工作表导出为 CSV/TSV 文件。"""
        import csv

        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=delimiter)
            for row_vals in self.values:
                writer.writerow(["" if v is None else v for v in row_vals])


def _match_value(val: Any, query: Any, whole: bool, case_sensitive: bool) -> bool:
    if isinstance(val, str) and isinstance(query, str):
        if case_sensitive:
            v, q = val, query
        else:
            v, q = val.lower(), query.lower()
        if whole:
            return v == q
        return q in v
    return val == query


class _Neg:
    """用于排序降序的包装器。"""

    __slots__ = ("val",)

    def __init__(self, val: Any) -> None:
        self.val = val

    def __lt__(self, other: "_Neg") -> bool:
        return self.val > other.val

    def __eq__(self, other: object) -> bool:
        return isinstance(other, _Neg) and self.val == other.val


class _AutoFilterProxy:
    """自动筛选代理。"""

    def __init__(self, worksheet: Worksheet) -> None:
        self._ws = worksheet

    @property
    def range(self) -> str | None:
        return self._ws._auto_filter_range

    @range.setter
    def range(self, value: str | None) -> None:
        if value is not None:
            range_to_index(value)  # 校验
        self._ws._auto_filter_range = value

    def set(self, address: str) -> "_AutoFilterProxy":
        """设置筛选条件（精简版仅记录区域）。"""
        self.range = address
        return self

    def apply(self) -> "_AutoFilterProxy":
        """应用筛选并隐藏不匹配行（精简版暂不隐藏行，仅记录已应用）。"""
        return self

    def clear(self) -> None:
        self._ws._auto_filter_range = None


__all__ = ["Worksheet"]
