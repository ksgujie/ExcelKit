"""Range：连续矩形区域的批量操作。"""

from __future__ import annotations

from typing import Any

from excelkit.address import index_to_range, range_to_index
from excelkit.errors import ExcelKitError, InvalidAddressError


class Range:
    """矩形区域，由 Worksheet.range("A1:C10") 创建。"""

    def __init__(self, worksheet: "Worksheet", address: str) -> None:
        self._worksheet = worksheet
        self.min_row, self.min_column, self.max_row, self.max_column = range_to_index(address)

    @property
    def address(self) -> str:
        """规范化的 A1 区域地址。"""
        from excelkit.address import index_to_range

        return index_to_range(
            self.min_row, self.min_column, self.max_row, self.max_column
        )

    @property
    def rows(self) -> int:
        """区域行数。"""
        return self.max_row - self.min_row + 1

    @property
    def columns(self) -> int:
        """区域列数。"""
        return self.max_column - self.min_column + 1

    # ---------- 值读写 ----------

    @property
    def values(self) -> list[list[Any]]:
        """以二维列表读取区域内有效值；公式单元格显示当前结果，无结果时为 None。"""
        result: list[list[Any]] = []
        for r in range(self.min_row, self.max_row + 1):
            row_vals: list[Any] = []
            for c in range(self.min_column, self.max_column + 1):
                row_vals.append(self._worksheet._get_cell_value(r, c))
            result.append(row_vals)
        return result

    def set_values(self, values: list[list[Any]]) -> Range:
        """写入等形状二维数据并返回当前区域；普通值会覆盖原公式。"""
        if len(values) != self.rows:
            raise ExcelKitError(
                f"值行数 {len(values)} 与区域行数 {self.rows} 不匹配"
            )
        for r_idx, row_vals in enumerate(values):
            if len(row_vals) != self.columns:
                raise ExcelKitError(
                    f"第 {r_idx} 行列数 {len(row_vals)} 与区域列数 {self.columns} 不匹配"
                )
            for c_idx, val in enumerate(row_vals):
                self._worksheet._set_cell_value(
                    self.min_row + r_idx, self.min_column + c_idx, val, formula=None
                )
        return self

    # ---------- 清除 ----------

    def clear(
        self,
        values: bool = True,
        styles: bool = True,
        hyperlinks: bool = False,
        notes: bool = False,
    ) -> Range:
        """按需清除区域内容和附属信息。

        参数:
            values: 是否清除值与公式。
            styles: 是否清除样式。
            hyperlinks: 是否清除超链接（精简版保留参数，暂不存储超链接）。
            notes: 是否清除批注（精简版保留参数，暂不存储批注）。
        """
        for r in range(self.min_row, self.max_row + 1):
            for c in range(self.min_column, self.max_column + 1):
                self._worksheet._clear_cell(r, c, values=values, styles=styles)
        return self

    # ---------- 复制 ----------

    def copy_to(
        self, target: Range, *, values: bool = True, formulas: bool = True, styles: bool = True
    ) -> Range:
        """复制到等尺寸区域，并按源目标偏移调整公式中的相对行列引用。

        精简版仅执行值/公式/样式的原样复制，不对公式引用做偏移调整。
        """
        if (self.rows, self.columns) != (target.rows, target.columns):
            raise ExcelKitError("源区域与目标区域尺寸不一致")
        for r_offset in range(self.rows):
            for c_offset in range(self.columns):
                src_r = self.min_row + r_offset
                src_c = self.min_column + c_offset
                tgt_r = target.min_row + r_offset
                tgt_c = target.min_column + c_offset
                if values or formulas:
                    val = self._worksheet._get_cell_value(src_r, src_c)
                    fmt = self._worksheet._get_cell_formula(src_r, src_c)
                    self._worksheet._set_cell_value(tgt_r, tgt_c, val, formula=fmt)
                if styles:
                    style = self._worksheet._get_cell_style(src_r, src_c)
                    self._worksheet._set_cell_style(tgt_r, tgt_c, style)
        return target

    # ---------- 合并 ----------

    def merge(self) -> Range:
        """合并当前连续区域。"""
        self._worksheet._merge_cells(self.min_row, self.min_column, self.max_row, self.max_column)
        return self

    def unmerge(self) -> Range:
        """取消合并当前区域。"""
        self._worksheet._unmerge_cells(self.min_row, self.min_column, self.max_row, self.max_column)
        return self

    # ---------- 记录转换 ----------

    def to_records(
        self, headers: bool = True, header_row: int | None = None
    ) -> list[dict[str, Any]]:
        """将区域转为字典列表。

        参数:
            headers: 是否将首行作为字段名。为 False 时自动生成 Column1、Column2 等。
            header_row: 字段所在的工作表绝对 0-based 行索引；为 None 时使用区域首行。
        """
        if header_row is None:
            field_row = self.min_row
        else:
            field_row = header_row

        if headers:
            field_values = self._worksheet._get_row_values(field_row, self.min_column, self.max_column)
            fields: list[str] = []
            for i, fv in enumerate(field_values):
                if fv is None or fv == "":
                    fields.append(f"Column{i + 1}")
                else:
                    fields.append(str(fv))
        else:
            fields = [f"Column{i + 1}" for i in range(self.columns)]

        data_start = field_row + 1 if headers else self.min_row
        records: list[dict[str, Any]] = []
        for r in range(data_start, self.max_row + 1):
            row_vals = self._worksheet._get_row_values(r, self.min_column, self.max_column)
            record: dict[str, Any] = {}
            for i, field in enumerate(fields):
                record[field] = row_vals[i] if i < len(row_vals) else None
            records.append(record)
        return records


__all__ = ["Range"]
