"""Cell 与 CellValue：单元格读写与类型转换。"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from excelkit.address import cell_address, cell_index
from excelkit.errors import InvalidAddressError
from excelkit.style import Style


class CellValue:
    """单元格值的快照封装，提供类型转换方法，不修改原单元格。"""

    def __init__(self, value: Any) -> None:
        self._value = value

    @property
    def value(self) -> Any:
        """只读值。"""
        return self._value

    def as_string(self) -> str:
        """转为字符串。"""
        if self._value is None:
            return ""
        return str(self._value)

    def as_int(self) -> int:
        """转为整数；不截断小数（非整数浮点抛 ValueError），拒绝 NaN/Inf。"""
        if self._value is None:
            raise ValueError("无法将 None 转为 int")
        if isinstance(self._value, bool):
            return int(self._value)
        if isinstance(self._value, int):
            return self._value
        if isinstance(self._value, float):
            if self._value != self._value or self._value in (float("inf"), float("-inf")):
                raise ValueError(f"无法将 {self._value!r} 转为 int")
            if not self._value.is_integer():
                raise ValueError(f"无法将非整数小数转为 int：{self._value!r}")
            return int(self._value)
        if isinstance(self._value, str):
            text = self._value.strip()
            # 允许 "3.0" 这类整数值字符串
            try:
                if "." in text or "e" in text.lower() or "E" in text:
                    f = float(text)
                    if f != f or f in (float("inf"), float("-inf")) or not f.is_integer():
                        raise ValueError
                    return int(f)
                return int(text)
            except ValueError:
                raise ValueError(f"无法将 {self._value!r} 转为 int") from None
        raise ValueError(f"无法将 {self._value!r} 转为 int")

    def as_float(self) -> float:
        """转为浮点数，拒绝 NaN 和无穷大。"""
        if self._value is None:
            raise ValueError("无法将 None 转为 float")
        if isinstance(self._value, bool):
            return float(self._value)
        if isinstance(self._value, (int, float)):
            result = float(self._value)
        elif isinstance(self._value, str):
            result = float(self._value.strip())
        else:
            raise ValueError(f"无法将 {self._value!r} 转为 float")
        if result != result or result in (float("inf"), float("-inf")):
            raise ValueError(f"无法将 {self._value!r} 转为 float")
        return result

    def as_bool(self) -> bool:
        """转为布尔值，仅接受 true/false/yes/no/是/否/1/0 及等价布尔或数值。"""
        if isinstance(self._value, bool):
            return self._value
        if isinstance(self._value, (int, float)):
            return bool(self._value)
        if isinstance(self._value, str):
            text = self._value.strip().lower()
            if text in ("true", "yes", "是", "1"):
                return True
            if text in ("false", "no", "否", "0"):
                return False
        raise ValueError(f"无法将 {self._value!r} 转为 bool")

    def as_date(self) -> date:
        """转为 date。支持 # 前缀与常见 YYYY-M-D / YYYY/M/D 格式。"""
        if isinstance(self._value, date) and not isinstance(self._value, datetime):
            return self._value
        if isinstance(self._value, datetime):
            return self._value.date()
        if isinstance(self._value, str):
            text = self._value.strip()
            if text.startswith("#"):
                text = text[1:].strip()
            # 兼容非零填充
            for sep in ("-", "/"):
                parts = text.split(sep)
                if len(parts) == 3 and all(p.isdigit() for p in parts):
                    try:
                        return date(int(parts[0]), int(parts[1]), int(parts[2]))
                    except ValueError:
                        pass
            for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
                try:
                    return datetime.strptime(text, fmt).date()
                except ValueError:
                    continue
        raise ValueError(f"无法将 {self._value!r} 转为 date")

    def as_datetime(self) -> datetime:
        """转为 datetime。支持 # 前缀与常见日期/时间格式。"""
        if isinstance(self._value, datetime):
            return self._value
        if isinstance(self._value, date):
            return datetime(self._value.year, self._value.month, self._value.day)
        if isinstance(self._value, str):
            text = self._value.strip()
            if text.startswith("#"):
                text = text[1:].strip()
            # 兼容非零填充
            for sep in ("-", "/"):
                if " " in text or "T" in text:
                    dt_part, tm_part = text.replace("T", " ").split(" ", 1)
                else:
                    dt_part, tm_part = text, ""
                dparts = dt_part.split(sep)
                if len(dparts) == 3 and all(p.isdigit() for p in dparts):
                    try:
                        if tm_part:
                            tparts = tm_part.split(":")
                            if len(tparts) >= 2 and all(p.isdigit() for p in tparts[:3]):
                                ss = int(tparts[2]) if len(tparts) > 2 else 0
                                return datetime(int(dparts[0]), int(dparts[1]), int(dparts[2]), int(tparts[0]), int(tparts[1]), ss)
                        else:
                            return datetime(int(dparts[0]), int(dparts[1]), int(dparts[2]))
                    except ValueError:
                        pass
            for fmt in (
                "%Y-%m-%d %H:%M:%S",
                "%Y/%m/%d %H:%M:%S",
                "%Y-%m-%d %H:%M",
                "%Y/%m/%d %H:%M",
                "%Y-%m-%d",
                "%Y/%m/%d",
            ):
                try:
                    return datetime.strptime(text, fmt)
                except ValueError:
                    continue
        raise ValueError(f"无法将 {self._value!r} 转为 datetime")

    def read(self) -> "CellValue":
        """返回自身（为与 Cell.read() 接口一致，保留此方法）。"""
        return self


class Cell:
    """单元格对象，通过 Worksheet["A1"] 或 Worksheet.cell(row, column) 获取。"""

    def __init__(self, worksheet: "Worksheet", address_or_row: str | int, column: int | None = None) -> None:
        self._worksheet = worksheet
        if isinstance(address_or_row, str):
            self._row, self._column = cell_index(address_or_row)
        else:
            if column is None:
                raise InvalidAddressError("cell() 需要行列两个整数参数，或一个 A1 地址字符串")
            self._row = int(address_or_row)
            self._column = int(column)
        if self._row < 0 or self._column < 0:
            raise InvalidAddressError(f"行列索引不能为负：({self._row}, {self._column})")

    @property
    def row(self) -> int:
        """0-based 行索引。"""
        return self._row

    @property
    def column(self) -> int:
        """0-based 列索引。"""
        return self._column

    @property
    def index(self) -> tuple[int, int]:
        """0-based (row, column) 元组，先行后列。"""
        return (self._row, self._column)

    @property
    def address(self) -> str:
        """规范化的 A1 地址。"""
        return cell_address(self._row, self._column)

    # ---------- 值与公式 ----------

    @property
    def value(self) -> Any:
        """普通单元格返回已保存值，公式单元格返回当前有效计算结果（无结果时为 None）。"""
        return self._worksheet._get_cell_value(self._row, self._column)

    @value.setter
    def value(self, value: Any) -> None:
        self._worksheet._set_cell_value(self._row, self._column, value, formula=None)

    @property
    def formula(self) -> str | None:
        """公式字符串，读取时始终带 '='；赋值 None 清除公式。"""
        return self._worksheet._get_cell_formula(self._row, self._column)

    @formula.setter
    def formula(self, formula: str | None) -> None:
        self._worksheet._set_cell_value(self._row, self._column, value=None, formula=formula)

    # ---------- 样式 ----------

    @property
    def style(self) -> Style | None:
        """完整不可变样式。"""
        return self._worksheet._get_cell_style(self._row, self._column)

    @style.setter
    def style(self, style: Style | None) -> None:
        self._worksheet._set_cell_style(self._row, self._column, style)

    # ---------- 转换 ----------

    def copy_style(self, source: "Cell") -> None:
        """从另一个 Cell 复制完整样式，不复制值和公式。"""
        self.style = source.style

    def set_value(self, value: Any) -> "Cell":
        """写入值并返回当前 Cell，用于链式类型转换。"""
        self.value = value
        return self

    def read(self) -> CellValue:
        """取得只读值快照，不修改单元格。"""
        return CellValue(self.value)

    def as_string(self) -> str:
        """转换为字符串，写回单元格并返回。"""
        result = self.read().as_string()
        self.value = result
        return result

    def as_int(self) -> int:
        """转换为整数，写回单元格并返回；非整数小数会抛 ValueError。"""
        result = self.read().as_int()
        self.value = result
        return result

    def as_float(self) -> float:
        """转换为浮点数，写回单元格并返回。"""
        result = self.read().as_float()
        self.value = result
        return result

    def as_bool(self) -> bool:
        """转换为布尔值，写回单元格并返回。"""
        result = self.read().as_bool()
        self.value = result
        return result

    def as_date(self) -> date:
        """转换为日期，写回单元格并返回。"""
        result = self.read().as_date()
        self.value = result
        return result

    def as_datetime(self) -> datetime:
        """转换为日期时间，写回单元格并返回。"""
        result = self.read().as_datetime()
        self.value = result
        return result


__all__ = ["Cell", "CellValue"]
