"""Python 端公式计算器（精简版）。

支持基础算术、比较、括号、单格与区域引用、跨表引用，以及常用函数：
SUM、AVERAGE、MIN、MAX、COUNT、COUNTA、IF、AND、OR、NOT、ABS、INT、ROUND、
CONCAT、LEN、LEFT、RIGHT、MID。

这不是完整 Excel 公式引擎，未支持的函数应交给 Excel/WPS 计算。
"""

from __future__ import annotations

import math
from typing import Any

from excelkit.errors import FormulaCalculationError


class _FormulaParser:
    """递归下降公式解析器。"""

    def __init__(self, workbook, formula: str, default_sheet=None) -> None:
        self.workbook = workbook
        self.default_sheet = default_sheet
        self.formula = formula.lstrip("=")
        self.pos = 0
        self.n = len(self.formula)

    def parse(self) -> Any:
        result = self._parse_expression()
        self._skip_ws()
        if self.pos != self.n:
            raise FormulaCalculationError(
                f"公式解析失败，未处理的尾部：{self.formula[self.pos:]!r}"
            )
        return result

    # ---------- 词法 ----------

    def _skip_ws(self) -> None:
        while self.pos < self.n and self.formula[self.pos].isspace():
            self.pos += 1

    def _peek(self) -> str:
        return self.formula[self.pos] if self.pos < self.n else ""

    def _next(self) -> str:
        if self.pos >= self.n:
            raise FormulaCalculationError(f"公式意外结束，位置 {self.pos}")
        ch = self.formula[self.pos]
        self.pos += 1
        return ch

    def _parse_number(self) -> float:
        start = self.pos
        while self.pos < self.n and (
            self.formula[self.pos].isdigit() or self.formula[self.pos] == "."
        ):
            self.pos += 1
        text = self.formula[start:self.pos]
        try:
            return float(text)
        except ValueError:
            raise FormulaCalculationError(f"无效的数字：{text!r}")

    def _parse_string(self) -> str:
        quote = self._next()
        start = self.pos
        while self.pos < self.n and self.formula[self.pos] != quote:
            if self.formula[self.pos] == "\\":
                self.pos += 2
            else:
                self.pos += 1
        if self.pos >= self.n:
            raise FormulaCalculationError("字符串未闭合")
        value = self.formula[start:self.pos]
        self.pos += 1
        return value

    def _parse_identifier(self) -> str:
        start = self.pos
        while self.pos < self.n and (
            self.formula[self.pos].isalnum() or self.formula[self.pos] in "._"
        ):
            self.pos += 1
        return self.formula[start:self.pos]

    # ---------- 递归下降 ----------

    def _parse_expression(self) -> Any:
        left = self._parse_comparison()
        while True:
            self._skip_ws()
            if self.pos < self.n and self.formula[self.pos] in "+-":
                op = self._next()
                right = self._parse_comparison()
                if op == "+":
                    left = _to_number(left) + _to_number(right)
                else:
                    left = _to_number(left) - _to_number(right)
            else:
                break
        return left

    def _parse_comparison(self) -> Any:
        left = self._parse_term()
        self._skip_ws()
        ch = self._peek()
        if ch in ("=", "<", ">"):
            op = self._next()
            if op == "<" and self._peek() == ">":
                self._next()
                op = "<>"
            elif op == "<" and self._peek() == "=":
                self._next()
                op = "<="
            elif op == ">" and self._peek() == "=":
                self._next()
                op = ">="
            right = self._parse_term()
            if op == "=":
                left = _eq(left, right)
            elif op == "<>":
                left = not _eq(left, right)
            elif op == "<":
                left = _cmp_lt(left, right)
            elif op == "<=":
                left = _cmp_le(left, right)
            elif op == ">":
                left = _cmp_gt(left, right)
            elif op == ">=":
                left = _cmp_ge(left, right)
        return left

    def _parse_term(self) -> Any:
        left = self._parse_factor()
        while True:
            self._skip_ws()
            if self.pos < self.n and self.formula[self.pos] in "*/%":
                op = self._next()
                right = self._parse_factor()
                if op == "*":
                    left = _to_number(left) * _to_number(right)
                elif op == "/":
                    denominator = _to_number(right)
                    if denominator == 0:
                        raise FormulaCalculationError("除数不能为零")
                    left = _to_number(left) / denominator
                else:
                    left = _to_number(left) % _to_number(right)
            else:
                break
        return left

    def _parse_factor(self) -> Any:
        self._skip_ws()
        ch = self._peek()
        if ch == "(":
            self._next()
            val = self._parse_expression()
            self._skip_ws()
            if self.pos >= self.n or self.formula[self.pos] != ")":
                raise FormulaCalculationError("括号不匹配")
            self._next()
            return val
        if ch == "-":
            self._next()
            return -self._parse_factor()
        if ch == '"':
            return self._parse_string()
        if ch.isdigit():
            return self._parse_number()
        ident = self._parse_identifier()
        self._skip_ws()
        # 处理区域引用：A1:B2
        if self.pos < self.n and self.formula[self.pos] == ":":
            self._next()
            self._skip_ws()
            ident2 = self._parse_identifier()
            ident = f"{ident}:{ident2}"
        # 处理跨表引用：Sheet1!A1 或 Sheet1!A1:B2
        if self.pos < self.n and self.formula[self.pos] == "!":
            self._next()
            self._skip_ws()
            sheet_ref = self._parse_identifier()
            if self.pos < self.n and self.formula[self.pos] == ":":
                self._next()
                self._skip_ws()
                sheet_ref = f"{sheet_ref}:{self._parse_identifier()}"
            ident = f"{ident}!{sheet_ref}"
        if self.pos < self.n and self.formula[self.pos] == "(":
            return self._call_function(ident)
        return self._resolve_reference(ident)

    def _call_function(self, name: str) -> Any:
        self._next()  # 消耗 '('
        args: list[Any] = []
        while True:
            self._skip_ws()
            if self.pos < self.n and self.formula[self.pos] == ")":
                self._next()
                break
            args.append(self._parse_expression())
            self._skip_ws()
            if self.pos < self.n and self.formula[self.pos] == ",":
                self._next()
                continue
            if self.pos < self.n and self.formula[self.pos] == ")":
                self._next()
                break
            raise FormulaCalculationError(f"函数 {name} 参数分隔符错误")
        return _apply_function(name, args)

    def _resolve_reference(self, ident: str) -> Any:
        if "!" in ident:
            sheet_name, ref = ident.split("!", 1)
            sheet = self.workbook.sheet(sheet_name)
        else:
            sheet = self.default_sheet or self.workbook.active
            ref = ident
        if ":" in ref:
            rng = sheet.range(ref)
            return rng.values
        row, col = cell_to_index(ref)
        return sheet._get_cell_value(row, col)


# ---------- 类型转换 ----------

def _to_number(val: Any) -> float:
    if isinstance(val, bool):
        return float(val)
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        return float(val)
    if val is None:
        return 0.0
    raise FormulaCalculationError(f"无法转为数字：{val!r}")


def _to_string(val: Any) -> str:
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if val is None:
        return ""
    return str(val)


def _truthy(val: Any) -> bool:
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val != 0
    if isinstance(val, str):
        return val != ""
    return val is not None


def _iterate_scalars(arg: Any):
    """把单元格值或区域（一维/二维列表）递归展开为标量迭代器。"""
    if isinstance(arg, list):
        for row in arg:
            yield from _iterate_scalars(row)
    else:
        yield arg


def _flatten_numbers(args: list[Any]) -> list[float]:
    result: list[float] = []
    for arg in args:
        for v in _iterate_scalars(arg):
            if isinstance(v, bool):
                result.append(float(v))
            elif isinstance(v, (int, float)):
                result.append(float(v))
            elif isinstance(v, str):
                try:
                    result.append(float(v))
                except ValueError:
                    pass
    return result


def _sum_values(arg: Any) -> float:
    total = 0.0
    for v in _iterate_scalars(arg):
        total += _to_number(v)
    return total


# ---------- 函数实现 ----------

def _eq(a: Any, b: Any) -> bool:
    if isinstance(a, str) or isinstance(b, str):
        return _to_string(a) == _to_string(b)
    try:
        return _to_number(a) == _to_number(b)
    except Exception:
        return _to_string(a) == _to_string(b)


def _cmp_lt(a: Any, b: Any) -> bool:
    return _cmp_key(a) < _cmp_key(b)


def _cmp_le(a: Any, b: Any) -> bool:
    return _cmp_key(a) <= _cmp_key(b)


def _cmp_gt(a: Any, b: Any) -> bool:
    return _cmp_key(a) > _cmp_key(b)


def _cmp_ge(a: Any, b: Any) -> bool:
    return _cmp_key(a) >= _cmp_key(b)


def _cmp_key(val: Any):
    if isinstance(val, bool):
        return (1, int(val))
    if isinstance(val, (int, float)):
        return (2, float(val))
    if isinstance(val, str):
        return (3, _to_string(val))
    if val is None:
        return (0, "")
    return (4, str(val))


def _apply_function(name: str, args: list[Any]) -> Any:
    upper = name.upper()
    if upper == "SUM":
        total = 0.0
        for arg in args:
            total += _sum_values(arg)
        return total
    if upper == "AVERAGE":
        nums = _flatten_numbers(args)
        if not nums:
            raise FormulaCalculationError("AVERAGE 参数不能为空")
        return sum(nums) / len(nums)
    if upper == "MIN":
        nums = _flatten_numbers(args)
        if not nums:
            raise FormulaCalculationError("MIN 参数不能为空")
        return min(nums)
    if upper == "MAX":
        nums = _flatten_numbers(args)
        if not nums:
            raise FormulaCalculationError("MAX 参数不能为空")
        return max(nums)
    if upper == "COUNT":
        count = 0
        for arg in args:
            for v in _iterate_scalars(arg):
                if isinstance(v, bool):
                    continue
                if isinstance(v, (int, float)):
                    count += 1
                elif isinstance(v, str):
                    try:
                        float(v)
                        count += 1
                    except ValueError:
                        pass
        return count
    if upper == "COUNTA":
        count = 0
        for arg in args:
            for v in _iterate_scalars(arg):
                if v is not None and v != "":
                    count += 1
        return count
    if upper == "IF":
        if len(args) < 2 or len(args) > 3:
            raise FormulaCalculationError("IF 参数个数应为 2 或 3")
        return args[1] if _truthy(args[0]) else (args[2] if len(args) == 3 else None)
    if upper == "AND":
        return all(_truthy(a) for a in args)
    if upper == "OR":
        return any(_truthy(a) for a in args)
    if upper == "NOT":
        if len(args) != 1:
            raise FormulaCalculationError("NOT 参数个数应为 1")
        return not _truthy(args[0])
    if upper == "ABS":
        if len(args) != 1:
            raise FormulaCalculationError("ABS 参数个数应为 1")
        return abs(_to_number(args[0]))
    if upper == "INT":
        if len(args) != 1:
            raise FormulaCalculationError("INT 参数个数应为 1")
        return int(math.floor(_to_number(args[0])))
    if upper == "ROUND":
        if len(args) not in (1, 2):
            raise FormulaCalculationError("ROUND 参数个数应为 1 或 2")
        ndigits = args[1] if len(args) == 2 else 0
        return round(_to_number(args[0]), int(ndigits))
    if upper == "CONCAT":
        return "".join(_to_string(a) for a in args)
    if upper == "LEN":
        if len(args) != 1:
            raise FormulaCalculationError("LEN 参数个数应为 1")
        return len(_to_string(args[0]))
    if upper == "LEFT":
        if len(args) not in (1, 2):
            raise FormulaCalculationError("LEFT 参数个数应为 1 或 2")
        n = int(args[1]) if len(args) == 2 else 1
        return _to_string(args[0])[:n]
    if upper == "RIGHT":
        if len(args) not in (1, 2):
            raise FormulaCalculationError("RIGHT 参数个数应为 1 或 2")
        n = int(args[1]) if len(args) == 2 else 1
        s = _to_string(args[0])
        return s[-n:] if n > 0 else ""
    if upper == "MID":
        if len(args) != 3:
            raise FormulaCalculationError("MID 参数个数应为 3")
        s = _to_string(args[0])
        start = int(args[1]) - 1
        length = int(args[2])
        return s[start:start + length]
    raise FormulaCalculationError(f"不支持的函数：{name}")


# ---------- 地址解析（延迟导入避免循环） ----------

def cell_to_index(address: str) -> tuple[int, int]:
    from excelkit.address import cell_to_index as _ci

    return _ci(address)


# ---------- 工作簿计算入口 ----------

def calculate_workbook(workbook, strict: bool = False) -> None:
    """计算工作簿内全部公式单元格，并把结果缓存到 cell['value']。"""
    for ws in workbook.sheets:
        for (r, c), cell in list(ws._cells.items()):
            formula = cell.get("formula")
            if formula is None:
                continue
            try:
                parser = _FormulaParser(workbook, formula, default_sheet=ws)
                value = parser.parse()
                cell["value"] = value
            except Exception as exc:
                if strict:
                    raise FormulaCalculationError(
                        f"工作表 [{ws.name}] 单元格 {ws.name}!{r + 1}{c + 1} 计算失败：{exc}"
                    ) from exc
                cell["value"] = None


__all__ = ["calculate_workbook", "FormulaCalculationError"]
