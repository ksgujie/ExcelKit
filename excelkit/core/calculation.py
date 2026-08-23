"""ExcelKit 受控公式解析、依赖计算和常用函数实现。"""

from __future__ import annotations

import ast
import math
import re
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from typing import TYPE_CHECKING, Any, Callable, Iterable

from ..address import cell_index, range_index
from ..errors import FormulaCalculationError

if TYPE_CHECKING:
    from .workbook import Workbook
    from .worksheet import Worksheet


_REFERENCE_PATTERN = re.compile(
    r"(?<![\w.])"
    r"(?:(?:'((?:[^']|'')+)'|([A-Za-z_\u0080-\uffff][\w.\u0080-\uffff]*))!)?"
    r"(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)(?::(\$?[A-Za-z]{1,3}\$?[1-9][0-9]*))?"
    r"(?![\w.])"
)
_STRING_PATTERN = re.compile(r'"(?:[^"]|"")*"')


def _flatten(values: Iterable[Any]) -> list[Any]:
    """功能：把公式函数收到的区域和嵌套序列展开为一维参数列表。

    使用方法：由 ``SUM``、``COUNT`` 等聚合函数内部调用。
    参数：``values`` 为函数位置参数序列，可包含列表或元组。
    返回：保持从左到右、从上到下顺序的一维值列表。
    """
    flattened: list[Any] = []
    for value in values:
        if isinstance(value, (list, tuple)):
            flattened.extend(_flatten(value))
        else:
            flattened.append(value)
    return flattened


def _numbers(values: Iterable[Any]) -> list[float | int]:
    """功能：提取聚合函数可以参与计算的非布尔有限数值。

    使用方法：由 ``SUM``、``AVERAGE``、``MIN`` 和 ``MAX`` 调用。
    参数：``values`` 可包含标量、区域二维列表、空值和文本。
    返回：只包含 ``int`` 或有限 ``float`` 的列表。
    """
    return [
        value
        for value in _flatten(values)
        if not isinstance(value, bool)
        and isinstance(value, (int, float))
        and math.isfinite(float(value))
    ]


def _sum(*values: Any) -> float | int:
    """功能：计算区域和标量中全部有效数值的总和。

    使用方法：公式 ``=SUM(A1:A10)``。
    参数：任意数量的标量或区域；文本、布尔值和空值忽略。
    返回：数值总和，没有数值时返回 ``0``。
    """
    return sum(_numbers(values))


def _average(*values: Any) -> float:
    """功能：计算区域和标量中有效数值的算术平均数。

    使用方法：公式 ``=AVERAGE(A1:A10)``。
    参数：任意数量的标量或区域。
    返回：浮点平均值；没有数值时抛出 ``FormulaCalculationError``。
    """
    numbers = _numbers(values)
    if not numbers:
        raise FormulaCalculationError("AVERAGE 没有可计算的数值")
    return sum(numbers) / len(numbers)


def _minimum(*values: Any) -> float | int:
    """功能：返回区域和标量中的最小有效数值。

    使用方法：公式 ``=MIN(A1:A10)``。
    参数：任意数量的标量或区域。
    返回：最小数值；没有数值时返回 ``0``。
    """
    numbers = _numbers(values)
    return min(numbers) if numbers else 0


def _maximum(*values: Any) -> float | int:
    """功能：返回区域和标量中的最大有效数值。

    使用方法：公式 ``=MAX(A1:A10)``。
    参数：任意数量的标量或区域。
    返回：最大数值；没有数值时返回 ``0``。
    """
    numbers = _numbers(values)
    return max(numbers) if numbers else 0


def _count(*values: Any) -> int:
    """功能：统计区域和标量中的有效数值数量。

    使用方法：公式 ``=COUNT(A1:A10)``。
    参数：任意数量的标量或区域。
    返回：不包含布尔值、文本和空值的数值数量。
    """
    return len(_numbers(values))


def _counta(*values: Any) -> int:
    """功能：统计区域和标量中的非空值数量。

    使用方法：公式 ``=COUNTA(A1:A10)``。
    参数：任意数量的标量或区域。
    返回：既不是 ``None`` 也不是空字符串的元素数量。
    """
    return sum(value not in (None, "") for value in _flatten(values))


def _and(*values: Any) -> bool:
    """功能：判断所有展开参数是否均为真。

    使用方法：公式 ``=AND(A1>0,B1>0)``。
    参数：任意数量的标量或区域。
    返回：全部值为真时返回 ``True``。
    """
    return all(bool(value) for value in _flatten(values))


def _or(*values: Any) -> bool:
    """功能：判断展开参数中是否至少有一个真值。

    使用方法：公式 ``=OR(A1>0,B1>0)``。
    参数：任意数量的标量或区域。
    返回：存在真值时返回 ``True``。
    """
    return any(bool(value) for value in _flatten(values))


def _concat(*values: Any) -> str:
    """功能：按顺序连接标量和区域中的全部值。

    使用方法：公式 ``=CONCAT(A1,"-",B1)``。
    参数：任意数量的标量或区域；空值转换为空字符串。
    返回：连接后的字符串。
    """
    return "".join("" if value is None else str(value) for value in _flatten(values))


def _left(value: Any, count: int = 1) -> str:
    """功能：取得文本左侧指定数量的字符。

    使用方法：公式 ``=LEFT(A1,3)``。
    参数：``value`` 为待转换文本；``count`` 为非负整数，默认1。
    返回：左侧文本；数量无效时抛出 ``FormulaCalculationError``。
    """
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise FormulaCalculationError("LEFT 的字符数必须是非负整数")
    return str(value)[:count]


def _right(value: Any, count: int = 1) -> str:
    """功能：取得文本右侧指定数量的字符。

    使用方法：公式 ``=RIGHT(A1,3)``。
    参数：``value`` 为待转换文本；``count`` 为非负整数，默认1。
    返回：右侧文本；数量无效时抛出 ``FormulaCalculationError``。
    """
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise FormulaCalculationError("RIGHT 的字符数必须是非负整数")
    return "" if count == 0 else str(value)[-count:]


def _mid(value: Any, start: int, count: int) -> str:
    """功能：按 Excel 的1-based起点取得文本中间片段。

    使用方法：公式 ``=MID(A1,2,3)``。
    参数：``value`` 为文本；``start`` 为正整数起点；``count`` 为非负整数长度。
    返回：截取后的字符串；索引无效时抛出 ``FormulaCalculationError``。
    """
    if (
        isinstance(start, bool)
        or not isinstance(start, int)
        or start <= 0
        or isinstance(count, bool)
        or not isinstance(count, int)
        or count < 0
    ):
        raise FormulaCalculationError("MID 的起点必须为正整数，长度必须为非负整数")
    return str(value)[start - 1:start - 1 + count]


def _round(value: Any, digits: int = 0) -> float | int:
    """功能：按 Excel 常用的四舍五入规则保留指定小数位。

    使用方法：公式 ``=ROUND(A1,2)``。
    参数：``value`` 为有限数值；``digits`` 为整数，可为负数。
    返回：四舍五入后的整数或浮点数。
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise FormulaCalculationError("ROUND 的值必须是数值")
    if isinstance(digits, bool) or not isinstance(digits, int):
        raise FormulaCalculationError("ROUND 的位数必须是整数")
    quantum = Decimal("1").scaleb(-digits)
    result = Decimal(str(value)).quantize(quantum, rounding=ROUND_HALF_UP)
    return int(result) if digits <= 0 else float(result)


_FUNCTIONS: dict[str, Callable[..., Any]] = {
    "SUM": _sum,
    "AVERAGE": _average,
    "MIN": _minimum,
    "MAX": _maximum,
    "COUNT": _count,
    "COUNTA": _counta,
    "AND": _and,
    "OR": _or,
    "NOT": lambda value: not bool(value),
    "ABS": abs,
    "INT": math.floor,
    "ROUND": _round,
    "CONCAT": _concat,
    "LEN": lambda value: len(str(value)),
    "LEFT": _left,
    "RIGHT": _right,
    "MID": _mid,
}


class _Evaluator:
    """在单个工作簿中解析引用、依赖关系和安全 AST 节点。"""

    def __init__(self, workbook: "Workbook") -> None:
        """功能：创建绑定到待计算工作簿的公式求值器。

        使用方法：由 :func:`calculate_workbook` 创建一次。
        参数：``workbook`` 为包含公式和数据的工作簿。
        返回：无；初始化依赖访问栈。
        """
        self.workbook = workbook
        self.visiting: set[tuple["Worksheet", int, int]] = set()

    def cell_value(self, worksheet: "Worksheet", row: int, column: int) -> Any:
        """功能：取得普通单元格值或递归计算公式单元格。

        使用方法：解析公式单元格引用时内部调用。
        参数：``worksheet`` 为所属表；``row``、``column`` 为0-based坐标。
        返回：普通值或公式计算结果。
        异常：循环引用或被引用公式计算失败时抛出 ``FormulaCalculationError``。
        """
        if (row, column) in worksheet._formulas:
            return self.calculate_cell(worksheet, row, column)
        return worksheet._values.get(row, column)

    def reference(self, worksheet: "Worksheet", reference: str) -> Any:
        """功能：把单格或矩形 A1 引用解析为当前计算值。

        使用方法：公式预处理替换引用时内部调用。
        参数：``worksheet`` 为引用目标表；``reference`` 为不带工作表名的A1文本。
        返回：单格值或按行组织的二维区域值。
        """
        clean = reference.replace("$", "")
        if ":" not in clean:
            row, column = cell_index(clean)
            return self.cell_value(worksheet, row, column)
        min_row, min_column, max_row, max_column = range_index(clean)
        return [
            [self.cell_value(worksheet, row, column) for column in range(min_column, max_column + 1)]
            for row in range(min_row, max_row + 1)
        ]

    def prepare(self, worksheet: "Worksheet", expression: str) -> tuple[str, dict[str, Any]]:
        """功能：把 Excel 引用替换为安全 AST 变量并规范化常用运算符。

        使用方法：每个公式解析 AST 前调用。
        参数：``worksheet`` 为默认引用表；``expression`` 为不含前导等号的公式。
        返回：Python 可解析表达式和变量值映射。
        异常：工作表或地址不存在时透传相应异常。
        """
        variables: dict[str, Any] = {}

        def replace_segment(segment: str) -> str:
            """功能：替换一个不在双引号字符串中的公式片段引用。

            使用方法：由公式字符串分段循环调用。
            参数：``segment`` 为不含 Excel 文本常量的字符串片段。
            返回：引用已替换为内部变量名的片段。
            """
            def replace(match: re.Match[str]) -> str:
                """功能：解析一个可能带工作表前缀的单格或区域引用。

                使用方法：由引用正则替换过程自动调用。
                参数：``match`` 为引用正则匹配结果。
                返回：保存对应当前值的内部变量名；无效列地址保持原文本。
                """
                quoted_sheet, plain_sheet, start, end = match.groups()
                reference = start if end is None else f"{start}:{end}"
                try:
                    cell_index(start.replace("$", ""))
                    if end is not None:
                        cell_index(end.replace("$", ""))
                except ValueError:
                    return match.group(0)
                target = worksheet
                sheet_name = quoted_sheet or plain_sheet
                if sheet_name is not None:
                    target = self.workbook.sheet(sheet_name.replace("''", "'"))
                variable = f"_ref_{len(variables)}"
                variables[variable] = self.reference(target, reference)
                return variable

            replaced = _REFERENCE_PATTERN.sub(replace, segment)
            replaced = replaced.replace("<>", "!=").replace("^", "**")
            return re.sub(r"(?<![<>=])=(?!=)", "==", replaced)

        parts: list[str] = []
        position = 0
        for match in _STRING_PATTERN.finditer(expression):
            parts.append(replace_segment(expression[position:match.start()]))
            parts.append(match.group(0))
            position = match.end()
        parts.append(replace_segment(expression[position:]))
        return "".join(parts), variables

    def node(self, node: ast.AST, variables: dict[str, Any]) -> Any:
        """功能：递归计算白名单中的安全 Python AST 节点。

        使用方法：由 :meth:`calculate_cell` 对解析后的表达式根节点调用。
        参数：``node`` 为 AST 节点；``variables`` 为单元格和区域引用值。
        返回：节点计算结果。
        异常：不支持语法、函数或运算失败时抛出 ``FormulaCalculationError``。
        """
        if isinstance(node, ast.Expression):
            return self.node(node.body, variables)
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name):
            upper = node.id.upper()
            if node.id in variables:
                return variables[node.id]
            if upper in {"TRUE", "FALSE"}:
                return upper == "TRUE"
            raise FormulaCalculationError(f"未知名称：{node.id}")
        if isinstance(node, ast.UnaryOp):
            value = self.node(node.operand, variables)
            operations = {ast.UAdd: lambda item: +item, ast.USub: lambda item: -item, ast.Not: lambda item: not item}
            operation = operations.get(type(node.op))
            if operation is None:
                raise FormulaCalculationError("不支持的单目运算")
            return operation(value)
        if isinstance(node, ast.BinOp):
            left = self.node(node.left, variables)
            right = self.node(node.right, variables)
            operations = {
                ast.Add: lambda a, b: a + b,
                ast.Sub: lambda a, b: a - b,
                ast.Mult: lambda a, b: a * b,
                ast.Div: lambda a, b: a / b,
                ast.FloorDiv: lambda a, b: a // b,
                ast.Mod: lambda a, b: a % b,
                ast.Pow: lambda a, b: a ** b,
            }
            operation = operations.get(type(node.op))
            if operation is None:
                raise FormulaCalculationError("不支持的二元运算")
            return operation(left, right)
        if isinstance(node, ast.Compare):
            left = self.node(node.left, variables)
            comparisons = {
                ast.Eq: lambda a, b: a == b,
                ast.NotEq: lambda a, b: a != b,
                ast.Lt: lambda a, b: a < b,
                ast.LtE: lambda a, b: a <= b,
                ast.Gt: lambda a, b: a > b,
                ast.GtE: lambda a, b: a >= b,
            }
            for operator, comparator in zip(node.ops, node.comparators):
                right = self.node(comparator, variables)
                comparison = comparisons.get(type(operator))
                if comparison is None or not comparison(left, right):
                    return False
                left = right
            return True
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            name = node.func.id.upper()
            if node.keywords:
                raise FormulaCalculationError("公式函数不支持关键字参数")
            if name == "IF":
                if len(node.args) not in {2, 3}:
                    raise FormulaCalculationError("IF 需要2个或3个参数")
                condition = bool(self.node(node.args[0], variables))
                selected = node.args[1] if condition else (node.args[2] if len(node.args) == 3 else False)
                return self.node(selected, variables) if isinstance(selected, ast.AST) else selected
            function = _FUNCTIONS.get(name)
            if function is None:
                raise FormulaCalculationError(f"不支持的公式函数：{name}")
            return function(*(self.node(argument, variables) for argument in node.args))
        raise FormulaCalculationError(f"不支持的公式语法：{type(node).__name__}")

    def calculate_cell(self, worksheet: "Worksheet", row: int, column: int) -> Any:
        """功能：计算一个公式单元格并更新缓存结果或错误状态。

        使用方法：工作簿计算循环及依赖引用内部调用。
        参数：``worksheet`` 为所属表；``row``、``column`` 为0-based坐标。
        返回：公式结果。
        异常：解析、循环引用、函数或运算错误统一转为 ``FormulaCalculationError``。
        """
        coordinate = (worksheet, row, column)
        if coordinate in self.visiting:
            raise FormulaCalculationError(f"检测到循环引用：{worksheet.name}!{worksheet.cell(row, column).address}")
        formula = worksheet._formulas.get((row, column))
        if formula is None:
            return worksheet._values.get(row, column)
        self.visiting.add(coordinate)
        try:
            prepared, variables = self.prepare(worksheet, formula.lstrip("="))
            tree = ast.parse(prepared, mode="eval")
            result = self.node(tree, variables)
            if isinstance(result, float) and not math.isfinite(result):
                raise FormulaCalculationError("公式结果必须是有限数值")
            if not isinstance(result, (str, bool, int, float, date, datetime)) and result is not None:
                raise FormulaCalculationError("公式结果不是受支持的单元格类型")
            worksheet._formula_values[(row, column)] = result
            worksheet._formula_errors.pop((row, column), None)
            return result
        except FormulaCalculationError:
            raise
        except Exception as error:
            raise FormulaCalculationError(str(error) or type(error).__name__) from error
        finally:
            self.visiting.remove(coordinate)


def calculate_workbook(workbook: "Workbook", *, strict: bool = False) -> None:
    """功能：按工作表和坐标顺序计算全部公式并记录逐格状态。

    使用方法：由 ``Workbook.calculate()`` 唯一公开入口调用。
    参数：``workbook`` 为待计算对象；``strict`` 控制首个错误是否立即抛出。
    返回：``None``；成功结果写入缓存，失败保留旧缓存并记录错误。
    异常：严格模式下抛出带单元格地址的 ``FormulaCalculationError``。
    """
    evaluator = _Evaluator(workbook)
    for worksheet in workbook.sheets:
        worksheet._formula_errors.clear()
    for worksheet in workbook.sheets:
        for row, column in sorted(worksheet._formulas):
            try:
                evaluator.calculate_cell(worksheet, row, column)
            except FormulaCalculationError as error:
                address = worksheet.cell(row, column).address
                message = f"{worksheet.name}!{address}: {error}"
                worksheet._formula_errors[(row, column)] = message
                if strict:
                    raise FormulaCalculationError(message) from error


__all__ = ["calculate_workbook"]
