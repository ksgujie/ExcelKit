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
_NAME_PATTERN = re.compile(r"(?<![\w.])([A-Za-z_\u0080-\uffff][\w.\u0080-\uffff]*)(?![\w.])")


def formula_dependencies(workbook: "Workbook", worksheet: "Worksheet", formula: str) -> tuple[Any, ...]:
    """功能：解析公式中直接引用的全部单元格。

    使用方法：``cell.dependencies`` 属性内部调用；也可供扩展模块调用。
    参数：``workbook`` 为公式所属工作簿；``worksheet`` 为默认引用工作表；
    ``formula`` 为带或不带 ``=`` 的 Excel 公式文本。
    返回：按公式出现顺序去重后的 ``Cell`` 元组；区域引用会展开为每个单元格，
    跨表引用会返回对应工作表的单元格对象。
    异常：工作表名称或地址无效时抛出原始地址异常；公式不是字符串时抛出
    ``TypeError``。
    """
    if not isinstance(formula, str):
        raise TypeError("formula 必须是字符串")
    result: list[Any] = []
    seen: set[tuple[int, int, int]] = set()
    expression = formula.lstrip("=")
    for segment in _STRING_PATTERN.split(expression)[::2]:
        for match in _REFERENCE_PATTERN.finditer(segment):
            quoted_sheet, plain_sheet, start, end = match.groups()
            try:
                start_row, start_column = cell_index(start.replace("$", ""))
                if end is None:
                    end_row, end_column = start_row, start_column
                else:
                    end_row, end_column = cell_index(end.replace("$", ""))
            except ValueError:
                continue
            target = worksheet
            sheet_name = quoted_sheet or plain_sheet
            if sheet_name is not None:
                target = workbook.sheet(sheet_name.replace("''", "'"))
            for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
                for column in range(min(start_column, end_column), max(start_column, end_column) + 1):
                    key = (id(target), row, column)
                    if key not in seen:
                        seen.add(key)
                        result.append(target.cell(row, column))
    return tuple(result)


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


def _round_direction(value: Any, digits: int, *, up: bool) -> float | int:
    """功能：按远离零或接近零方向进行 Excel 风格数值舍入。

    使用方法：分别由 ``ROUNDUP`` 与 ``ROUNDDOWN`` 公式函数调用。
    参数：``value`` 为有限数值；``digits`` 为可正可负的整数；``up`` 为真时远离零。
    返回：舍入后的整数或浮点数。
    异常：参数类型或取值无效时抛出 ``FormulaCalculationError``。
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise FormulaCalculationError("舍入的值必须是数值")
    if isinstance(digits, bool) or not isinstance(digits, int):
        raise FormulaCalculationError("舍入位数必须是整数")
    factor = 10 ** digits
    scaled = float(value) * factor
    rounded = math.ceil(abs(scaled)) if up else math.floor(abs(scaled))
    result = math.copysign(rounded / factor, scaled)
    return int(result) if digits <= 0 else result


def _criterion_matches(value: Any, criterion: Any) -> bool:
    """功能：按常见 Excel 条件表达式判断一个值是否匹配。

    使用方法：由 ``SUMIF``、``COUNTIF``、``AVERAGEIF`` 内部调用。
    参数：``value`` 为待判断值；``criterion`` 为文本、数值或含 ``>=`` 等运算符的条件。
    返回：匹配时返回 ``True``。
    """
    if not isinstance(criterion, str):
        return value == criterion
    match = re.match(r"^(<=|>=|<>|=|<|>)(.*)$", criterion)
    if match is None:
        return str(value) == criterion
    operator, expected = match.groups()
    try:
        comparison_value: Any = float(expected)
        actual: Any = float(value)
    except (TypeError, ValueError):
        comparison_value = expected
        actual = "" if value is None else str(value)
    return {
        "=": actual == comparison_value,
        "<>": actual != comparison_value,
        ">": actual > comparison_value,
        ">=": actual >= comparison_value,
        "<": actual < comparison_value,
        "<=": actual <= comparison_value,
    }[operator]


def _conditional_values(criteria_range: Any, criterion: Any, sum_range: Any = None) -> list[Any]:
    """功能：按条件区域筛选对应数据区域的值。

    使用方法：由条件聚合函数内部调用。
    参数：``criteria_range`` 为区域或序列；``criterion`` 为条件；``sum_range`` 为可选
    的同长度目标区域，省略时使用条件区域本身。
    返回：所有匹配位置的目标值列表。
    异常：区域元素数不一致时抛出 ``FormulaCalculationError``。
    """
    criteria = _flatten([criteria_range])
    target = _flatten([criteria_range if sum_range is None else sum_range])
    if len(criteria) != len(target):
        raise FormulaCalculationError("条件区域与目标区域大小必须一致")
    return [target[index] for index, value in enumerate(criteria) if _criterion_matches(value, criterion)]


def _sumif(criteria_range: Any, criterion: Any, sum_range: Any = None) -> float | int:
    """功能：对满足条件的位置求和。

    使用方法：公式 ``=SUMIF(A1:A10,\">0\",B1:B10)``。
    参数：``criteria_range`` 为条件区域；``criterion`` 为 Excel 条件；``sum_range``
    为可选的同尺寸求和区域。
    返回：匹配位置中有效数值的总和；没有匹配数值时返回 ``0``。
    """
    return sum(_numbers(_conditional_values(criteria_range, criterion, sum_range)))


def _countif(criteria_range: Any, criterion: Any) -> int:
    """功能：统计满足条件的位置数量。

    使用方法：公式 ``=COUNTIF(A1:A10,\"通过\")``。
    参数：``criteria_range`` 为条件区域；``criterion`` 为 Excel 条件。
    返回：满足条件的元素数量。
    """
    return len(_conditional_values(criteria_range, criterion))


def _averageif(criteria_range: Any, criterion: Any, average_range: Any = None) -> float:
    """功能：计算满足条件位置的数值平均值。

    使用方法：公式 ``=AVERAGEIF(A1:A10,\">0\")``。
    参数：``criteria_range`` 为条件区域；``criterion`` 为 Excel 条件；
    ``average_range`` 为可选的同尺寸目标区域。
    返回：匹配位置有效数值的浮点平均值。
    异常：没有可计算的数值时抛出 ``FormulaCalculationError``。
    """
    values = _numbers(_conditional_values(criteria_range, criterion, average_range))
    if not values:
        raise FormulaCalculationError("AVERAGEIF 没有可计算的数值")
    return sum(values) / len(values)


def _criteria_indexes(*criteria_pairs: Any) -> list[int]:
    """功能：计算多组条件区域同时满足条件的位置索引。

    使用方法：由 ``SUMIFS`` 和 ``COUNTIFS`` 内部调用。
    参数：``criteria_pairs`` 必须按 ``区域、条件`` 成对传入；所有区域长度必须一致。
    返回：满足全部条件的 0-based 一维位置索引列表。
    异常：参数数量不是偶数或区域长度不一致时抛出 ``FormulaCalculationError``。
    """
    if not criteria_pairs or len(criteria_pairs) % 2:
        raise FormulaCalculationError("条件函数必须传入成对的区域和条件")
    first = _flatten([criteria_pairs[0]])
    ranges = [first]
    criteria = list(criteria_pairs[1::2])
    for item in criteria_pairs[2::2]:
        values = _flatten([item])
        if len(values) != len(first):
            raise FormulaCalculationError("所有条件区域大小必须一致")
        ranges.append(values)
    return [
        index for index in range(len(first))
        if all(_criterion_matches(ranges[pos][index], criteria[pos]) for pos in range(len(ranges)))
    ]


def _sumifs(sum_range: Any, *criteria_pairs: Any) -> float | int:
    """功能：对同时满足多组条件的位置求和。

    使用方法：公式 ``=SUMIFS(C2:C100,A2:A100,"销售",B2:B100,"华东")``。
    参数：``sum_range`` 为求和区域；其余参数按条件区域、条件成对传入。
    返回：匹配位置的有效数值总和；没有匹配数值时返回 ``0``。
    异常：区域大小或参数结构无效时抛出 ``FormulaCalculationError``。
    """
    values = _flatten([sum_range])
    if criteria_pairs:
        first_length = len(_flatten([criteria_pairs[0]]))
        if len(values) != first_length:
            raise FormulaCalculationError("SUMIFS 求和区域与条件区域大小必须一致")
    indexes = _criteria_indexes(*criteria_pairs)
    if any(index >= len(values) for index in indexes):
        raise FormulaCalculationError("SUMIFS 求和区域与条件区域大小必须一致")
    return sum(_numbers(values[index] for index in indexes))


def _countifs(*criteria_pairs: Any) -> int:
    """功能：统计同时满足多组条件的位置数量。

    使用方法：公式 ``=COUNTIFS(A2:A100,"销售",B2:B100,"华东")``。
    参数：参数必须按条件区域、条件成对传入。
    返回：同时满足全部条件的元素数量。
    异常：参数结构或区域大小无效时抛出 ``FormulaCalculationError``。
    """
    return len(_criteria_indexes(*criteria_pairs))


def _index(array: Any, row_number: int, column_number: int = 1) -> Any:
    """功能：按 Excel 的 1-based 行列序号取得数组元素。

    使用方法：公式 ``=INDEX(A2:C10,2,3)`` 或 ``=INDEX(A2:A10,2)``。
    参数：``array`` 为一维或二维区域；``row_number`` 和 ``column_number`` 为正整数，
    均采用 Excel 1-based 序号。
    返回：指定位置的元素。
    异常：数组维度、序号或位置越界时抛出 ``FormulaCalculationError``。
    """
    if any(isinstance(item, bool) or not isinstance(item, int) or item < 1 for item in (row_number, column_number)):
        raise FormulaCalculationError("INDEX 的行列序号必须是正整数")
    rows = array if isinstance(array, (list, tuple)) else [array]
    if not rows:
        raise FormulaCalculationError("INDEX 数组不能为空")
    if not isinstance(rows[0], (list, tuple)):
        if column_number != 1 or row_number > len(rows):
            raise FormulaCalculationError("INDEX 位置超出数组范围")
        return rows[row_number - 1]
    if row_number > len(rows) or any(not isinstance(row, (list, tuple)) for row in rows):
        raise FormulaCalculationError("INDEX 位置超出数组范围")
    row = rows[row_number - 1]
    if column_number == 1 and all(len(item) == 1 for item in rows):
        return row[0]
    if column_number > len(row):
        raise FormulaCalculationError("INDEX 位置超出数组范围")
    return row[column_number - 1]


def _match(value: Any, lookup_array: Any, match_type: int = 0) -> int:
    """功能：在一维区域中查找值并返回 Excel 1-based 位置。

    使用方法：公式 ``=MATCH("李四",A2:A10,0)``；``match_type`` 为 0 时精确匹配，
    为 1 或 -1 时按升序或降序近似匹配。
    参数：``value`` 为查找值；``lookup_array`` 为一维区域；``match_type`` 为 0、1、-1。
    返回：匹配项的 1-based 位置。
    异常：数组不是一维、匹配类型无效或未找到时抛出 ``FormulaCalculationError``。
    """
    if isinstance(match_type, bool) or not isinstance(match_type, int) or match_type not in (-1, 0, 1):
        raise FormulaCalculationError("MATCH 的匹配类型必须是 -1、0 或 1")
    values = _flatten([lookup_array])
    if match_type == 0:
        for index, item in enumerate(values):
            if item == value:
                return index + 1
    elif match_type == 1:
        selected = None
        for index, item in enumerate(values):
            if item == value:
                return index + 1
            try:
                if item <= value:
                    selected = index + 1
            except TypeError:
                continue
        if selected is not None:
            return selected
    else:
        selected = None
        for index, item in enumerate(values):
            if item == value:
                return index + 1
            try:
                if item >= value:
                    selected = index + 1
                elif selected is not None:
                    break
            except TypeError:
                continue
        if selected is not None:
            return selected
    raise FormulaCalculationError("MATCH 未找到匹配项")


def _text(value: Any, format_text: str) -> str:
    """功能：按常用 Excel 格式代码把值格式化为文本。

    使用方法：公式 ``=TEXT(A1,"#,##0.00")`` 或 ``=TEXT(A1,"yyyy-mm-dd")``。
    参数：``value`` 为数值、日期或时间；``format_text`` 为格式代码字符串。
    返回：格式化后的字符串；不支持的格式使用 Python 字符串表示。
    异常：格式参数不是字符串时抛出 ``FormulaCalculationError``。
    """
    if not isinstance(format_text, str):
        raise FormulaCalculationError("TEXT 的格式代码必须是字符串")
    if isinstance(value, datetime):
        if format_text in {"yyyy-mm-dd", "yyyy/mm/dd"}:
            return value.strftime("%Y-%m-%d" if "-" in format_text else "%Y/%m/%d")
        if format_text in {"hh:mm", "hh:mm:ss"}:
            return value.strftime("%H:%M" if format_text == "hh:mm" else "%H:%M:%S")
    if isinstance(value, date):
        if format_text in {"yyyy-mm-dd", "yyyy/mm/dd"}:
            return value.strftime("%Y-%m-%d" if "-" in format_text else "%Y/%m/%d")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if format_text.endswith("%"):
            decimals = len(format_text.split(".", 1)[1].rstrip("%")) if "." in format_text else 0
            return f"{value * 100:.{decimals}f}%"
        if format_text in {"#,##0", "#,##0.00", "0.00"}:
            decimals = len(format_text.split(".", 1)[1]) if "." in format_text else 0
            return f"{value:,.{decimals}f}"
    return str(value)


def _date(year: int, month: int, day: int) -> date:
    """功能：按年、月、日构造日期。

    使用方法：公式 ``=DATE(2026,8,1)``。
    参数：``year``、``month``、``day`` 均为整数日期分量。
    返回：对应的 ``date`` 对象。
    异常：类型或日期无效时抛出 ``FormulaCalculationError``。
    """
    if any(isinstance(item, bool) or not isinstance(item, int) for item in (year, month, day)):
        raise FormulaCalculationError("DATE 的年、月、日必须是整数")
    try:
        return date(year, month, day)
    except ValueError as error:
        raise FormulaCalculationError(str(error)) from error


def _lookup_vector(value: Any, vector: Any, return_vector: Any, if_not_found: Any = None) -> Any:
    """功能：在一维查找向量中精确匹配并返回同位置结果。

    使用方法：由 ``XLOOKUP`` 内部调用。
    参数：``value`` 为目标；``vector`` 和 ``return_vector`` 为同长度区域；
    ``if_not_found`` 为未找到的返回值。
    返回：匹配项对应的返回值或未找到值。
    异常：向量长度不一致时抛出 ``FormulaCalculationError``。
    """
    lookup = _flatten([vector])
    results = _flatten([return_vector])
    if len(lookup) != len(results):
        raise FormulaCalculationError("XLOOKUP 查找区域和返回区域大小必须一致")
    for index, item in enumerate(lookup):
        if item == value:
            return results[index]
    if if_not_found is not None:
        return if_not_found
    raise FormulaCalculationError("XLOOKUP 未找到匹配项")


def _vlookup(value: Any, table: Any, column_index: int, approximate: bool = True) -> Any:
    """功能：在二维区域首列中查找并返回指定 1-based 列的值。

    使用方法：公式 ``=VLOOKUP(A1,D2:F10,2,FALSE)``。
    参数：``value`` 为查找值；``table`` 为二维区域；``column_index`` 为 Excel
    1-based 返回列序号；``approximate`` 为是否允许近似匹配。
    返回：匹配行中指定列的值。
    异常：索引、区域或查找结果无效时抛出 ``FormulaCalculationError``。
    """
    if isinstance(column_index, bool) or not isinstance(column_index, int) or column_index < 1:
        raise FormulaCalculationError("VLOOKUP 的列索引必须是正整数")
    rows = table if isinstance(table, list) else [table]
    selected = None
    for row in rows:
        if not isinstance(row, (list, tuple)) or len(row) < column_index:
            raise FormulaCalculationError("VLOOKUP 区域或列索引无效")
        if row[0] == value:
            return row[column_index - 1]
        if approximate and row[0] <= value:
            selected = row
    if selected is not None:
        return selected[column_index - 1]
    raise FormulaCalculationError("VLOOKUP 未找到匹配项")


def _hlookup(value: Any, table: Any, row_index: int, approximate: bool = True) -> Any:
    """功能：在二维区域首行中查找并返回指定 1-based 行的值。

    使用方法：公式 ``=HLOOKUP(A1,B1:E3,2,FALSE)``。
    参数：``value`` 为查找值；``table`` 为二维区域；``row_index`` 为 Excel
    1-based 返回行序号；``approximate`` 为是否允许近似匹配。
    返回：匹配列中指定行的值。
    异常：索引、区域或查找结果无效时抛出 ``FormulaCalculationError``。
    """
    if isinstance(row_index, bool) or not isinstance(row_index, int) or row_index < 1:
        raise FormulaCalculationError("HLOOKUP 的行索引必须是正整数")
    if not isinstance(table, list) or len(table) < row_index or not table:
        raise FormulaCalculationError("HLOOKUP 区域或行索引无效")
    header = table[0]
    selected = None
    for column, item in enumerate(header):
        if item == value:
            return table[row_index - 1][column]
        if approximate and item <= value:
            selected = column
    if selected is not None:
        return table[row_index - 1][selected]
    raise FormulaCalculationError("HLOOKUP 未找到匹配项")


def _date_error(name: str, value: Any) -> Any:
    """功能：为日期拆分函数生成统一的参数类型错误。

    使用方法：由 ``YEAR``、``MONTH``、``DAY`` 内部调用。
    参数：``name`` 为函数名称；``value`` 为收到的非日期值。
    返回：无，始终抛出异常。
    异常：抛出 ``FormulaCalculationError``。
    """
    raise FormulaCalculationError(f"{name} 需要日期或日期时间值，而不是 {type(value).__name__}")


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
    "ROUNDUP": lambda value, digits=0: _round_direction(value, digits, up=True),
    "ROUNDDOWN": lambda value, digits=0: _round_direction(value, digits, up=False),
    "CONCAT": _concat,
    "LEN": lambda value: len(str(value)),
    "LEFT": _left,
    "RIGHT": _right,
    "MID": _mid,
    "SUMIF": _sumif,
    "COUNTIF": _countif,
    "AVERAGEIF": _averageif,
    "SUMIFS": _sumifs,
    "COUNTIFS": _countifs,
    "INDEX": _index,
    "MATCH": _match,
    "TEXT": _text,
    "DATE": _date,
    "YEAR": lambda value: value.year if isinstance(value, (date, datetime)) else _date_error("YEAR", value),
    "MONTH": lambda value: value.month if isinstance(value, (date, datetime)) else _date_error("MONTH", value),
    "DAY": lambda value: value.day if isinstance(value, (date, datetime)) else _date_error("DAY", value),
    "TODAY": lambda: date.today(),
    "NOW": lambda: datetime.now().replace(microsecond=0),
    "VLOOKUP": _vlookup,
    "HLOOKUP": _hlookup,
    "XLOOKUP": _lookup_vector,
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

            def replace_named_range(match: re.Match[str]) -> str:
                """功能：把公式中的工作簿命名区域替换为内部变量。

                使用方法：由本公式片段的名称替换过程自动调用。
                参数：``match`` 为候选标识符匹配对象。
                返回：名称不是工作簿命名区域或是函数调用时返回原文本；否则返回变量名。
                """
                name = match.group(1)
                following = replaced[match.end():match.end() + 1]
                if following == "(" or name.upper() in {"TRUE", "FALSE"}:
                    return name
                try:
                    named_range = self.workbook.named_range(name)
                except KeyError:
                    return name
                variable = f"_ref_{len(variables)}"
                variables[variable] = self.reference(
                    named_range.worksheet, named_range.range.address
                )
                return variable

            replaced = _NAME_PATTERN.sub(replace_named_range, replaced)
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
            if name == "IFERROR":
                if len(node.args) != 2:
                    raise FormulaCalculationError("IFERROR 需要2个参数")
                try:
                    return self.node(node.args[0], variables)
                except Exception:
                    return self.node(node.args[1], variables)
            if name == "IFNA":
                if len(node.args) != 2:
                    raise FormulaCalculationError("IFNA 需要2个参数")
                try:
                    return self.node(node.args[0], variables)
                except FormulaCalculationError:
                    return self.node(node.args[1], variables)
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
