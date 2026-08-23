"""工作簿标量标签和循环行块的唯一模板渲染实现。"""

from __future__ import annotations

import ast
import math
import operator
import re
from collections.abc import Iterable, Mapping
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Tuple, Union

from .address import MAX_ROW
from .core.worksheet import _normalize_value
from .errors import TemplateError

if TYPE_CHECKING:
    from .core.workbook import Workbook
    from .core.worksheet import Worksheet
    from .style import Style

_RAW_PLACEHOLDER_PATTERN = re.compile(r"\{([^{}\r\n]+)\}")
_PATH_ROOT_TEXT = r"(?:[^\W\d]|_)[\w-]*"
_PATH_TEXT = rf"{_PATH_ROOT_TEXT}(?:\.(?:[\w-]+|@index))*"
_PLAIN_PATH_PATTERN = re.compile(rf"^(?:{_PATH_TEXT}|\.)$")
_VARIABLE_PATTERN = re.compile(rf"(?<![\w.@]){_PATH_TEXT}(?![\w.@])")
_FORMAT_FILTER_PATTERN = re.compile(
    r'''^(.*?)\s*\|\s*format\s*:\s*(?:"([^"]*)"|'([^']*)')\s*$'''
)
_START_PATTERN = re.compile(r"^\{loop\s+([\w-]+(?:\.[\w-]+)*)\}$")
_END_PATTERN = re.compile(r"^\{/loop\}$")
_CELL_REFERENCE_PATTERN = re.compile(
    r"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)(?!\d|\s*\()"
)
_MISSING = object()
_MAX_EXPRESSION_LENGTH = 256
_MAX_EXPRESSION_NODES = 64
_BINARY_OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.FloorDiv: operator.floordiv,
    ast.Mod: operator.mod,
}
_UNARY_OPERATORS = {
    ast.UAdd: operator.pos,
    ast.USub: operator.neg,
}


def _member(value: Any, name: str) -> Any:
    """功能：从映射、序列或对象属性中读取一个路径成员。

    使用方法：模板路径逐段解析时内部调用。
    参数：``value`` 为当前对象；``name`` 为一个路径段，序列索引采用 0-based。
    返回：找到的成员值；成员不存在时返回内部缺失标记。
    """
    if isinstance(value, Mapping):
        return value.get(name, _MISSING)
    if isinstance(value, (list, tuple)) and name.isdigit():
        index = int(name)
        return value[index] if 0 <= index < len(value) else _MISSING
    if not name.startswith("_") and hasattr(value, name):
        return getattr(value, name)
    return _MISSING


def _path(value: Any, expression: str) -> Any:
    """功能：按点分路径从给定对象解析嵌套值。

    使用方法：``_path(data, 'customer.name')`` 由模板解析内部调用。
    参数：``value`` 为路径根对象；``expression`` 为点分成员路径。
    返回：解析成功的任意值；任一路径段不存在时返回内部缺失标记。
    """
    current = value
    for name in expression.split("."):
        current = _member(current, name)
        if current is _MISSING:
            break
    return current


def _resolve(
    expression: str,
    context: Mapping[str, Any],
    item: Any = _MISSING,
    item_index: Optional[int] = None,
    loop_name: Optional[str] = None,
) -> Any:
    """功能：按显式循环前缀或根数据路径解析标签。

    使用方法：替换单元格标签和解析循环集合时内部调用。
    参数：``expression`` 为标签路径；``context`` 为根映射；``item`` 为当前循环
    元素或缺失标记；``item_index`` 为当前元素的 0-based 索引；``loop_name``
    为循环集合路径，最后一个路径段作为当前元素标签前缀。
    返回：解析出的任意值；没有匹配项时返回内部缺失标记。
    """
    if expression == ".":
        return item
    if item is not _MISSING and loop_name:
        alias = loop_name.rsplit(".", 1)[-1]
        if expression == alias:
            return item
        if expression == f"{alias}.@index":
            return item_index
        if expression.startswith(f"{alias}."):
            return _path(item, expression[len(alias) + 1:])
    return _path(context, expression)


def _numeric(value: Any, expression: str) -> Any:
    """功能：校验模板算术表达式的操作数和中间结果是否为受支持数值。

    使用方法：安全表达式求值器在执行每个一元或二元运算前后调用。
    参数：``value`` 为待校验对象；``expression`` 为用于错误信息的原始表达式。
    返回：原 ``int`` 或有限 ``float`` 数值。
    异常：布尔值、非数值对象或非有限浮点数抛出 :class:`TemplateError`。
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TemplateError(f"模板表达式只支持整数和浮点数：{{{expression}}}")
    if isinstance(value, float) and not math.isfinite(value):
        raise TemplateError(f"模板表达式不能产生无穷大或 NaN：{{{expression}}}")
    return value


def _evaluate_node(node: ast.AST, variables: Mapping[str, Any], expression: str) -> Any:
    """功能：递归计算已经通过 Python AST 解析的安全数值节点。

    使用方法：由 :func:`_calculate` 调用；不会执行函数、下标、属性或任意代码。
    参数：``node`` 为当前 AST 节点；``variables`` 为内部变量和值的映射；
    ``expression`` 为原始模板表达式。
    返回：计算得到的 ``int`` 或有限 ``float``。
    异常：节点、运算符或数据类型不在白名单中时抛出 :class:`TemplateError`。
    """
    if isinstance(node, ast.Expression):
        return _evaluate_node(node.body, variables, expression)
    if isinstance(node, ast.Constant):
        return _numeric(node.value, expression)
    if isinstance(node, ast.Name) and node.id in variables:
        return _numeric(variables[node.id], expression)
    if isinstance(node, ast.UnaryOp) and type(node.op) in _UNARY_OPERATORS:
        operand = _evaluate_node(node.operand, variables, expression)
        return _numeric(_UNARY_OPERATORS[type(node.op)](operand), expression)
    if isinstance(node, ast.BinOp) and type(node.op) in _BINARY_OPERATORS:
        left = _evaluate_node(node.left, variables, expression)
        right = _evaluate_node(node.right, variables, expression)
        try:
            result = _BINARY_OPERATORS[type(node.op)](left, right)
        except ArithmeticError as error:
            raise TemplateError(
                f"模板表达式计算失败：{{{expression}}}；{error}"
            ) from error
        return _numeric(result, expression)
    raise TemplateError(
        "模板表达式只允许数值、数据路径、括号以及 +、-、*、/、//、% 运算"
        f"：{{{expression}}}"
    )


def _calculate(
    expression: str,
    context: Mapping[str, Any],
    item: Any,
    item_index: Optional[int],
    loop_name: Optional[str],
) -> Any:
    """功能：把数据路径替换为内部变量并安全计算数值表达式。

    使用方法：处理 ``items.@index + 1`` 或 ``quantity * price`` 等标签时调用。
    参数：``expression`` 为不含格式过滤器的表达式；``context`` 为根数据；
    ``item``、``item_index`` 和 ``loop_name`` 描述当前循环作用域。
    返回：计算结果；任一数据路径缺失时返回内部缺失标记。
    异常：表达式过长、过于复杂、语法无效或包含非白名单能力时抛出
    :class:`TemplateError`。
    """
    if len(expression) > _MAX_EXPRESSION_LENGTH:
        raise TemplateError(
            f"模板表达式长度不能超过 {_MAX_EXPRESSION_LENGTH} 个字符"
        )
    variables: Dict[str, Any] = {}
    missing = False

    def replace(match: re.Match[str]) -> str:
        """功能：把一个模板数据路径替换成无副作用的内部 AST 变量名。

        使用方法：由变量路径正则替换过程自动调用。
        参数：``match`` 为一个根路径或循环元素路径匹配。
        返回：形如 ``_value_0`` 的内部变量名。
        """
        nonlocal missing
        path_expression = match.group(0)
        resolved = _resolve(
            path_expression, context, item, item_index, loop_name
        )
        variable_name = f"_value_{len(variables)}"
        variables[variable_name] = resolved
        if resolved is _MISSING:
            missing = True
        return variable_name

    translated = _VARIABLE_PATTERN.sub(replace, expression)
    if missing:
        return _MISSING
    if not variables:
        raise TemplateError(f"模板计算表达式必须包含数据路径：{{{expression}}}")
    try:
        tree = ast.parse(translated, mode="eval")
    except SyntaxError as error:
        raise TemplateError(f"无效的模板计算表达式：{{{expression}}}") from error
    if sum(1 for _node in ast.walk(tree)) > _MAX_EXPRESSION_NODES:
        raise TemplateError(
            f"模板表达式不能超过 {_MAX_EXPRESSION_NODES} 个语法节点"
        )
    return _evaluate_node(tree, variables, expression)


def _evaluate_expression(
    expression: str,
    context: Mapping[str, Any],
    item: Any = _MISSING,
    item_index: Optional[int] = None,
    loop_name: Optional[str] = None,
) -> Any:
    """功能：统一解析普通路径、数值计算和 ``format`` 显示格式过滤器。

    使用方法：每次替换 ``{...}`` 模板标签时调用。
    参数：``expression`` 为标签内部文本；``context`` 为根数据；``item``、
    ``item_index`` 和 ``loop_name`` 为可选循环作用域。
    返回：原始路径值、数值计算结果或格式化后的字符串；路径缺失时返回内部标记。
    异常：过滤器、格式字符串或算术表达式无效时抛出 :class:`TemplateError`。
    """
    expression = expression.strip()
    format_spec: Optional[str] = None
    filter_match = _FORMAT_FILTER_PATTERN.fullmatch(expression)
    if filter_match is not None:
        expression = filter_match.group(1).strip()
        format_spec = (
            filter_match.group(2)
            if filter_match.group(2) is not None
            else filter_match.group(3)
        )
    elif "|" in expression:
        raise TemplateError(
            "模板只支持一个格式过滤器，写法为 | format:\"格式字符串\""
        )

    if _PLAIN_PATH_PATTERN.fullmatch(expression):
        result = _resolve(expression, context, item, item_index, loop_name)
    else:
        result = _calculate(expression, context, item, item_index, loop_name)
    if result is _MISSING or format_spec is None:
        return result
    try:
        return format(result, format_spec)
    except (TypeError, ValueError) as error:
        raise TemplateError(
            f"模板格式字符串无效：{{{expression} | format:\"{format_spec}\"}}"
        ) from error


def _placeholder_matches(value: str) -> List[re.Match[str]]:
    """功能：找出包含数据路径的模板标签并忽略普通 Excel 花括号内容。

    使用方法：渲染字符串前调用，避免把 Excel 数组常量 ``{1,2}`` 当作模板。
    参数：``value`` 为可能含有一个或多个标签的字符串。
    返回：按原字符串位置排列的正则匹配列表。
    """
    return [
        match for match in _RAW_PLACEHOLDER_PATTERN.finditer(value)
        if match.group(1).strip() == "."
        or _VARIABLE_PATTERN.search(match.group(1)) is not None
    ]


def _render_text(
    value: str,
    context: Mapping[str, Any],
    strict: bool,
    item: Any = _MISSING,
    item_index: Optional[int] = None,
    loop_name: Optional[str] = None,
    drop_on_missing: bool = False,
) -> Any:
    """功能：替换字符串中的模板标签，并为整格标签保留原数据类型。

    使用方法：渲染普通值和带标签公式时内部调用。
    参数：``value`` 为模板字符串；``context`` 为根数据；``strict`` 控制缺失标签
    是否报错；``item``、``item_index``、``loop_name`` 为可选循环作用域；
    ``drop_on_missing`` 为 ``True`` 时任一缺失值返回内部缺失标记，供公式整体删除。
    返回：整格只有一个标签时返回原类型值；混合文本返回替换后的字符串；非严格
    模式的整格缺失标签返回 ``None``，公式缺失返回内部缺失标记。
    异常：严格模式遇到缺失标签时抛出 :class:`TemplateError`。
    """
    matches = _placeholder_matches(value)
    if not matches:
        return value
    if len(matches) == 1 and matches[0].span() == (0, len(value)):
        expression = matches[0].group(1).strip()
        result = _evaluate_expression(
            expression, context, item, item_index, loop_name
        )
        if result is _MISSING:
            if strict:
                raise TemplateError(f"模板数据缺少标签：{{{expression}}}")
            return _MISSING if drop_on_missing else None
        return result

    parts: List[str] = []
    position = 0
    for match in matches:
        parts.append(value[position:match.start()])
        expression = match.group(1).strip()
        result = _evaluate_expression(
            expression, context, item, item_index, loop_name
        )
        if result is _MISSING:
            if strict:
                raise TemplateError(f"模板数据缺少标签：{{{expression}}}")
            if drop_on_missing:
                return _MISSING
        elif result is not None:
            parts.append(str(result))
        elif drop_on_missing:
            return _MISSING
        position = match.end()
    parts.append(value[position:])
    return "".join(parts)


def _translate_formula(formula: str, row_offset: int) -> str:
    """功能：按行偏移调整公式中的相对 A1 行引用。

    使用方法：循环复制公式或移动循环下方公式时内部调用。
    参数：``formula`` 为带或不带等号的公式；``row_offset`` 为有符号行偏移量。
    返回：相对行号已调整、绝对行号保持不变的公式字符串。
    异常：调整后行号越出 Excel 上限时抛出 :class:`TemplateError`。
    """
    if row_offset == 0:
        return formula

    def replace(match: re.Match[str]) -> str:
        """功能：调整一个公式单元格引用的相对行号。

        使用方法：由正则替换过程自动调用。
        参数：``match`` 为包含绝对标记、列字母和行号的正则匹配。
        返回：调整后的单元格引用字符串。
        """
        column_absolute, column, row_absolute, row_text = match.groups()
        if row_absolute:
            return match.group(0)
        row_number = int(row_text) + row_offset
        if not 1 <= row_number <= MAX_ROW:
            raise TemplateError("循环展开后的公式行引用超出 Excel 上限")
        return f"{column_absolute}{column}{row_absolute}{row_number}"

    segments = re.split(r'("(?:[^"]|"")*")', formula)
    return "".join(
        segment if index % 2 else _CELL_REFERENCE_PATTERN.sub(replace, segment)
        for index, segment in enumerate(segments)
    )


def _loop_blocks(
    values: Dict[Tuple[int, int], Any],
    formulas: Dict[Tuple[int, int], str],
) -> List[Tuple[int, int, str]]:
    """功能：校验并配对 ``{loop 数据}`` 与 ``{/loop}`` 标记行。

    使用方法：每张工作表开始渲染前调用。
    参数：``values``、``formulas`` 为工作表值和公式快照。
    返回：按从上到下顺序排列的 ``(开始行, 结束行, 数据路径)`` 列表。
    异常：标记行包含其他内容、标记未配对或循环嵌套时抛出
    :class:`TemplateError`。
    """
    candidates: Dict[int, List[Tuple[Tuple[int, int], str, Optional[str]]]] = {}
    for coordinate, value in sorted(values.items()):
        if not isinstance(value, str):
            continue
        text = value.strip()
        start = _START_PATTERN.fullmatch(text)
        end = _END_PATTERN.fullmatch(text)
        if start:
            candidates.setdefault(coordinate[0], []).append(
                (coordinate, "start", start.group(1))
            )
        elif end:
            candidates.setdefault(coordinate[0], []).append((coordinate, "end", None))

    markers: List[Tuple[int, str, Optional[str]]] = []
    for row, row_markers in sorted(candidates.items()):
        if len(row_markers) != 1:
            raise TemplateError("每个循环标记行只能包含一个循环标签")
        marker_coordinate, kind, name = row_markers[0]
        has_other_value = any(
            coordinate[0] == row and coordinate != marker_coordinate
            for coordinate in values
        )
        has_formula = any(coordinate[0] == row for coordinate in formulas)
        if has_other_value or has_formula:
            raise TemplateError("循环开始行和结束行不能包含其他值或公式")
        markers.append((row, kind, name))

    blocks: List[Tuple[int, int, str]] = []
    opened: Optional[Tuple[int, str]] = None
    for row, kind, name in markers:
        if kind == "start":
            if opened is not None:
                raise TemplateError("当前模板循环不支持嵌套")
            opened = (row, name or "")
        else:
            if opened is None:
                raise TemplateError("循环结束标签 {/loop} 没有对应的开始标签")
            start_row, start_name = opened
            if row <= start_row + 1:
                raise TemplateError(f"循环 {start_name!r} 必须至少包含一行模板内容")
            blocks.append((start_row, row, start_name))
            opened = None
    if opened is not None:
        raise TemplateError(f"循环 {opened[1]!r} 缺少结束标签 {{/loop}}")
    return blocks


def _sequence(value: Any, name: str, strict: bool) -> List[Any]:
    """功能：把循环数据物化为可重复遍历的列表。

    使用方法：展开每个循环块前调用。
    参数：``value`` 为标签解析结果；``name`` 为用于错误信息的数据路径；
    ``strict`` 为 ``False`` 时缺失集合按空列表处理，为 ``True`` 时缺失即报错。
    返回：由原可迭代对象物化得到的列表。
    异常：字符串、字节、映射或不可迭代值会抛出 :class:`TemplateError`。
    """
    if value is _MISSING:
        if strict:
            raise TemplateError(f"模板数据缺少循环集合：{{loop {name}}}")
        return []
    if isinstance(value, (str, bytes, Mapping)) or not isinstance(value, Iterable):
        raise TemplateError(f"循环数据 {name!r} 必须是列表或其他非映射可迭代对象")
    return list(value)


def _render_sheet(
    worksheet: "Worksheet", context: Mapping[str, Any], strict: bool
) -> Tuple[
    Dict[Tuple[int, int], Any],
    Dict[Tuple[int, int], str],
    Dict[Tuple[int, int], "Style"],
    int,
    int,
]:
    """功能：在工作表快照中替换标量标签并展开循环行块。

    使用方法：工作簿渲染器为每张表调用，成功后才统一提交结果。
    参数：``worksheet`` 为源工作表；``context`` 为根数据；``strict`` 控制缺失
    普通标签是否报错。
    返回：新值、新公式、新样式、最大行索引和最大列索引组成的元组。
    异常：模板结构、数据类型、标签解析或行数无效时抛出
    :class:`TemplateError`。
    """
    values = dict(worksheet._values.items())
    formulas = dict(worksheet._formulas)
    styles = dict(worksheet._styles)
    blocks = _loop_blocks(values, formulas)
    loop_rows = {
        row for start_row, end_row, _name in blocks
        for row in range(start_row, end_row + 1)
    }

    for coordinate, value in list(values.items()):
        if coordinate[0] not in loop_rows and isinstance(value, str):
            rendered = _normalize_value(_render_text(value, context, strict))
            if rendered is None:
                values.pop(coordinate)
            else:
                values[coordinate] = rendered
    for coordinate, formula in list(formulas.items()):
        if coordinate[0] not in loop_rows:
            rendered_formula = _render_text(
                formula, context, strict, drop_on_missing=True
            )
            if rendered_formula is _MISSING or rendered_formula is None:
                formulas.pop(coordinate)
            elif not isinstance(rendered_formula, str):
                raise TemplateError("公式模板渲染结果必须是字符串")
            else:
                formulas[coordinate] = rendered_formula

    for start_row, end_row, name in reversed(blocks):
        items = _sequence(_path(context, name), name, strict)
        template_start = start_row + 1
        template_height = end_row - template_start
        expanded_height = template_height * len(items)
        removed_height = end_row - start_row + 1
        delta = expanded_height - removed_height
        coordinates = set(values) | set(formulas) | set(styles)
        current_max_row = max((row for row, _column in coordinates), default=-1)
        if current_max_row + delta >= MAX_ROW:
            raise TemplateError("循环展开后的工作表超过 Excel 最大行数")

        template_values = {
            (row, column): value for (row, column), value in values.items()
            if template_start <= row < end_row
        }
        template_formulas = {
            (row, column): formula for (row, column), formula in formulas.items()
            if template_start <= row < end_row
        }
        template_styles = {
            (row, column): style for (row, column), style in styles.items()
            if template_start <= row < end_row
        }

        new_values: Dict[Tuple[int, int], Any] = {}
        new_formulas: Dict[Tuple[int, int], str] = {}
        new_styles: Dict[Tuple[int, int], "Style"] = {}
        for (row, column), value in values.items():
            if start_row <= row <= end_row:
                continue
            new_row = row + delta if row > end_row else row
            new_values[(new_row, column)] = value
        for (row, column), formula in formulas.items():
            if start_row <= row <= end_row:
                continue
            new_row = row + delta if row > end_row else row
            new_formulas[(new_row, column)] = (
                _translate_formula(formula, delta) if row > end_row else formula
            )
        for (row, column), style in styles.items():
            if start_row <= row <= end_row:
                continue
            new_row = row + delta if row > end_row else row
            new_styles[(new_row, column)] = style

        for item_index, item in enumerate(items):
            for (row, column), value in template_values.items():
                destination_row = (
                    start_row + item_index * template_height + row - template_start
                )
                rendered = (
                    _render_text(value, context, strict, item, item_index, name)
                    if isinstance(value, str) else value
                )
                rendered = _normalize_value(rendered)
                if rendered is not None:
                    new_values[(destination_row, column)] = rendered
            for (row, column), formula in template_formulas.items():
                destination_row = (
                    start_row + item_index * template_height + row - template_start
                )
                rendered_formula = _render_text(
                    formula,
                    context,
                    strict,
                    item,
                    item_index,
                    name,
                    drop_on_missing=True,
                )
                if rendered_formula is _MISSING or rendered_formula is None:
                    continue
                if not isinstance(rendered_formula, str):
                    raise TemplateError("公式模板渲染结果必须是字符串")
                new_formulas[(destination_row, column)] = _translate_formula(
                    rendered_formula, destination_row - row
                )
            for (row, column), style in template_styles.items():
                destination_row = (
                    start_row + item_index * template_height + row - template_start
                )
                new_styles[(destination_row, column)] = style
        values, formulas, styles = new_values, new_formulas, new_styles

    coordinates = set(values) | set(formulas) | set(styles)
    max_row = max((row for row, _column in coordinates), default=-1)
    max_column = max((column for _row, column in coordinates), default=-1)
    return values, formulas, styles, max_row, max_column


def render_workbook(
    workbook: "Workbook",
    context: Optional[Mapping[str, Any]] = None,
    *,
    by_sheet: Optional[
        Mapping[Union[str, int], Mapping[str, Any]]
    ] = None,
    strict: bool = False,
) -> "Workbook":
    """功能：使用公共或分工作表上下文原子渲染标量标签和循环行块。

    使用方法：由 ``workbook.render(data)`` 唯一公开入口调用。
    参数：``workbook`` 为待渲染工作簿；``context`` 为共享根映射或 ``None``；
    ``by_sheet`` 以工作表名称或0-based索引映射到独立根数据；``strict`` 为
    ``True`` 时缺失数据立即报错，为 ``False`` 时标签按空值处理。
    返回：传入的同一个 Workbook，支持继续 ``save()`` 链式调用。
    异常：参数、工作表标识、数据或模板无效时抛出相应异常；失败时所有目标
    工作表保持原状。
    """
    if context is None:
        context = {}
    elif not isinstance(context, Mapping):
        raise TypeError("模板数据必须是映射对象")
    if not isinstance(strict, bool):
        raise TypeError("strict 必须是布尔值")
    targets: List[Tuple["Worksheet", Mapping[str, Any]]] = []
    if by_sheet is None:
        targets = [(sheet, context) for sheet in workbook.sheets]
    else:
        if not isinstance(by_sheet, Mapping):
            raise TypeError("by_sheet 必须是工作表标识到独立数据的映射")
        seen = set()
        for identifier, local_context in by_sheet.items():
            worksheet = workbook.sheet(identifier)
            if worksheet in seen:
                raise ValueError("by_sheet 不能用不同标识重复指定同一张工作表")
            if not isinstance(local_context, Mapping):
                raise TypeError("by_sheet 中每张工作表的数据都必须是映射对象")
            merged_context = dict(context)
            merged_context.update(local_context)
            targets.append((worksheet, merged_context))
            seen.add(worksheet)

    results = [
        _render_sheet(worksheet, sheet_context, strict)
        for worksheet, sheet_context in targets
    ]
    for (worksheet, _sheet_context), result in zip(targets, results):
        values, formulas, styles, max_row, max_column = result
        worksheet._values._values = values
        worksheet._formulas = formulas
        worksheet._styles = styles
        worksheet._max_row = max_row
        worksheet._max_column = max_column
    return workbook


__all__ = ["render_workbook"]
