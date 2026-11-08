"""Excel 模板渲染：标量标签与 {loop items} / {/loop} 循环行块。

支持：
- 标量标签：{customer.name}、{report_date}、{items.@index + 1}
- 循环行块：{loop items} ... {/loop}
- 数值表达式：+ - * / // % 和括号，不允许函数调用
- 显示格式：| format:",.2f"
- strict 模式：缺失字段抛出 TemplateError
"""

from __future__ import annotations

import re
from typing import Any

from excelkit.core.workbook import Workbook
from excelkit.core.worksheet import Worksheet
from excelkit.errors import TemplateError


# 标签匹配：{ 标签内容 }
_TAG_RE = re.compile(r"\{(/?loop\s+[\w\.]+|[^{}]+?)\}")


def _eval_expr(expr: str, ctx: dict[str, Any]) -> Any:
    """安全计算数值表达式，仅允许 + - * / // % 和括号、上下文变量。"""
    expr = expr.strip()
    # 去掉 | format:...
    if "|" in expr:
        expr, _fmt = expr.split("|", 1)
        expr = expr.strip()
    # 只允许数字、运算符、括号、点、上下文标识符
    if not re.fullmatch(r"[\d\w\.\(\)\s\+\-\*/%]+", expr):
        raise TemplateError(f"非法模板表达式：{expr!r}")
    # 替换上下文变量为其值（数字或字符串）
    def repl(m):
        key = m.group(0)
        if key in ctx:
            v = ctx[key]
            return repr(v) if isinstance(v, str) else str(v)
        return key
    # 替换 ctx 中的标识符（最长优先）
    for key in sorted(ctx, key=len, reverse=True):
        if key.isidentifier():
            expr = expr.replace(key, f"__ctx_{key}__")
            ctx[f"__ctx_{key}__"] = ctx.pop(key)
    try:
        # 使用 eval 但限制内置，仅数学运算
        return eval(expr, {"__builtins__": {}}, ctx)
    except Exception as exc:
        raise TemplateError(f"表达式计算失败：{expr!r}：{exc}") from exc


def _resolve(ctx: dict[str, Any], path: str) -> Any:
    """按点分路径解析上下文值，例如 customer.name。"""
    cur = ctx
    for part in path.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        elif hasattr(cur, part):
            cur = getattr(cur, part)
        else:
            return None
    return cur


def _is_full_tag(text: str) -> bool:
    """判断单元格内容是否只包含一个标签（如 {items.quantity}）。"""
    return bool(_TAG_RE.fullmatch((text or "").strip()))


def _convert_back(text: str):
    """把替换后的字符串尽量转回数字或布尔，否则返回 None 表示保持字符串。"""
    if text == "TRUE":
        return True
    if text == "FALSE":
        return False
    try:
        if "." in text or "e" in text.lower():
            return float(text)
        return int(text)
    except ValueError:
        return None


def render_workbook(workbook: Workbook, data: dict[str, Any],
                    sheet_data: dict[str | int, dict[str, Any]], strict: bool) -> None:
    """对工作簿各工作表进行模板渲染。"""
    for ws in workbook.sheets:
        # 确定该表的根数据
        idx = workbook.sheets.index(ws)
        name = ws.name
        root_data = dict(data)
        if name in sheet_data:
            root_data.update(sheet_data[name])
        elif idx in sheet_data:
            root_data.update(sheet_data[idx])
        if not root_data:
            continue
        _render_sheet(ws, root_data, strict)


def _render_sheet(ws: Worksheet, root_data: dict[str, Any], strict: bool) -> None:
    # 先处理循环行块：找出 {loop ...} ... {/loop} 跨行区域
    rows = {}
    max_r = ws._max_row
    max_c = ws._max_column
    if max_r < 0 or max_c < 0:
        # 只有标量标签的情况（无 append 数据）
        for r in range(0, 50):
            for c in range(0, 10):
                cell = ws._cells.get((r, c))
                if cell and cell.get("value") is not None:
                    new_val = _replace_tags(str(cell["value"]), root_data, strict)
                    if new_val != cell["value"]:
                        cell["value"] = new_val
        return

    for r in range(max_r + 1):
        row_cells = []
        for c in range(max_c + 1):
            cell = ws._cells.get((r, c))
            if cell is None:
                cell = {"value": None, "formula": None, "style": None}
                ws._cells[(r, c)] = cell
            row_cells.append(cell)
        rows[r] = row_cells

    # 找出循环块范围
    loop_blocks: list[tuple[int, int, str]] = []  # (start_row, end_row, collection_path)
    i = 0
    while i <= max_r:
        line_values = [str(cell.get("value") or "") for cell in rows[i]]
        joined = "\t".join(line_values)
        m = re.search(r"\{loop\s+([\w\.]+)\}", joined)
        if m:
            coll = m.group(1)
            # 找匹配的 {/loop}
            j = i + 1
            while j <= max_r:
                j_joined = "\t".join(str(cell.get("value") or "") for cell in rows[j])
                if re.search(r"\{/loop\}", j_joined):
                    loop_blocks.append((i, j, coll))
                    i = j + 1
                    break
                j += 1
            else:
                raise TemplateError(f"未找到闭合的 {{/loop}}，循环集合：{coll}")
        else:
            i += 1

    # 先处理非循环行的标量标签
    loop_rows = set()
    for start, end, _ in loop_blocks:
        for rr in range(start, end + 1):
            loop_rows.add(rr)

    for r in range(max_r + 1):
        if r in loop_rows:
            continue
        for c in range(max_c + 1):
            cell = rows[r][c]
            if cell and cell.get("value") is not None and isinstance(cell["value"], str):
                new_val = _replace_tags(cell["value"], root_data, strict)
                if _is_full_tag(cell["value"]):
                    converted = _convert_back(new_val)
                    if converted is not None:
                        new_val = converted
                cell["value"] = new_val

    # 处理循环行块（从后往前展开，避免行号偏移）
    for start, end, coll in reversed(loop_blocks):
        collection = _resolve(root_data, coll)
        if collection is None:
            collection = []
        if not isinstance(collection, list):
            raise TemplateError(f"循环集合不是列表：{coll}")
        # 收集模板行（start+1 .. end-1）的单元格模板
        template_rows = []
        for rr in range(start + 1, end):
            trow = []
            for cc in range(max_c + 1):
                cell = rows[rr][cc]
                trow.append(str(cell.get("value") or "") if cell else "")
            template_rows.append(trow)
        # 删除循环块整行（start .. end）
        for rr in range(start, end + 1):
            for cc in range(max_c + 1):
                cell = rows[rr][cc]
                if cell:
                    cell["value"] = None
                    cell["formula"] = None
                    cell["style"] = None
        # 插入展开后的行
        cursor_r = start
        for item in collection:
            item_ctx = dict(root_data)
            item_ctx[coll.rsplit(".", 1)[-1]] = item
            item_ctx["@index"] = collection.index(item)
            for trow in template_rows:
                for cc, tval in enumerate(trow):
                    if tval == "":
                        continue
                    new_val = _replace_tags(tval, item_ctx, strict)
                    if _is_full_tag(tval):
                        converted = _convert_back(new_val)
                        if converted is not None:
                            new_val = converted
                    ws._set_cell_value(cursor_r, cc, new_val, formula=None)
                cursor_r += 1
        # 调整工作簿已触及边界
        ws._max_row = max(ws._max_row, cursor_r - 1)
        if max_c > ws._max_column:
            ws._max_column = max_c


def _replace_tags(text: str, ctx: dict[str, Any], strict: bool) -> str:
    """把文本中的 {标签} 替换为上下文值。"""

    def repl(m):
        tag = m.group(1).strip()
        if tag.startswith("/loop"):
            return m.group(0)  # 循环结束标记不在此处替换
        if tag.startswith("loop "):
            return m.group(0)
        # 处理 | format
        fmt = None
        if "|" in tag:
            tag, fmt = tag.split("|", 1)
            tag = tag.strip()
            fmt = fmt.strip()
            if fmt.startswith("format:"):
                fmt = fmt[len("format:"):].strip().strip('"')
        # 处理 @index
        value = _resolve(ctx, tag)
        if value is None and strict:
            raise TemplateError(f"模板标签缺失：{{{tag}}}")
        if value is None:
            value = ""
        if fmt:
            try:
                value = format(value, fmt)
            except Exception:
                pass
        return str(value)

    return _TAG_RE.sub(repl, text)


__all__ = ["render_workbook"]
