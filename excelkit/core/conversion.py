"""单元格写回转换与只读转换共用的唯一类型转换实现。"""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from typing import Any

_DATE_TEXT_PATTERN = re.compile(
    r"^#?(\d{4})([-/])(\d{1,2})\2(\d{1,2})"
    r"(?: (\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$"
)


def _temporal(value: str) -> date | datetime:
    """功能：把支持的日期或日期时间字符串解析为 Python 时间对象。

    使用方法：由日期自动识别、``as_date()`` 和 ``as_datetime()`` 共用。
    参数：``value`` 为可带 ``#`` 前缀、使用 ``-`` 或 ``/`` 分隔的字符串，
    时间部分可为 ``H:M`` 或 ``H:M:S``。
    返回：没有时间部分时返回 :class:`date`，否则返回 :class:`datetime`。
    异常：格式或日期时间取值无效时抛出 ``ValueError``。
    """
    match = _DATE_TEXT_PATTERN.fullmatch(value.strip())
    if match is None:
        raise ValueError(
            "日期时间字符串必须使用 YYYY-M-D、YYYY/M/D 或追加 H:M[:S]"
        )
    try:
        year, month, day = (int(match.group(index)) for index in (1, 3, 4))
        if match.group(5) is None:
            return date(year, month, day)
        return datetime(
            year,
            month,
            day,
            int(match.group(5)),
            int(match.group(6)),
            int(match.group(7) or 0),
        )
    except ValueError as error:
        raise ValueError(f"无效的日期时间值：{value!r}") from error


def normalize_value(value: Any) -> Any:
    """功能：自动识别带 ``#`` 前缀的日期及日期时间普通值。

    使用方法：工作表所有普通值写入入口统一调用。
    参数：``value`` 为任意对象；仅严格匹配 ``#YYYY-M-D``、``#YYYY/M/D``
    及其可选时间形式的字符串参与转换。
    返回：匹配时返回 ``date`` 或 ``datetime``，否则返回原值。
    异常：匹配日期格式但取值无效时抛出 ``ValueError``。
    """
    if not isinstance(value, str) or not value.startswith("#"):
        return value
    try:
        return _temporal(value)
    except ValueError:
        if _DATE_TEXT_PATTERN.fullmatch(value.strip()) is not None:
            raise
        return value


def as_string(value: Any) -> str:
    """功能：把任意单元格值转换为字符串。

    使用方法：由写回和只读的 ``as_string()`` 共用。
    参数：``value`` 为任意对象，``None`` 转换为空字符串。
    返回：转换后的 ``str``。
    异常：对象自身的 ``str()`` 实现失败时透传相应异常。
    """
    return "" if value is None else str(value)


def as_int(value: Any) -> int:
    """功能：把值无损转换为整数。

    使用方法：由写回和只读的 ``as_int()`` 共用。
    参数：``value`` 可为整数、整数字符串或没有小数部分的有限浮点数；布尔值
    不作为整数接受。
    返回：转换后的 ``int``。
    异常：类型不支持时抛出 ``TypeError``；格式无效或会损失小数时抛出
    ``ValueError``。
    """
    if isinstance(value, bool):
        raise TypeError("as_int() 不接受布尔值")
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            raise ValueError("as_int() 只能转换没有小数部分的有限浮点数")
        return int(value)
    if isinstance(value, str):
        try:
            return int(value.strip(), 10)
        except ValueError as error:
            raise ValueError("as_int() 需要合法的整数字符串") from error
    raise TypeError("as_int() 只能转换字符串、整数或浮点数")


def as_float(value: Any) -> float:
    """功能：把值转换为有限浮点数。

    使用方法：由写回和只读的 ``as_float()`` 共用。
    参数：``value`` 可为整数、浮点数或合法浮点数字符串；布尔值不接受。
    返回：转换后的有限 ``float``。
    异常：类型不支持时抛出 ``TypeError``；格式无效、无穷大或 NaN 时抛出
    ``ValueError``。
    """
    if isinstance(value, bool):
        raise TypeError("as_float() 不接受布尔值")
    if not isinstance(value, (str, int, float)):
        raise TypeError("as_float() 只能转换字符串、整数或浮点数")
    try:
        converted = float(value.strip() if isinstance(value, str) else value)
    except ValueError as error:
        raise ValueError("as_float() 需要合法的浮点数字符串") from error
    if not math.isfinite(converted):
        raise ValueError("as_float() 不接受无穷大或 NaN")
    return converted


def as_bool(value: Any) -> bool:
    """功能：按明确白名单把值转换为布尔值。

    使用方法：由写回和只读的 ``as_bool()`` 共用。
    参数：``value`` 可为布尔值、数值 1/0，或 ``true/false``、``yes/no``、
    ``是/否``、``1/0`` 文本；英文忽略大小写和首尾空白。
    返回：转换后的 ``bool``。
    异常：类型不支持时抛出 ``TypeError``；值不在白名单时抛出 ``ValueError``。
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("as_bool() 数值只接受 1 和 0")
        if value not in (0, 1):
            raise ValueError("as_bool() 数值只接受 1 和 0")
        return bool(value)
    if isinstance(value, str):
        normalized = value.strip().casefold()
        if normalized in {"true", "yes", "是", "1"}:
            return True
        if normalized in {"false", "no", "否", "0"}:
            return False
        raise ValueError("as_bool() 文本不在受支持的真假值白名单中")
    raise TypeError("as_bool() 只能转换字符串、布尔值或 0/1 数值")


def as_date(value: Any) -> date:
    """功能：把值转换为日期并在需要时舍弃时间部分。

    使用方法：由写回和只读的 ``as_date()`` 共用。
    参数：``value`` 可为支持格式的字符串、``date`` 或 ``datetime``。
    返回：转换后的 :class:`date`；``datetime`` 只保留年月日。
    异常：类型不支持时抛出 ``TypeError``；字符串格式或取值无效时抛出
    ``ValueError``。
    """
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        converted = _temporal(value)
        return converted.date() if isinstance(converted, datetime) else converted
    raise TypeError("as_date() 只能转换字符串、date 或 datetime 值")


def as_datetime(value: Any) -> datetime:
    """功能：把值转换为日期时间并为纯日期补充午夜时间。

    使用方法：由写回和只读的 ``as_datetime()`` 共用。
    参数：``value`` 可为支持格式的字符串、``date`` 或 ``datetime``。
    返回：转换后的 :class:`datetime`；纯日期的时间为 ``00:00:00``。
    异常：类型不支持时抛出 ``TypeError``；字符串格式或取值无效时抛出
    ``ValueError``。
    """
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        converted = _temporal(value)
        if isinstance(converted, datetime):
            return converted
        return datetime(converted.year, converted.month, converted.day)
    raise TypeError("as_datetime() 只能转换字符串、date 或 datetime 值")


__all__ = [
    "normalize_value",
    "as_string",
    "as_int",
    "as_float",
    "as_bool",
    "as_date",
    "as_datetime",
]
