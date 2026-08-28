"""工作簿文档属性对象。"""

from __future__ import annotations

from datetime import datetime
from typing import Optional


class WorkbookProperties:
    """保存 XLSX 核心文档属性的可变对象。"""

    __slots__ = (
        "title", "subject", "author", "keywords", "comments", "category",
        "created", "modified", "last_modified_by",
    )

    def __init__(
        self,
        *,
        title: str = "",
        subject: str = "",
        author: str = "",
        keywords: str = "",
        comments: str = "",
        category: str = "",
        created: Optional[datetime] = None,
        modified: Optional[datetime] = None,
        last_modified_by: str = "",
    ) -> None:
        """功能：创建工作簿核心文档属性集合。

        使用方法：``workbook.properties.title = "销售报表"``；也可直接构造后传入
        ``WorkbookProperties(...)`` 的字段值。
        参数：文本字段必须是字符串；``created``、``modified`` 为可选
        ``datetime``；时间不带时区时按本地文档时间写出。
        返回：无；创建后字段可直接修改。
        异常：字段类型不正确时抛出 ``TypeError``。
        """
        for name, value in (
            ("title", title), ("subject", subject), ("author", author),
            ("keywords", keywords), ("comments", comments),
            ("category", category), ("last_modified_by", last_modified_by),
        ):
            if not isinstance(value, str):
                raise TypeError(f"{name} 必须是字符串")
            setattr(self, name, value)
        for name, value in (("created", created), ("modified", modified)):
            if value is not None and not isinstance(value, datetime):
                raise TypeError(f"{name} 必须是 datetime 或 None")
            setattr(self, name, value)


__all__ = ["WorkbookProperties"]
