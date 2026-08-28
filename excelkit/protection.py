"""工作簿和工作表保护设置。"""

from __future__ import annotations


class Protection:
    """表示可序列化的基础保护选项。"""

    __slots__ = ("enabled", "password", "select_locked", "select_unlocked")

    def __init__(self) -> None:
        """功能：创建未启用的保护设置；参数无；返回无。"""
        self.enabled = False
        self.password = None
        self.select_locked = True
        self.select_unlocked = True


__all__ = ["Protection"]
