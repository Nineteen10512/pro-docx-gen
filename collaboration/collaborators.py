"""Collaborator — DOCX 协作参与者描述（v1.5.3 C 能力域）。

每位协作者由 author 标识，自动分配一个 0~15 的颜色索引（用于 w:color w:val="auto"）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class Collaborator:
    """协作参与者。

    字段：
    - author      : 协作者显示名（Word 中显示）
    - initials    : 缩写（用于批注标记）
    - color_idx   : 自动分配的颜色索引（0~15），对应 Word 16 色
    - joined_at   : 加入会话时间
    - role        : "editor" | "reviewer" | "commenter"
    """
    author: str
    initials: str = ""
    color_idx: int = 0
    joined_at: str = ""
    role: str = "editor"

    def __post_init__(self):
        if not self.initials:
            # 自动从 author 取首字母（支持中文，取首字）
            self.initials = self.author[:2] if self.author else "?"
        if not self.joined_at:
            self.joined_at = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")


def _next_color_idx(used: set[int]) -> int:
    """分配下一个未使用的颜色索引（0~15）。"""
    for i in range(16):
        if i not in used:
            return i
    # 全部占用时回环
    return 0


__all__ = ["Collaborator", "_next_color_idx"]
