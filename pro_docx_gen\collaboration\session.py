"""CollaborationSession — DOCX 协作会话管理（v1.5.1 C 能力域）。

提供：
- start_collaboration(docx_path) : 创建会话并开启 trackChanges
- add_collaborator(session, name) : 注册协作者（自动分配颜色）
- list_collaborators(session)     : 列出所有协作者
- session.apply_suggestion(...)   : 由 suggest.py 调用，记录到 session.suggestions
- session.export() / merge()

设计：
- session 状态是轻量 dict 描述，不直接持有 docx 锁
- 所有 modify 都在 docx 文件上做 w:ins/w:del
- trackChanges 在 settings.xml 开启（由 start_collaboration 写入）
"""
from __future__ import annotations

import json
import os
import shutil
from dataclasses import dataclass, field
from typing import Any, Optional

from .collaborators import Collaborator, _next_color_idx


@dataclass
class CollaborationSession:
    """协作会话状态。

    字段：
    - session_id    : 会话唯一 ID（uuid4 hex）
    - docx_path     : 当前 docx 路径
    - base_path     : 原始 docx 路径（用于 export_clean_copy 对比）
    - track_changes : 是否开启修订追踪
    - collaborators : 协作者列表（按加入顺序）
    - suggestions   : 建议记录列表（[{id, type, author, status, target, text, created_at}]）
    - comments      : 批注列表
    - created_at    : 会话创建时间
    """
    session_id: str
    docx_path: str
    base_path: str = ""
    track_changes: bool = True
    collaborators: list[Collaborator] = field(default_factory=list)
    suggestions: list[dict] = field(default_factory=list)
    comments: list[dict] = field(default_factory=list)
    created_at: str = ""

    def add_suggestion(self, sug: dict) -> None:
        self.suggestions.append(sug)

    def add_comment(self, cmt: dict) -> None:
        self.comments.append(cmt)

    def to_dict(self) -> dict[str, Any]:
        return {
            "session_id": self.session_id,
            "docx_path": self.docx_path,
            "base_path": self.base_path,
            "track_changes": self.track_changes,
            "collaborators": [
                {
                    "author": c.author,
                    "initials": c.initials,
                    "color_idx": c.color_idx,
                    "joined_at": c.joined_at,
                    "role": c.role,
                }
                for c in self.collaborators
            ],
            "suggestions": self.suggestions,
            "comments": self.comments,
            "created_at": self.created_at,
        }


def _enable_track_changes_in_docx(docx_path: str) -> None:
    """在 docx 的 settings.xml 写入 w:trackChanges 节点。

    - 不存在则追加
    - 已存在则不动
    - 解析失败抛清晰错误（不允许静默降级）
    """
    import zipfile
    from lxml import etree

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    nsmap = {"w": W}

    tmp_out = docx_path + ".track_tmp"
    settings_modified = False

    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/settings.xml":
                    try:
                        root = etree.fromstring(data)
                    except etree.XMLSyntaxError as e:
                        raise RuntimeError(f"settings.xml 解析失败: {e}") from e
                    # 查 w:trackChanges
                    existing = root.find(f"{{{W}}}trackChanges")
                    if existing is None:
                        track = etree.SubElement(root, f"{{{W}}}trackChanges")
                        # Word 默认属性 val=true；省略表示 true
                    settings_modified = True
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(item, data)
            if not settings_modified:
                # settings.xml 不存在，创建一个最小骨架
                settings_xml = (
                    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<w:settings xmlns:w="{W}">'
                    f'<w:trackChanges/>'
                    f'</w:settings>'
                ).encode("utf-8")
                zout.writestr("word/settings.xml", settings_xml)
    shutil.move(tmp_out, docx_path)


def start_collaboration(
    docx_path: str,
    track_changes: bool = True,
    session_id: Optional[str] = None,
) -> CollaborationSession:
    """创建协作会话并开启 trackChanges。

    Args:
        docx_path: 已存在的 docx 路径
        track_changes: 是否开启修订追踪（默认 True）
        session_id: 可选自定义会话 ID；不传则生成 uuid4

    Returns:
        CollaborationSession 实例

    Raises:
        FileNotFoundError: docx 不存在
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"找不到 docx 文件: {docx_path}")

    if track_changes:
        _enable_track_changes_in_docx(docx_path)

    if session_id is None:
        import uuid
        session_id = uuid.uuid4().hex[:12]

    from datetime import datetime
    sess = CollaborationSession(
        session_id=session_id,
        docx_path=docx_path,
        base_path=docx_path,
        track_changes=track_changes,
        created_at=datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
    )
    return sess


def add_collaborator(
    session: CollaborationSession,
    author: str,
    role: str = "editor",
) -> Collaborator:
    """注册协作者并自动分配颜色索引。

    Args:
        session: CollaborationSession
        author: 协作者显示名
        role: "editor" | "reviewer" | "commenter"

    Returns:
        Collaborator 实例
    """
    used = {c.color_idx for c in session.collaborators}
    color_idx = _next_color_idx(used)
    initials = author[:2] if author else "?"
    c = Collaborator(
        author=author,
        initials=initials,
        color_idx=color_idx,
        role=role,
    )
    session.collaborators.append(c)
    return c


def list_collaborators(session: CollaborationSession) -> list[dict]:
    """列出所有协作者（轻量字典）。"""
    return [
        {
            "author": c.author,
            "initials": c.initials,
            "color_idx": c.color_idx,
            "role": c.role,
            "joined_at": c.joined_at,
        }
        for c in session.collaborators
    ]


__all__ = [
    "CollaborationSession",
    "start_collaboration",
    "add_collaborator",
    "list_collaborators",
]
