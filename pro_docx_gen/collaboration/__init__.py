"""Collaboration package — DOCX 多人协作编辑（v1.5.1 C 能力域）。

提供：
- session.py       : CollaborationSession + start_collaboration + add_collaborator
- collaborators.py : Collaborator 描述
- suggest.py       : suggest_insert / suggest_replace / suggest_delete
- comments.py      : add_review_comment / reply_comment / list_comments
- merge.py         : list_suggestions/accept_by_id/reject_by_id/accept_by_author/reject_by_author + merge_collaboration + export_clean_copy + export_review_pdf
- diff.py          : diff_documents 双版本对比

设计原则：
- 全 optional，不开新行为时与 v1.5 字节级兼容
- 协作者颜色用 w:color w:val="auto"，由应用自动分配
- 全部基于 lxml 构造 OOXML 节点（python-docx + lxml）

@since v1.5.1
"""

from .session import CollaborationSession, start_collaboration, add_collaborator, list_collaborators
from .collaborators import Collaborator
from .suggest import suggest_insert, suggest_replace, suggest_delete
from .comments import add_review_comment, reply_comment, list_comments
from .merge import (
    list_suggestions, accept_by_id, reject_by_id,
    accept_by_author, reject_by_author,
    merge_collaboration, export_clean_copy, export_review_pdf,
)
from .diff import diff_documents

__all__ = [
    "CollaborationSession", "start_collaboration",
    "add_collaborator", "list_collaborators",
    "Collaborator",
    "suggest_insert", "suggest_replace", "suggest_delete",
    "add_review_comment", "reply_comment", "list_comments",
    "list_suggestions", "accept_by_id", "reject_by_id",
    "accept_by_author", "reject_by_author",
    "merge_collaboration", "export_clean_copy", "export_review_pdf",
    "diff_documents",
]
