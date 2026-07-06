"""Collaboration package — DOCX 多人协作编辑（v1.5.3 C 能力域）。

提供：
- session.py       : CollaborationSession + start_collaboration()
- collaborators.py : Collaborator 描述 + add_collaborator() 自动分配颜色
- suggest.py       : suggest_insert / suggest_replace / suggest_delete
- comments.py      : add_review_comment / reply_comment（带 thread）
- merge.py         : list/accept/reject 策略 + merge_collaboration + export_clean_copy
- diff.py          : diff_documents 双版本对比

设计原则：
- 全 optional，不开新行为时与 v1.5 字节级兼容
- 协作者颜色用 w:color w:val="auto"，由应用自动分配
- 全部基于 lxml 构造 OOXML 节点（python-docx + lxml）

@since v1.5.3
"""
