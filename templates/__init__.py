"""DOCX templates — 包入口，统一注册 10 套模板（v1.5.3 D 能力域）。

约定：
- 模板数据文件是**纯数据**（thin），< 30 行，不写逻辑
- import 即注册（无运行时开销）
- 旧 API（academic.new_paper / business.new_report / teaching.new_lesson_plan）继续工作
- 不修改 pro_ppt_gen 任何文件

@since v1.5.3
"""
from .registry import (
    DOCXTEMPLATE_REGISTRY,
    register,
    list_templates,
    get_template,
)

# 3 套 v1.5 已有模板（迁移为 DOCXTemplate 描述符）
from . import academic_v151
from . import business_v151
from . import teaching_v151

# 7 套 v1.5.3 新增模板（纯数据）
from . import thesis_full
from . import lesson_plan_gaokao
from . import business_proposal
from . import meeting_minutes
from . import resume_cn
from . import contract_cn
from . import reading_notes

__all__ = [
    "DOCXTEMPLATE_REGISTRY",
    "register",
    "list_templates",
    "get_template",
]
