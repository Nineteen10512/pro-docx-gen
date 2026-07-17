"""DOCX templates — 包入口，统一注册 18 套白底模板（v1.5.1 D 能力域，v1.6.6 精简到 16，v1.7.0 扩到 18）。

约定：
- 模板数据文件是**纯数据**（thin），< 30 行，不写逻辑
- import 即注册（无运行时开销）
- 旧 API（academic.new_paper / business.new_report / teaching.new_lesson_plan）继续工作
- 不修改 pro_ppt_gen 任何文件
- v1.6.6 精简：删除 brand_luxury / proposal_elegant / data_analysis_tech / tech_whitepaper
  四个黑金/深色科技模板（heading 颜色违规），保留 16 套白底模板。
- v1.7.0 新增：business_compact（商务紧凑）+ research_report（研报专用）
  两套白底模板，仍走 _heading_rgb() 强制覆盖，self_audit 必过。

@since v1.5.1
@updated v1.6.6 (template cleanup)
@updated v1.7.0 (research/compact)
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

# 7 套 v1.5.1 新增模板（纯数据）
from . import thesis_full
from . import lesson_plan_gaokao
from . import business_proposal
from . import meeting_minutes
from . import resume_cn
from . import contract_cn
from . import reading_notes

# 6 套 v1.6.0 variant 白底模板（4 套黑金/深色科技模板已删除）
from . import academic_corporate
from . import business_minimal
from . import education_formal
from . import marketing_bold
from . import product_launch_bold
from . import report_minimal

# 2 套 v1.7.0 新增模板
from . import business_compact
from . import research_report

__all__ = [
    "DOCXTEMPLATE_REGISTRY",
    "register",
    "list_templates",
    "get_template",
]
