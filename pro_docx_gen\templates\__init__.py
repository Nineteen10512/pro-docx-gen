"""DOCX templates package for PRO-DOCX v1.6.0."""

from .registry import (
    DOCXTEMPLATE_REGISTRY,
    register,
    list_templates,
    get_template,
)

from . import academic_v151
from . import business_v151
from . import teaching_v151

from . import thesis_full
from . import lesson_plan_gaokao
from . import business_proposal
from . import meeting_minutes
from . import resume_cn
from . import contract_cn
from . import reading_notes

from . import academic_corporate
from . import business_minimal
from . import tech_whitepaper
from . import education_formal
from . import brand_luxury
from . import marketing_bold
from . import data_analysis_tech
from . import product_launch_bold
from . import report_minimal
from . import proposal_elegant

__all__ = [
    "DOCXTEMPLATE_REGISTRY",
    "register",
    "list_templates",
    "get_template",
]
