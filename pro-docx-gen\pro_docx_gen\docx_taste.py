"""v1.6.0 DOCX 审美预检 — taste_check() for DOCX documents.

基于 semantic JSON 做文档质量检查，不渲染文件：
- 段落密度检查（过密/过疏）
- 标题层级一致性
- 表格/图题一致性
- 场景-变体匹配度
- AI 套话检测

@since v1.6.0
"""
from __future__ import annotations

import re
from typing import Any, Optional


# ─── AI 套话正则 ──────────────────────────────────────────────

_AI_TELL_PATTERNS = [
    (re.compile(r"众所周知"), "AI 套话：'众所周知' 建议替换为具体数据引用"),
    (re.compile(r"赋能"), "AI 套话：'赋能' 过于泛化，建议具体说明"),
    (re.compile(r"颠覆(?!性)"), "AI 套话：'颠覆' 建议替换为具体成果描述"),
    (re.compile(r"打通.*闭环"), "AI 套话：'打通闭环' 建议拆解为具体步骤"),
    (re.compile(r"抓手"), "AI 套话：'抓手' 建议替换为具体方法"),
    (re.compile(r"对齐.*颗粒度"), "AI 套话：'对齐颗粒度' 建议替换为具体标准"),
    (re.compile(r"底层逻辑"), "AI 套话：'底层逻辑' 建议替换为具体原理"),
    (re.compile(r"降维打击"), "AI 套话：'降维打击' 建议替换为具体竞争策略"),
    (re.compile(r"倒逼"), "AI 套话：'倒逼' 建议替换为'推动/促使'"),
    (re.compile(r"组合拳"), "AI 套话：'组合拳' 建议替换为'综合方案'"),
    (re.compile(r"护城河"), "AI 套话：'护城河' 建议替换为'核心竞争力'"),
    (re.compile(r"引爆"), "AI 套话：'引爆' 建议替换为'快速推广'"),
    (re.compile(r"破圈"), "AI 套话：'破圈' 过于口语化，建议替换"),
    (re.compile(r"在.*的今天"), "AI 套话：'在……的今天' 句式建议替换为直接陈述"),
    (re.compile(r"随着.*的.*发展"), "AI 套话：'随着……的发展' 句式建议替换为具体时间节点"),
]

# ─── 场景-变体兼容性矩阵 ──────────────────────────────────────

_SCENE_VARIANT_MAP = {
    "business_report": ["corporate_formal", "minimal_clean", "bold_impact"],
    "product_launch": ["modern_tech", "bold_impact", "minimal_clean", "elegant_luxury"],
    "education": ["corporate_formal", "minimal_clean"],
    "brand_story": ["elegant_luxury", "minimal_clean", "bold_impact"],
    "data_analysis": ["corporate_formal", "modern_tech", "minimal_clean"],
    "creative_showcase": ["bold_impact", "modern_tech", "minimal_clean"],
    "academic": ["corporate_formal", "minimal_clean"],
    "marketing": ["bold_impact", "elegant_luxury", "modern_tech", "minimal_clean"],
}


def taste_check(
    content: dict,
    variant: Optional[str] = None,
    scene: Optional[str] = None,
    lang: str = "cn",
    strict: bool = False,
) -> dict:
    """对 DOCX semantic JSON 做审美/质量预检。

    Args:
        content: 语义 JSON 字典（与 generate() 相同格式）
        variant: 当前使用的风格变体（用于场景匹配检查）
        scene: 文档场景（business_report/education/...）
        lang: 语言 "cn" | "en"
        strict: 严格模式，阈值提高到 85

    Returns:
        {
            "score": int (0-100),
            "passed": bool,
            "issues": [{"level": "ERROR"|"WARNING"|"NOTICE", "code": str, "message": str}],
            "details": {...},
        }
    """
    issues: list[dict] = []
    score = 100
    details: dict[str, Any] = {}

    # 1) 段落密度检查
    _check_paragraph_density(content, issues, details)

    # 2) 标题层级一致性
    _check_heading_consistency(content, issues, details)

    # 3) 表格/图题一致性
    _check_table_caption_consistency(content, issues, details)

    # 4) 场景-变体匹配
    if variant and scene:
        _check_scene_variant_match(scene, variant, issues)

    # 5) AI 套话检测
    _check_ai_tells(content, issues, lang)

    # 6) 空内容检测
    _check_empty_content(content, issues)

    # 计算扣分
    for issue in issues:
        if issue["level"] == "ERROR":
            score -= 10
        elif issue["level"] == "WARNING":
            score -= 5
        elif issue["level"] == "NOTICE":
            score -= 2

    score = max(0, min(100, score))
    threshold = 85 if strict else 75
    passed = score >= threshold

    details["score"] = score
    details["passed"] = passed
    details["issue_count"] = len(issues)

    return {
        "score": score,
        "passed": passed,
        "issues": issues,
        "details": details,
    }


def _check_paragraph_density(content: dict, issues: list, details: dict):
    """检查段落密度：过密（>500字/paragraph）或过疏（<10字/paragraph）。"""
    sections = content.get("sections", [])
    total_paras = 0
    dense_paras = 0
    sparse_paras = 0

    for sec in sections:
        for node in sec.get("content", []):
            if node.get("type") == "paragraph":
                text = node.get("text", "")
                total_paras += 1
                if len(text) > 500:
                    dense_paras += 1
                elif len(text) < 10 and text.strip():
                    sparse_paras += 1

    if dense_paras > 0:
        issues.append({
            "level": "WARNING",
            "code": "dense_paragraph",
            "message": f"{dense_paras} 个段落超过 500 字，建议拆分或精简",
        })
    if sparse_paras > 3:
        issues.append({
            "level": "NOTICE",
            "code": "sparse_paragraph",
            "message": f"{sparse_paras} 个段落少于 10 字，建议合并或补充",
        })

    details["paragraph_count"] = total_paras
    details["dense_paragraphs"] = dense_paras
    details["sparse_paragraphs"] = sparse_paras


def _check_heading_consistency(content: dict, issues: list, details: dict):
    """检查标题层级：不能用 h3 跳过 h2，或 h1 后直接 h3。"""
    sections = content.get("sections", [])
    levels = [sec.get("level", 1) for sec in sections]
    if not levels:
        return

    for i, (prev, curr) in enumerate(zip(levels, levels[1:]), 1):
        if curr - prev > 1:
            issues.append({
                "level": "WARNING",
                "code": "heading_skip",
                "message": f"标题层级跳跃：第 {i} 节 level={prev} → 第 {i+1} 节 level={curr}，建议不要跳过中间层级",
            })
            break

    # 检查是否有重复标题
    titles = [sec.get("title", "") for sec in sections]
    seen = set()
    for i, t in enumerate(titles):
        if t and t in seen:
            issues.append({
                "level": "NOTICE",
                "code": "duplicate_heading",
                "message": f"重复标题：第 {i+1} 节 '{t}' 与前面重复",
            })
        seen.add(t)

    details["heading_count"] = len(levels)


def _check_table_caption_consistency(content: dict, issues: list, details: dict):
    """检查表格/图表是否有 caption 或 title。"""
    sections = content.get("sections", [])
    tables_without_caption = 0
    figures_without_caption = 0

    for sec in sections:
        for node in sec.get("content", []):
            if node.get("type") == "table":
                if not node.get("caption"):
                    tables_without_caption += 1
            elif node.get("type") == "figure":
                if not node.get("caption"):
                    figures_without_caption += 1

    if tables_without_caption > 0:
        issues.append({
            "level": "NOTICE",
            "code": "table_no_caption",
            "message": f"{tables_without_caption} 个表格缺少 caption，建议补充题注",
        })
    if figures_without_caption > 0:
        issues.append({
            "level": "NOTICE",
            "code": "figure_no_caption",
            "message": f"{figures_without_caption} 个图表缺少 caption，建议补充题注",
        })

    details["tables_no_caption"] = tables_without_caption
    details["figures_no_caption"] = figures_without_caption


def _check_scene_variant_match(scene: str, variant: str, issues: list):
    """检查场景-变体兼容性。"""
    compatible = _SCENE_VARIANT_MAP.get(scene, [])
    if compatible and variant not in compatible:
        issues.append({
            "level": "WARNING",
            "code": "scene_variant_mismatch",
            "message": f"场景 '{scene}' 与变体 '{variant}' 不匹配。建议：{compatible}",
        })


def _check_ai_tells(content: dict, issues: list, lang: str):
    """检测 AI 套话/模板味文案。"""
    if lang != "cn":
        return

    sections = content.get("sections", [])
    all_text_parts = []

    # 收集 meta
    meta = content.get("meta", {})
    for key in ("title", "subtitle", "abstract"):
        val = meta.get(key, "")
        if isinstance(val, str):
            all_text_parts.append(val)

    # 收集 sections
    for sec in sections:
        all_text_parts.append(sec.get("title", ""))
        for node in sec.get("content", []):
            for key in ("text", "body", "caption", "title"):
                val = node.get(key, "")
                if isinstance(val, str):
                    all_text_parts.append(val)
            # bullets
            for bullet in node.get("bullets", []) or node.get("items", []) or []:
                if isinstance(bullet, str):
                    all_text_parts.append(bullet)
                elif isinstance(bullet, dict):
                    all_text_parts.append(bullet.get("text", ""))

    full_text = " ".join(all_text_parts)

    found = []
    for pattern, msg in _AI_TELL_PATTERNS:
        if pattern.search(full_text):
            found.append(msg)
            if len(found) >= 5:
                break

    if found:
        issues.append({
            "level": "NOTICE",
            "code": "ai_copy_detected",
            "message": f"检测到 {len(found)} 处 AI 套话：{'; '.join(found[:3])}",
        })


def _check_empty_content(content: dict, issues: list):
    """检查是否有空章节。"""
    sections = content.get("sections", [])
    empty_sections = []
    for i, sec in enumerate(sections, 1):
        if not sec.get("content"):
            empty_sections.append(i)

    if empty_sections:
        issues.append({
            "level": "WARNING",
            "code": "empty_section",
            "message": f"第 {', '.join(map(str, empty_sections))} 节内容为空，建议填充或删除",
        })