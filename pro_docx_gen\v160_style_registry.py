"""v1.6.0 Style Variant Registry — DOCX 端 5 变体元数据。

与 pro-ppt-gen 的 v160_style_registry.py 保持结构一致，共用同一套 VariantProfile。
DOCX 端通过 variant_tokens.py 将 VariantProfile 映射为具体 token 覆盖。

@since v1.6.0
"""
from dataclasses import dataclass, field


@dataclass
class VariantProfile:
    name: str
    display_name: str
    description: str
    # 视觉特征
    colors: dict = field(default_factory=dict)
    font_style: str = ""          # "sans" / "serif" / "geometric" / "humanist"
    spacing: str = ""             # "spacious" / "normal" / "compact"
    decoration: str = ""          # "minimal" / "geometric" / "organic" / "bold" / "luxury"
    # 匹配关键词
    scene_keywords: list = field(default_factory=list)
    # 适用场景
    suitable_scenes: list = field(default_factory=list)


VARIANT_PROFILES: dict[str, VariantProfile] = {
    "corporate_formal": VariantProfile(
        name="corporate_formal",
        display_name="企业正式",
        description="蓝白/灰白标准企业色，网格对齐，信息密度适中",
        colors={"primary": "深蓝", "accent": "标准蓝", "bg": "白/浅灰"},
        font_style="sans",
        spacing="normal",
        decoration="minimal",
        scene_keywords=["汇报", "报告", "年度", "总结", "述职", "董事会", "政府", "公文", "正式"],
        suitable_scenes=["business_report", "academic", "data_analysis"],
    ),
    "minimal_clean": VariantProfile(
        name="minimal_clean",
        display_name="极简干净",
        description="单色/双色低饱和，大留白，无衬线细体，Apple Keynote 风格",
        colors={"primary": "黑白", "accent": "一抹蓝/灰", "bg": "大量白"},
        font_style="sans",
        spacing="spacious",
        decoration="minimal",
        scene_keywords=["极简", "简约", "干净", "留白", "苹果", "咨询", "设计"],
        suitable_scenes=["business_report", "creative_showcase", "product_launch", "brand_story"],
    ),
    "modern_tech": VariantProfile(
        name="modern_tech",
        display_name="现代科技",
        description="深色模式或渐变蓝紫，发光强调，几何无衬线",
        colors={"primary": "深蓝紫", "accent": "青蓝霓虹", "bg": "深色/渐变"},
        font_style="geometric",
        spacing="normal",
        decoration="geometric",
        scene_keywords=["科技", "AI", "技术", "互联网", "SaaS", "数字化", "智能", "编程"],
        suitable_scenes=["product_launch", "data_analysis", "creative_showcase"],
    ),
    "bold_impact": VariantProfile(
        name="bold_impact",
        display_name="大胆冲击",
        description="强对比黑白+单强调色，超大标题粗体，适合TED/发布会",
        colors={"primary": "黑", "accent": "单一强调色", "bg": "强对比"},
        font_style="sans",
        spacing="compact",
        decoration="bold",
        scene_keywords=["冲击", "大胆", "TED", "发布会", "创意", "震撼", "视觉", "海报"],
        suitable_scenes=["product_launch", "creative_showcase", "brand_story", "marketing"],
    ),
    "elegant_luxury": VariantProfile(
        name="elegant_luxury",
        display_name="优雅奢华",
        description="黑金/白金/深蓝金，衬线标题细体，对称细边框",
        colors={"primary": "黑金/白金", "accent": "金色", "bg": "深色/米色"},
        font_style="serif",
        spacing="spacious",
        decoration="luxury",
        scene_keywords=["奢华", "高端", "奢侈", "品牌", "珠宝", "地产", "酒店", "精品"],
        suitable_scenes=["brand_story", "product_launch", "marketing"],
    ),
}


def get_variant(name: str) -> VariantProfile | None:
    return VARIANT_PROFILES.get(name)


def list_variants() -> list[dict]:
    return [
        {
            "name": v.name,
            "display_name": v.display_name,
            "description": v.description,
            "suitable_scenes": v.suitable_scenes,
        }
        for v in VARIANT_PROFILES.values()
    ]


def match_variant(scene: str = "", keywords: str = "") -> str:
    """Auto-match best variant from scene + content keywords.

    Returns variant name string. Default: corporate_formal.
    """
    scores: dict[str, int] = {}
    for name, v in VARIANT_PROFILES.items():
        s = 0
        if scene in v.suitable_scenes:
            s += 3
        if keywords:
            for kw in v.scene_keywords:
                if kw in keywords:
                    s += 2
        scores[name] = s

    if not scores:
        return "corporate_formal"

    best = max(scores, key=scores.get)
    if scores[best] == 0:
        return "corporate_formal"
    return best