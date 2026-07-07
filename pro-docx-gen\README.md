# PRO-PPTX v1.5.2

PaperJSX 语义编译架构 — 专业 PPT 生成技能（独立包）。

## v1.5.2 更新内容（2026-07）

### 🛡️ 设计铁律（新增）

1. **对比度铁律（深色区域禁深色字 / 浅色区域禁浅色字）**
   - `shared/quality.py` 新增 `contrast_check()`：批量校验 (bg, fg) 对是否满足 WCAG AA 正文 ≥4.5:1 对比度
   - `pro_ppt_gen/taste.py` 新增 `_check_contrast()`：语义层对标题栏/章节页/KPI/正文四类场景预检，不达标发 warning
   - `pro_ppt_gen/engine/renderer.py` 新增 `_ensure_readable()` 兜底：渲染前自动把不达标颜色反色为黑/白（哪个与底色对比强用哪个），即使上游 LLM 传错色也保证文字看得见
   - textbox/bullets 文本框默认 `fill.background()`（noFill）透明，避免白底盖在深色标题栏上造成视觉白对白

2. **图表多色铁律（禁止全同色）**
   - `pro_ppt_gen/taste.py` 新增 `_check_chart_color_variety()`：多 series 被显式指定同一颜色时告警
   - `pro_ppt_gen/engine/chart_renderer.py` 重构 series 上色逻辑：
     - 单系列柱/条图：每根柱子按调色板循环上色（per-dPt）
     - 单系列饼/环图：每块扇形按调色板循环上色
     - 多系列图：每个 series 分配调色板独立颜色
   - `shared/themes.py` `THEME_CHART_PALETTES` 从每主题 4 色扩充到 6 色，且全部对各自主题底色做过 WCAG 对比度校验
   - `pro_ppt_gen/tokens/themes.py` / `pro_docx_gen/tokens/themes.py` 同步扩充 chart_palette_override 至 6 色
   - `pro_docx_gen/engine/chart_renderer.py` 同步实现柱/条图单系列多彩色

### 🔧 其他改进
- `taste_check()` 支持 `dict` 类型 theme（`{"name":..., "overrides":{...}}`），dict theme override 能正确被对比度预检识别
- 版本号升级到 1.5.2
- 新增 `smoke_tests/test_v152_rules.py` 专门验证三条新铁律
- 向后兼容：所有 v1.5.x 语义 JSON 可直接生成，无新增强制依赖

## 使用

```python
from pro_ppt_gen import generate
generate(content, "output.pptx", theme="business")
```

详见 SKILL.md。
