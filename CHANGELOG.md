# PRO-DOCX 版本日志 (CHANGELOG)

> PaperJSX 语义编译架构 — 专业级 DOCX/Word 文档生成与翻译技能

---

## v1.6.0（2026-07-06）— AI 视觉审美引擎

### 🆕 结构保真翻译工作流
- 新增 `translation.py` 翻译模块
- `collect_translation_segments(docx_path)` — 收集可翻译文本段
- `apply_translation_map(docx_path, translations, output_path, target_lang, auto_format_tables)` — 翻译回写
- `assess_translation_risk(docx_path, translations)` — 翻译风险评估
- 默认保留原文档排版、图片、表格、页眉页脚、脚注/尾注引用
- 支持 Replace / Bilingual / Review 三种翻译模式

### 🆕 表格翻译层
- 表格内容可正常翻译，不再跳过
- 引擎自动评估文本膨胀风险
- 按需执行 mild shrink / autofit / 表格收紧策略

### 🆕 5 套全局风格变体
- `corporate_formal` — 商务正式
- `minimal_clean` — 极简清爽
- `modern_tech` — 科技现代
- `bold_impact` — 冲击力强
- `elegant_luxury` — 优雅奢华
- 支持 `variant` 和 `auto_style` 参数

### 🆕 DOCX 审美预检
- 新增 `docx_taste.py` — 生成前检查文档气质、密度与观感风险
- 段落密度、标题一致性、表格图题一致性、场景匹配度

### 🆕 图片保真修复
- 翻译回写保护图片、域代码、批注/脚注引用等非纯文本节点
- 文本回写到原段落与原单元格附近

### 🆕 自动风格决策
- 新增 `theme_extractor.py` — 主题自动提取
- 新增 `variant_tokens.py` — 变体设计令牌
- 新增 `v160_style_registry.py` — v1.6.0 风格注册表
- `auto_style=True` 按内容自动推荐风格

### 🔧 稳定性
- 默认 smoke suite 新增 `DOCX image preservation` + `image regression`
- 100% 向后兼容 v1.5.x JSON / Markdown 输入
- 不新增强制依赖

---

## v1.5.1（2026-07）

### 🆕 10 套场景化模板
- `lesson_plan_gaokao` 高考教案 / `thesis_full` 完整论文
- `business_proposal` 商业提案 / `meeting_minutes` 会议纪要
- `resume_cn` 中文简历 / `contract_cn` 中文合同
- `reading_notes` 读书笔记 / `brand_luxury` 品牌奢华
- `data_analysis_tech` 数据分析 / `tech_whitepaper` 技术白皮书
- 覆盖原 academic / business / teaching 三套

### 🆕 本地模板主题提取
- `template_path` 参数支持 .docx/.dotx/.wpt 文件
- `scan_local_templates()` 扫描默认模板目录
- `auto_taste_match=True` 自动 WCAG AA 对比度校验

### 🆕 多人协作能力
- `collaboration/` 模块：start / add_collaborator / suggest / comment
- reply / accept / reject / merge / export_clean / export_review_pdf / diff
- 颜色由 Word/WPS 自动分配

### 🆕 content=None 支持
- 仅传 `template_name` 即可出模板骨架

---

## v1.4.0（2026-07）

### 🆕 Markdown 直入
- `generate_from_markdown(md_str)` 直接渲染为 DOCX

### 🆕 自动目录 TOC
- Word 原生 TOC 域代码，打开自动更新页码

### 🆕 参考文献格式化引擎
- 4 种 style：APA / GB7714 / MLA / IEEE
- 与 PRO-PPTX 共享 `shared/citation.py`

### 🆕 多级列表自动编号
- 3 级嵌套，numbering.xml 每层编号样式

### 🆕 交叉引用
- `{ref label}` 自动解析为 REF 域代码（Ctrl+点击跳转）

### 🆕 多栏排版 + 首字下沉 + 行号
- 栏间分隔线、首字下沉、行号显示

### 🆕 智能结构识别
- cn×academic / cn×business / en×academic / en×business 四模式
- 字体/字号/对齐按 GB/T 7713、GB/T 9704、APA7、McKinsey 自动切换

### 🆕 数学公式
- `engine/omml.py` — OMML 公式渲染

---

## v1.3.0（2026-07）

### 🆕 图题/表题规范
- 自动编号 + 标签体系

### 🆕 修订追踪
- `engine/revisions.py` — Track Changes 支持

---

## v1.2.0（2026-07）

### 🆕 DOCX 基础生成
- 段落、标题、列表、表格、图片
- 页眉页脚、页码

---

## v1.0.0（2026-07-04）

### 🎉 初始发布
- PaperJSX 语义编译架构
- 确定性布局引擎
- 基本 DOCX 生成能力