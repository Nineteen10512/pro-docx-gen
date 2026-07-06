---
name: pro-docx-gen
description: Use when users need to create or edit DOCX/Word documents, reports, papers, lesson plans, references, TOC, footnotes, comments, track changes, equations, cross-references, multicolumn layout, PDF/image/text extraction, or Markdown-to-DOCX output with PRO-DOCX semantic JSON.
---

# PRO-DOCX v1.5.3 — PaperJSX 专业 Word 文档生成

> LLM 只写语义，不写数值。字号、段距、颜色全部由确定性代码按 Design Tokens 计算。

---

## 版本亮点 v1.5.3（2026-07）

- **🆕 10 套场景化模板**（`template_name` 参数）：高考教案 `lesson_plan_gaokao` / 完整论文 `thesis_full` / 商业提案 `business_proposal` / 会议纪要 `meeting_minutes` / 中文简历 `resume_cn` / 中文合同 `contract_cn` / 读书笔记 `reading_notes`，覆盖原 academic/business/teaching。
- **🆕 本地 WPS/Office 模板主题提取**：`template_path` 参数传 .docx/.dotx/.wpt 文件自动提色板/字体；`scan_local_templates()` 扫描默认模板目录；`auto_taste_match=True` 时自动做 WCAG AA 对比度校验。
- **🆕 多人协作能力**（`pro_docx_gen.collaboration`）：`start_collaboration` / `add_collaborator` / `suggest_insert|replace|delete` / `add_review_comment` / `reply_comment` / `accept_by_id|author` / `reject_by_id|author` / `merge_collaboration` / `export_clean_copy` / `export_review_pdf` / `diff_documents` / `generate_with_collaboration`。协作者颜色用 `w:val="auto"`，由 Word/WPS 自动分配。
- **content=None 支持**：仅传 `template_name` 即可出模板骨架（教案/合同/简历等场景直接出空白可编辑文档）。
- **向后兼容 100%**：track_changes 默认 False；所有 v1.0–v1.5 JSON / Markdown 输入不改字段即可继续使用。

## 版本亮点 v1.4

- **Markdown 直入**：`generate_from_markdown(md_str)` 支持 Markdown 字符串直接渲染为 docx
- **自动目录 TOC（多级）**：新增 `toc` 节点，插入 Word 原生 TOC 域代码（w:sdt + fldChar），打开 Word 自动更新页码
- **参考文献格式化引擎**：`references` 节点支持 4 种 style（apa/gb7714/mla/ieee），与 PRO-PPTX 共享 `shared/citation.py`
- **References 专业排版**：悬挂缩进 2 字符、方括号编号（GB7714/IEEE）、DOI 蓝色下划线、按 article/book/thesis/conference/webpage 分组
- **多级列表自动编号**：`list` 节点支持 items 嵌套（3 级），通过 numbering.xml 设置每层编号样式
- **交叉引用 + inline `{ref label}`**：`figure/table/equation` 支持 label 字段，正文 `{ref label}` 文本自动解析为 REF 域代码（Ctrl+点击跳转）
- **多栏排版 + 栏间分隔线**：`meta.columns` 支持 int/dict（count/space/sep），段落级连续分节符支持
- **首字下沉**：`paragraph.drop_cap=true` 自动生成首字下沉（3 倍字号加粗）
- **行号显示**：`meta.line_numbers` 控制行号起始、增量、重启规则
- **智能结构识别四模式**：`meta.auto_structure=cn×academic|cn×business|en×academic|en×business`，字体/字号/对齐按 GB/T 7713、GB/T 9704、APA7、McKinsey 自动切换
- **caption 自动转节点**：识别 "图X"/"Fig. X"/"表X"/"Table X"/"公式X" 开头段落，自动转 figure/table/equation 节点并分配 label
- **统一主题字典**：色板下沉到 `shared/themes.py`，与 PRO-PPTX 完全对齐
- **输出文件命名**：默认 `{title}_v{version}.docx`
- **中英双语 docstring**：关键 API 双语文档
- **中文错误提示 + 修复建议**：校验失败给出具体原因与修改建议
- **中文字体 eastAsia 全面审计**：15+ 场景 run 全部设置 eastAsia 字体，零缺失乱码

## 版本亮点 v1.3

- **编辑已有文档**：`load()` / `update_document()` 支持追加章节、替换章节、替换文本、追加段落、删除章节
- **修订追踪（Track Changes）**：新增 `revision` 节点，自动渲染为 Word 原生 w:ins/w:del 红色下划线/删除线，可"接受/拒绝"
- **批注 Comments**：段落/标题支持 `comment` 字段 + 独立 `comment` 节点，自动注入 word/comments.xml
- **脚注/尾注 Footnote/Endnote**：自动注入 word/footnotes.xml（尾注在 v1.2 简化为文末列表）
- **水印 Watermark**：文字水印（VML shape，支持自定义文字/角度/颜色）
- **页面边框、页面设置增强**：纸张大小（A4/Letter/A3/B5/Legal）、方向（portrait/landscape）、四边独立边距、装订线、首页不同、奇偶页不同
- **页眉页脚增强**：图片页眉、"Page X of Y" / "第 X 页，共 Y 页" 域代码
- **13 套主题**（3 套经典 + 10 套新增）：tech/dark/minimal/nature/warm/premium/chinese_red/ocean/forest/sunset，全部与 PPT v1.2 配色对齐，支持自然语言主题名（"科技蓝""中国红""极简"等）与自定义主题 dict deep merge
- **新语义节点**：`equation`（LaTeX 公式占位）、`signature_block` / `signature_line`（签名区）、`page_border`
- **元数据扩展**：keywords/category/comments/status/subject/company/manager 写入 core/extended properties
- **表格表头跨页重复**（w:tblHeader）
- **PDF/图片预览**：`to_pdf()` 通过 LibreOffice 转 PDF，`to_images()` 通过 pdftoppm 输出每页 PNG
- **文本提取**：`extract_text()` 优先 pandoc 转 Markdown（保留修订标记），回退 python-docx

## 核心原则

- **语义优先**：传入 `{"type": "heading", "level": 1, "text": "绪论"}` 而不是 `font_size=18`
- **设计令牌**：所有样式由 tokens 字典统一管理，主题切换只改 token 覆盖
- **确定性布局**：列宽、缩进、图片尺寸全部由代码计算，无需人工调整
- **中文字体安全**：所有 run 都同时设置 w:ascii / w:hAnsi / w:eastAsia，避免 Word 默认回退导致中文乱码

---

## 安装与环境

沙箱已安装/建议安装：python-docx (≥1.0)、matplotlib、numpy、Pillow、lxml、latex2mathml。
可选外部工具（已在沙箱预装）：
- LibreOffice (`soffice`)：PDF 转换
- pdftoppm (`poppler-utils`)：PDF 转图片预览
- pandoc：docx → Markdown 文本提取（保留修订标记）

---

## API 概览

```python
from skills.pro_docx_gen import (
    generate,              # 语义 JSON/Markdown → docx
    load,                  # 加载已有 docx → 结构摘要
    update_document,       # 编辑已有 docx
    to_pdf, to_images,     # PDF/图片预览
    extract_text,          # 文本提取
    outline, word_count,   # 大纲 / 字数估算
    list_themes,           # 主题列表
)
```

---

## 一、三阶段工作流（创建新文档）

### 阶段 1：大纲确认

```python
from skills.pro_docx_gen import outline

doc = {
    "meta": {"title": "深度学习在NLP中的应用研究"},
    "theme": "academic",  # 或 "科技蓝"、"中国红" 等
    "sections": [
        {"title": "绪论", "content": []},
        {"title": "相关工作", "content": []},
        {"title": "方法", "content": []},
        {"title": "实验", "content": []},
        {"title": "结论", "content": []},
    ]
}
print(outline(doc))
```

### 阶段 2：内容填充

根据大纲填充完整 JSON，调用 `word_count(doc, lang="cn")` 检查字数。

### 阶段 3：渲染输出

```python
from skills.pro_docx_gen import generate
generate(doc, "output/paper.docx", theme="academic", lang="cn")
```

Markdown 输入同样支持：

```python
generate("# Title\n\nHello **world**.", "out.docx", theme="teaching", lang="cn")
```

### 随包 Smoke Tests

解包后可在包根目录运行：

```bash
python smoke_tests/run_smoke_tests.py
```

该脚本当前只做随包结构与导入级冒烟检查：确认关键目录和脚本存在，并验证临时 `skills/` 布局下 `shared` 兼容别名可用。

### 交付质量门

正式交付前建议运行：

```bash
python quality_gates/run_quality_gate.py output/report.docx --json-report output/quality_report.json
```

质量门当前只检查 OOXML 包结构、DOCX/PPTX 主文档 XML 是否存在，以及占位符文本（如 `TODO`、`TBD`、`[Image not found]`、`[Image unavailable]`）。若需要把警告也视为失败，可加 `--strict`。

---

## 二、编辑已有文档（v1.2 新增）

### 加载结构摘要

```python
info = load("draft.docx")
print(info["sections"])      # 章节列表（heading + paragraph_count）
print(info["core_props"])    # 标题/作者/关键词
print(info["tables_count"])
```

### update_document 支持的 action

```python
from skills.pro_docx_gen import update_document

edits = [
    # 1) 末尾追加章节
    {"action": "append_section", "title": "附录A", "content": [
        {"type": "paragraph", "text": "补充材料。"},
    ]},
    # 2) 替换整章（按标题匹配）
    {"action": "replace_section", "title_match": "研究方法", "content": [
        {"type": "paragraph", "text": "我们采用新的方法..."},
        {"type": "list", "ordered": True, "items": ["步骤1", "步骤2"]},
    ]},
    # 3) 全文替换文本
    {"action": "replace_text", "find": "旧公司名", "replace": "新公司名"},
    # 4) 在指定标题后追加段落
    {"action": "append_paragraphs", "after_heading": "结论", "content": [
        {"type": "paragraph", "text": "补充说明。"},
    ]},
    # 5) 末尾追加一个段落
    {"action": "append_paragraph", "text": "—— 本文档由 Pro Docx Gen 生成。"},
    # 6) 删除章节
    {"action": "delete_section", "title_match": "已过时的章节"},
]
update_document("draft.docx", edits, "revised.docx", theme="academic", lang="cn")
```

## 三、修订追踪与批注（v1.2 新增）

### revision 节点（Track Changes）

修订节点会渲染为 Word 原生修订标记（红色下划线/删除线），审阅者可直接在 Word 中"接受/拒绝"：

```json
{"type": "paragraph", "text": "基础文本。"},
{"type": "revision", "action": "insert", "text": "这是新增的文本。",
 "author": "AI Assistant", "date": "2026-07-04T12:00:00Z"},
{"type": "revision", "action": "delete", "text": "这段将被删除。"},
{"type": "revision", "action": "replace", "old_text": "旧说法", "new_text": "新说法"}
```

### comment 批注

可以直接挂在 paragraph / heading 上（整段范围），或独立使用：

```json
{"type": "paragraph", "text": "这是有批注的段落。",
 "comment": {"text": "这里需要引用权威来源", "author": "Reviewer", "date": "2026-07-04"}}
```

v1.2 批注范围简化为整段（不做精确到词的选择）。author 默认 "ProDocx Gen"，date 默认当前日期。

---

## 四、脚注、尾注、水印、签名、公式（v1.2 新增）

```json
{"type": "footnote", "text": "这是一个脚注，自动注入 footnotes.xml"},
{"type": "endnote", "text": "这是尾注，集中在文档末尾列表展示"},
{"type": "watermark", "text": "DRAFT", "rotation": -45},
{"type": "equation", "latex": "E = mc^2", "caption": "质能方程"},
{"type": "signature_line", "signer": "申请人", "date": "日期"},
{"type": "signature_block", "name": "张三", "date": "2026-07-04", "title": "签字确认", "lines": 1}
```

文档级水印（每页出现）在顶层设置：

```python
doc = {
    "meta": {"title": "Demo"},
    "watermark": {"enabled": True, "text": "CONFIDENTIAL", "rotation": -45},
    "sections": [...]
}
```

注意：公式会通过 `latex2mathml` 转为 OMML（Word 可编辑公式）；脚注使用原生 footnotes.xml part，尾注简化为文末列表（不注入 endnotes.xml part 但保留引用标记）。

---

## 五、PDF/图片预览与文本提取（v1.2 新增）

```python
from skills.pro_docx_gen import to_pdf, to_images, extract_text

pdf = to_pdf("report.docx", output_dir="output/")        # → report.pdf
images = to_images(pdf, dpi=150)                         # → [report-1.png, ...]
md = extract_text("report.docx", fmt="markdown",
                  track_changes="accept")                # → Markdown 文本
```

## 六、语义节点类型（共 22 种）

### v1.1 节点（12 种，完全兼容）

| 节点 | 说明 |
|---|---|
| `heading` | 标题，`level` 1–5 |
| `paragraph` | 段落，`style`: normal/quote/code/abstract/footnote |
| `list` | 列表，`ordered`: true/false，支持嵌套 `sub_items` |
| `table` | 表格，`headers`/`rows`/`caption`/`col_widths`/`header_repeat` |
| `figure` | 图片，`path`/`caption`/`width_inches` |
| `chart` | 图表（matplotlib 生成 PNG），支持 column/bar/line/pie/doughnut/area/scatter/radar/stacked_column/stacked_bar 共 10 种 |
| `kpi_card` | KPI 卡片，`value`/`label`/`subtext` |
| `callout` | 提示框，`variant`: info/warning/success/danger |
| `page_break` | 分页 |
| `toc` | 自动目录域（需在 Word 中右键"更新域"） |
| `references` | 参考文献块（自动加 References 标题） |
| `appendix` | 附录 |

### v1.2 新增节点（10 种）

| 节点 | 说明 |
|---|---|
| `revision` | 修订标记：action=insert/delete/replace |
| `comment` | 独立批注块（建议直接挂在 `paragraph.comment` 字段） |
| `footnote` | 脚注，自动注入 footnotes.xml |
| `endnote` | 尾注，文末集中列表 |
| `watermark` | 水印（放在文档任意位置都会触发文档级水印；推荐顶层 `watermark` 字段） |
| `page_border` | 页面边框 |
| `equation` | 公式，LaTeX → OMML（Word 可编辑公式） |
| `signature_block` | 签名区表格（签字人/日期/横线） |
| `signature_line` | 签名线（签字人+日期+下划线） |
| `paragraph.comment` 字段 | 给段落/标题挂批注（整段范围） |

---

## 七、主题系统 v1.2

### 内置主题（13 套）

| Key | 名称 | 主色 | 场景 |
|---|---|---|---|
| `academic` | 学术 | 深蓝 | 论文、学术报告 |
| `business` | 商务 | 专业蓝 | 企业报告、方案 |
| `teaching` | 教学 | 教学蓝+橙 | 教案、真题集 |
| `tech` | 科技蓝 | 霓虹蓝 | 技术白皮书、产品文档 |
| `dark` | 暗黑 | 深灰+品红 | 演示配套文档 |
| `minimal` | 极简白 | 纯黑 | 简洁备忘录 |
| `nature` | 自然绿 | 墨绿 | 自然/环保主题 |
| `warm` | 暖橙 | 暖橙色 | 亲和力内容 |
| `premium` | 高端黑金 | 黑+金 | 邀请函、正式发布 |
| `chinese_red` | 中国红 | 朱红+金 | 国风/正式庆典/红头文件 |
| `ocean` | 海洋蓝 | 青蓝 | 海洋/环境/数据 |
| `forest` | 森林 | 深绿+木色 | 自然/户外 |
| `sunset` | 日落橙 | 橙+紫 | 活力/营销/生活 |

自然语言别名自动识别：`"科技蓝"`→`tech`、`"商务"`→`business`、`"中国红"`→`chinese_red`、`"极简"`→`minimal`、`"黑金"`→`premium` 等。

```python
from skills.pro_docx_gen import list_themes
print(list_themes())
```

### 自定义主题

`theme` 参数支持传入 dict，自动 deep merge 到基础主题：

```python
custom = {"color": {"accent": RGBColor(0xFF, 0x66, 0x00)},
          "spacing": {"line_spacing": 1.5}}
generate(doc, "out.docx", theme=custom, lang="cn")
```

---

## 八、顶层字段速查

```python
doc = {
    "meta": {
        "title": "文档标题",           # 必填
        "subtitle": "副标题",
        "author": "作者", "institution": "机构", "date": "日期",
        "keywords": ["AI", "NLP"], "category": "研究报告", "subject": "主题",
        "comments": "备注", "status": "草稿", "company": "公司", "manager": "经理",
        "page_setup": {               # v1.2 页面设置
            "size": "A4",             # A4/Letter/A3/B5/Legal
            "orientation": "portrait", # portrait/landscape
            "margin_top": 1.0, "margin_bottom": 1.0,
            "margin_left": 1.2, "margin_right": 1.2,  # 英寸
            "gutter": 0.0,
            "different_first_page": False,
            "different_odd_even": False,
        },
    },
    "theme": "academic",               # 主题名或自定义 dict
    "lang": "cn",                      # en/cn（传 lang 参数）
    "abstract": {"text": "摘要文本...", "keywords": ["关键词1"]},
    "toc": {"levels": [1, 2, 3]},     # 或 true
    "header": {                        # v1.2 增强
        "text": "页眉文字",
        "image_path": "logo.png",
        "image_width_inches": 1.0,
        "show_on_first_page": False,
        "different_odd_even": False,
    },
    "footer": {                        # v1.2 增强
        "text": "页脚文字",
        "page_number": True,
        "page_x_of_y": False,         # "Page X of Y"
        "page_x_of_y_cn": True,       # "第 X 页，共 Y 页"
        "show_on_first_page": True,
    },
    "watermark": {                     # v1.2 文档级水印
        "enabled": True, "text": "DRAFT", "rotation": -45,
    },
    "page_border": {"enabled": True},  # v1.2 页面边框
    "sections": [
        {"title": "章节名", "level": 1, "content": [/* nodes */]},
    ],
    "references": [/* 参考文献条目 */],
    "appendices": [{"title": "附录", "content": [/* nodes */]}],
}
```

---

## 九、Markdown 语法支持

直接传入 Markdown 字符串即可：

- `# / ## / ###` → 标题/章节
- 普通段落 → paragraph
- `- / *` 无序列表，`1.` 有序列表（支持缩进嵌套）
- `> 引用` → quote 段落
- `` `code` `` 与 ```` ```code blocks```` → code 段落
- \`\`\`chart ... \`\`\` → 图表节点（YAML-like 语法）
- GFM 表格 `| a | b |` → table 节点
- `![caption](path)` → figure 节点
- `---` → page_break

v1.2 新增（在代码块中使用）：
- \`\`\`revision / \`\`\`comment 代码块仍建议通过 JSON API 使用（更精确控制 author/date）

---

## 十、参考文献格式

- `reference.format = "harvard"` → Authors (Year) 'Title', Source, Vol(Issue), pp.xx. doi:xxx
- `"apa"` → APA 风格（同 harvard 近似）
- `"gbt7714"` → [序号] 作者. 题名[J]. 刊名, 年, 卷(期): 页码.

---

## 十一、架构与文件组织

```
skills/pro_docx_gen/
├── __init__.py              # 版本号 & API 导出 (v1.4.0)
├── docx_jsx.py              # 对外 API（generate/load/update_document/to_pdf/...）
├── SKILL.md                 # 本文件
├── requirements.txt
├── compiler/
│   ├── parser.py            # JSON → 扁平节点（v1.2 新节点展开）
│   ├── validators.py        # 输入验证（v1.2 新节点校验）
│   ├── markdown_parser.py   # Markdown → JSON
│   └── __init__.py
├── engine/
│   ├── renderer.py          # OOXML 渲染（v1.2 含修订/批注/脚注/水印/页XofY等）
│   ├── layout.py            # 布局计算（支持四边边距）
│   ├── chart_renderer.py    # matplotlib 图表
│   └── __init__.py
├── tokens/
│   ├── design_tokens.py     # BASE_TOKENS（含 page/header/footer/watermark/revision 等）
│   ├── themes.py            # 14 套主题 + 别名解析 + list_themes
│   └── __init__.py
├── templates/               # 三套经典主题的细调模板
├── smoke_tests/             # 随包冒烟测试
└── quality_gates/           # 交付前结构/渲染质量门
```

---

## 十二、已知限制（v1.4）

1. **批注范围**：批注固定为整段范围，不支持词级精确批注。
2. **尾注 endnote**：尾注简化为文末段落列表，不单独注入 word/endnotes.xml part（脚注 footnote 已完整支持）。
3. **公式 equation**：OMML 公式依赖 `latex2mathml`；未安装时无法把 LaTeX 转为 Word 可编辑公式。
4. **奇偶页不同页眉图片**：首页不同已支持，奇偶页不同设置 global flag 但内容由 Word 默认复制。
5. **update_document 的 replace_section**：已处理段落和表格等 block 级元素，但复杂嵌套内容（如文本框/形状）可能被一并删除。
6. **水印**：通过 VML 实现，部分 Word 版本（WPS/Mac Word）显示效果可能略有差异，但都能显示。
7. **PDF 预览**：依赖 LibreOffice；若用户环境缺失会给出友好提示。

---

## 十三、与 PPT 技能联动

PRO-PPTX v1.4 与 PRO-DOCX v1.4 共享同一套主题色板（14 套）：
在同一份材料中分别生成 PPT 汇报与 DOCX 附件，天然配色一致，无需额外配置。

---

*版本：v1.4.0 | 架构：PaperJSX 语义编译 | Design Tokens 驱动 | 确定性布局*
