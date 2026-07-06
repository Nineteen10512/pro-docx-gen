"""Markdown Parser — Markdown 字符串 → 语义节点列表。

支持常见 MD 语法：
- # / ## / ### / #### / #####  标题
- 普通段落
- - / *  无序列表
- 1. / 2.  有序列表（自动识别）
- >  引用块
- ``` 代码块
- | 表格 (GFM 简单表格)
- ![alt](path) 图片
- --- 分页
"""

import re


def markdown_to_nodes(md_text: str) -> list[dict]:
    """将 Markdown 字符串解析为语义节点列表。

    Returns:
        语义节点列表（不含 meta/sections 包装，外层需补充）。
    """
    lines = md_text.split("\n")
    nodes = []
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 空行跳过
        if not stripped:
            i += 1
            continue

        # 代码块 ```
        if stripped.startswith("```"):
            i, code_block = _parse_code_block(lines, i)
            nodes.append(code_block)
            continue

        # 标题 #
        heading_match = re.match(r"^(#{1,5})\s+(.+)$", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            text = heading_match.group(2).strip()
            nodes.append({"type": "heading", "level": level, "text": text})
            i += 1
            continue

        # 水平线 → 分页
        if re.match(r"^(-{3,}|\*{3,}|_{3,})$", stripped):
            nodes.append({"type": "page_break"})
            i += 1
            continue

        # 表格 |
        if stripped.startswith("|") and i + 1 < len(lines) and re.match(
            r"^\|[\s\-:|]+\|$", lines[i + 1].strip()
        ):
            i, table = _parse_table(lines, i)
            nodes.append(table)
            continue

        # 引用块 >
        if stripped.startswith(">"):
            i, quote = _parse_quote(lines, i)
            nodes.append(quote)
            continue

        # 图片 ![alt](path)
        img_match = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)(.*)$", stripped)
        if img_match:
            alt = img_match.group(1)
            path = img_match.group(2)
            caption = img_match.group(3).strip() or alt
            nodes.append({
                "type": "figure",
                "path": path,
                "caption": caption,
            })
            i += 1
            continue

        # [TOC] 目录标记（大小写不敏感，独占一行）
        if re.match(r"^\[TOC\]\s*$", stripped, re.IGNORECASE):
            nodes.append({"type": "toc"})
            i += 1
            continue

        # 列表
        ul_match = re.match(r"^[\-\*]\s+(.+)$", stripped)
        ol_match = re.match(r"^\d+\.\s+(.+)$", stripped)
        if ul_match or ol_match:
            ordered = bool(ol_match)
            i, list_node = _parse_list(lines, i, ordered)
            nodes.append(list_node)
            continue

        # 普通段落（连续非空非特殊行合并）
        i, para = _parse_paragraph(lines, i)
        if para:
            nodes.append(para)

    return nodes


def markdown_to_document(md_text: str, title: str = "Untitled", **meta_kwargs) -> dict:
    """将 Markdown 解析为完整文档 JSON（可直接传给 generate）。

    一级标题 (# ...) 被视为章节分隔；首个 # 之前的内容放入第一节"引言"。
    """
    lines = md_text.split("\n")
    sections = []
    current_section_title = "Introduction"
    current_section_content = []

    # 先提取首个一级标题作为文档标题（如果有）
    doc_title = title
    content_lines = []
    first_h1_found = False
    for line in lines:
        m = re.match(r"^#\s+(.+)$", line.strip())
        if m and not first_h1_found:
            doc_title = m.group(1).strip()
            first_h1_found = True
            continue
        content_lines.append(line)

    nodes = markdown_to_nodes("\n".join(content_lines))

    # 将 heading(level=1) 切分为 sections
    for node in nodes:
        if node["type"] == "heading" and node.get("level", 1) == 1:
            if current_section_content or current_section_title != "Introduction":
                sections.append({
                    "title": current_section_title,
                    "level": 1,
                    "content": current_section_content,
                })
            current_section_title = node["text"]
            current_section_content = []
        else:
            current_section_content.append(node)

    if current_section_content or current_section_title:
        sections.append({
            "title": current_section_title,
            "level": 1,
            "content": current_section_content,
        })

    doc = {
        "meta": {"title": doc_title, **meta_kwargs},
        "theme": "academic",
        "sections": sections,
    }
    return doc


# ─── 行级解析辅助 ──────────────────────────────────────────────────

def _parse_code_block(lines: list[str], start: int) -> tuple[int, dict]:
    """从 ``` 开始到下一个 ``` 结束。

    支持语言标记：
    - ```chart   → 解析为 chart 节点（简易 YAML-like 语法）
    - 其它        → 原样作为代码段落
    """
    opening = lines[start].strip()
    lang = opening[3:].strip().lower()  # e.g. "chart", "python", ""

    i = start + 1
    buf = []
    while i < len(lines):
        if lines[i].strip().startswith("```"):
            i += 1
            break
        buf.append(lines[i])
        i += 1

    if lang == "chart":
        try:
            return i, _parse_chart_yaml(buf)
        except Exception:
            # 解析失败则回退为代码块
            return i, {
                "type": "paragraph",
                "text": "\n".join(buf),
                "style": "code",
            }

    return i, {
        "type": "paragraph",
        "text": "\n".join(buf),
        "style": "code",
    }


def _parse_chart_yaml(buf: list[str]) -> dict:
    """简易 YAML-like 解析器（仅支持 chart 代码块需要的子集）：
      key: value
      key: [a, b, c]
      series:
        - name: xxx
          values: [1, 2, 3]
    返回 chart 语义节点。
    """
    import json

    # 首先尝试用 pyyaml（如果可用）
    try:
        import yaml  # type: ignore
        data = yaml.safe_load("\n".join(buf))
        if isinstance(data, dict):
            data["type"] = "chart"
            # 兼容 type 字段与 chart_type 字段
            if "type" in data and "chart_type" not in data and data["type"] != "chart":
                data["chart_type"] = data.pop("type")
            return data
    except Exception:
        pass

    # 简易手写解析
    data: dict = {"type": "chart"}
    i = 0
    while i < len(buf):
        line = buf[i]
        stripped = line.rstrip()
        if not stripped.strip() or stripped.lstrip().startswith("#"):
            i += 1
            continue

        # 顶层 key（无缩进或缩进 < 2）
        m = re.match(r"^(\s*)([\w_]+)\s*:\s*(.*)$", line)
        if not m:
            i += 1
            continue
        indent = len(m.group(1))
        key = m.group(2)
        rest = m.group(3).strip()

        if indent >= 2:
            i += 1
            continue

        if rest == "":
            # 块值：可能是 series 列表
            block_lines = []
            i += 1
            while i < len(buf):
                nxt = buf[i]
                nm = re.match(r"^(\s*)([\w_\-]+)?", nxt)
                if nxt.strip() == "":
                    block_lines.append(nxt)
                    i += 1
                    continue
                nindent = len(nm.group(1)) if nm else 0
                if nindent < 2:
                    break
                block_lines.append(nxt)
                i += 1
            data[key] = _parse_block_value(key, block_lines)
            continue

        # 行内值
        parsed = _parse_scalar(rest)
        if key in ("name", "title", "caption") and parsed is not None and not isinstance(parsed, str):
            parsed = str(parsed)
        data[key] = parsed
        i += 1

    # 将 type 字段映射到 chart_type（语法里写 "type: column" 更自然）
    if "type" in data and data["type"] != "chart":
        if "chart_type" not in data:
            data["chart_type"] = data.pop("type")
        else:
            data.pop("type")
    return data


def _parse_scalar(s: str):
    """解析行内标量：数字/bool/列表/字符串。"""
    s = s.strip()
    if s == "":
        return None
    if s.lower() in ("true", "yes"):
        return True
    if s.lower() in ("false", "no"):
        return False
    if s.lower() in ("null", "none", "~"):
        return None
    # 列表 [a, b, c]
    if s.startswith("[") and s.endswith("]"):
        inner = s[1:-1].strip()
        if not inner:
            return []
        parts = [p.strip() for p in _split_csv_respecting_quotes(inner)]
        return [_parse_scalar(p) for p in parts]
    # 数字
    try:
        if "." in s:
            return float(s)
        return int(s)
    except ValueError:
        pass
    # 去除引号
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        return s[1:-1]
    return s


def _split_csv_respecting_quotes(s: str) -> list[str]:
    """按逗号分割，保留引号内的逗号。"""
    parts = []
    buf = []
    quote = None
    for ch in s:
        if ch in ('"', "'") and quote is None:
            quote = ch
            buf.append(ch)
        elif ch == quote:
            quote = None
            buf.append(ch)
        elif ch == "," and quote is None:
            parts.append("".join(buf).strip())
            buf = []
        else:
            buf.append(ch)
    if buf:
        parts.append("".join(buf).strip())
    return parts


def _parse_block_value(key: str, lines: list[str]):
    """解析块值（主要处理 series 列表）。"""
    if key == "series":
        return _parse_series_list(lines)
    # 其它块统一作为字符串
    return "\n".join(lines).strip()


def _parse_series_list(lines: list[str]) -> list[dict]:
    """解析 series 列表：
      - name: xxx
        values: [1,2,3]
      - name: yyy
        values: [4,5,6]
    """
    items = []
    current: dict | None = None
    current_indent = 0
    for line in lines:
        if not line.strip():
            continue
        # 新列表项 "-"
        m = re.match(r"^(\s*)-\s+(.*)$", line)
        if m:
            if current is not None:
                items.append(current)
            current = {}
            indent = len(m.group(1))
            current_indent = indent + 2
            rest = m.group(2).strip()
            kv = re.match(r"^([\w_]+)\s*:\s*(.*)$", rest)
            if kv:
                k = kv.group(1)
                v = kv.group(2).strip()
                if v:
                    current[k] = _parse_scalar(v)
                else:
                    current[k] = None
            continue
        # 列表项内的 key
        if current is not None:
            kv = re.match(r"^(\s*)([\w_]+)\s*:\s*(.*)$", line)
            if kv:
                k = kv.group(2)
                v = kv.group(3).strip()
                parsed = _parse_scalar(v) if v else None
                # name/title/label 这类语义字段即使看起来像数字也保留为字符串
                if k in ("name", "title", "caption") and parsed is not None and not isinstance(parsed, str):
                    parsed = str(parsed)
                current[k] = parsed
    if current is not None:
        items.append(current)
    return items


def _parse_table(lines: list[str], start: int) -> tuple[int, dict]:
    """GFM 简单表格：表头 | 分隔行 | 数据行。"""
    header_line = lines[start].strip()
    sep_line = lines[start + 1].strip()
    headers = [c.strip() for c in header_line.strip("|").split("|")]
    rows = []
    i = start + 2
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped.startswith("|"):
            break
        cells = [c.strip() for c in stripped.strip("|").split("|")]
        # 补齐/截断到 headers 长度
        if len(cells) < len(headers):
            cells.extend([""] * (len(headers) - len(cells)))
        elif len(cells) > len(headers):
            cells = cells[:len(headers)]
        rows.append(cells)
        i += 1
    return i, {"type": "table", "headers": headers, "rows": rows}


def _parse_quote(lines: list[str], start: int) -> tuple[int, dict]:
    """连续 > 行合并为一个引用段落。"""
    i = start
    buf = []
    while i < len(lines):
        stripped = lines[i].strip()
        if stripped.startswith(">"):
            buf.append(stripped.lstrip(">").strip())
            i += 1
        else:
            break
    return i, {"type": "paragraph", "text": " ".join(buf), "style": "quote"}


def _parse_list(lines: list[str], start: int, ordered: bool) -> tuple[int, dict]:
    """解析有序/无序列表，支持缩进表示嵌套。"""
    items = []
    i = start
    while i < len(lines):
        stripped = lines[i].rstrip()
        if not stripped.strip():
            # 空行结束列表
            break
        s = stripped.lstrip()
        indent = len(stripped) - len(s)
        is_ul = bool(re.match(r"^[\-\*]\s+", s))
        is_ol = bool(re.match(r"^\d+\.\s+", s))
        if not (is_ul or is_ol):
            break
        m = re.match(r"^(?:[\-\*]|\d+\.)\s+(.+)$", s)
        text = m.group(1).strip() if m else ""
        items.append({"text": text, "indent": indent})
        i += 1

    # 按 indent 构建嵌套结构
    nested = _build_nested_list(items)
    return i, {"type": "list", "ordered": ordered, "items": nested}


def _build_nested_list(flat_items: list[dict]) -> list:
    """根据缩进构建嵌套列表项。"""
    if not flat_items:
        return []
    root = []
    stack = [(-1, root)]  # (indent, container_list)
    for item in flat_items:
        indent = item["indent"]
        node = {"text": item["text"]}
        # 弹出比当前缩进更深或相等的层级
        while len(stack) > 1 and stack[-1][0] >= indent:
            stack.pop()
        parent = stack[-1][1]
        # 检查上一个同级是否有 sub_items
        if parent and isinstance(parent[-1], dict) and parent[-1].get("sub_items") is None:
            # 只有缩进比上一个大才嵌套
            if indent > stack[-1][0]:
                pass
        parent.append(node)
        node["sub_items"] = []
        stack.append((indent, node["sub_items"]))
    # 清理空 sub_items
    def cleanup(items):
        result = []
        for it in items:
            sub = it.pop("sub_items", [])
            sub = cleanup(sub)
            if sub:
                it["sub_items"] = sub
            result.append(it)
        return result
    return cleanup(root)


def _parse_paragraph(lines: list[str], start: int) -> tuple[int, dict | None]:
    """合并连续普通文本为段落。"""
    i = start
    buf = []
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            break
        # 遇到特殊语法停止
        if (stripped.startswith("#") or stripped.startswith("```")
                or stripped.startswith("|") or stripped.startswith(">")
                or stripped.startswith("![") or re.match(r"^[\-\*]\s", stripped)
                or re.match(r"^\d+\.\s", stripped)
                or re.match(r"^(-{3,}|\*{3,})$", stripped)):
            break
        buf.append(stripped)
        i += 1
    if not buf:
        return i, None
    return i, {"type": "paragraph", "text": " ".join(buf), "style": "normal"}
