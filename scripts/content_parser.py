"""
Content parsers for .md and .docx input files.

Returns a structured elements list consumed by FormatEngine.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Literal

ElementType = Literal["chapter", "h1", "h2", "h3", "body", "blockquote",
                       "figure", "table", "footnote", "abstract_cn", "abstract_en",
                       "toc", "references", "appendix", "acknowledgement",
                       "keywords_cn", "keywords_en", "table_caption", "figure_caption",
                       "image", "list"]

ParsedElement = dict  # {type: ElementType, ...}

_TABLE_SEP_CELL = re.compile(r':?-+:?')


def _is_table_sep(sep_line: str, header_line: str) -> bool:
    """判定 GFM 表格分隔行（`|:---|---:|` 之类），**连字符数不限（≥1）**。

    旧判据是 `"---" in 下一行`：只要某一格写成 `:--:`，整行凑不出三个连续连字符，
    判据即落空，该表会被当成正文段落逐行排出。2026-09-24 第3章表3.1 即因此漏掉。
    另校验与表头列数一致，避免正文里含竖线的行被误判成表格。
    """
    s = sep_line.strip()
    if not (s.startswith("|") and s.endswith("|")):
        return False
    cells = [c.strip() for c in s.strip("|").split("|")]
    if not cells or not all(_TABLE_SEP_CELL.fullmatch(c) for c in cells):
        return False
    h = header_line.strip()
    if not (h.startswith("|") and h.endswith("|")):
        return False          # 表头也须是竖线行，否则正文里的「a | b」会被误判
    return len(cells) == len(h.strip("|").split("|"))


def _clean_inline_markdown(text: str) -> str:
    """Remove inline markdown formatting (**bold**, *italic*, `code`)."""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    return text


def _render_mermaid(code: str) -> str | None:
    """Render mermaid diagram to PNG using Playwright with inline mermaid.js. Returns path or None."""
    import tempfile, os, http.server, threading, socket
    tmp_png = None
    server = None
    try:
        from playwright.sync_api import sync_playwright

        # Check if mermaid.min.js exists locally from npm
        mermaid_js = None
        for p in ['node_modules/mermaid/dist/mermaid.min.js',
                  '/usr/local/lib/node_modules/mermaid/dist/mermaid.min.js']:
            if os.path.exists(p):
                mermaid_js = open(p).read()
                break

        if not mermaid_js:
            # Download once and cache
            import urllib.request, ssl
            ctx = ssl._create_unverified_context()
            req = urllib.request.Request('https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js')
            r = urllib.request.urlopen(req, context=ctx, timeout=10)
            mermaid_js = r.read().decode('utf-8')

        html = f'<!DOCTYPE html><html><head><meta charset="utf-8"><script>{mermaid_js}</script>'
        html += '<script>mermaid.initialize({startOnLoad:true,theme:"default"});</script>'
        html += f'</head><body><pre class="mermaid">{code}</pre></body></html>'

        # Serve HTML via localhost to avoid file:// CSP issues
        html_bytes = html.encode('utf-8')
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.bind(('127.0.0.1', 0))
        port = sock.getsockname()[1]
        sock.close()

        class Handler(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(html_bytes)
            def log_message(self, *args): pass

        server = http.server.HTTPServer(('127.0.0.1', port), Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()

        tmp_png = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        tmp_png.close()

        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={'width': 1200, 'height': 800})
            page.goto(f'http://127.0.0.1:{port}', timeout=15000)
            page.wait_for_selector('svg', timeout=10000)
            svg = page.query_selector('svg')
            if svg:
                svg.screenshot(path=tmp_png.name)
            browser.close()

        if os.path.getsize(tmp_png.name) > 500:
            return tmp_png.name
    except Exception:
        pass
    finally:
        if server:
            server.shutdown()
    return None


def parse_markdown(filepath: str) -> list[ParsedElement]:
    """Parse a .md file into structured elements."""
    path = Path(filepath)
    text = path.read_text(encoding="utf-8")
    return parse_markdown_text(text)


def parse_markdown_text(text: str, thesis_extras: bool = False) -> list[ParsedElement]:
    """Parse markdown text into structured elements.

    thesis_extras=True 时启用论文合并稿后处理（_apply_thesis_extras）：
    图表题、来源注、列表项、关键词等元素再分类。旧流程一律用 False。
    """
    lines = text.split("\n")

    elements: list[ParsedElement] = []
    footnotes_def: dict[str, str] = {}
    footnote_counter = 0

    # First pass: extract footnote definitions [^n]: ...
    fn_pattern = re.compile(r'^\[\^(\d+)\]:\s*(.+)$')
    clean_lines = []
    for line in lines:
        m = fn_pattern.match(line)
        if m:
            footnotes_def[m.group(1)] = m.group(2)
        else:
            clean_lines.append(line)

    # Second pass: parse content
    i = 0
    while i < len(clean_lines):
        line = clean_lines[i]

        # Skip empty lines
        if not line.strip():
            i += 1
            continue

        # Headers
        if line.startswith("# "):
            elements.append({"type": "chapter", "text": line[2:].strip()})
        elif line.startswith("## "):
            elements.append({"type": "h1", "text": line[3:].strip()})
        elif line.startswith("### "):
            elements.append({"type": "h2", "text": line[4:].strip()})
        elif line.startswith("#### "):
            elements.append({"type": "h3", "text": line[5:].strip()})

        # Blockquote
        elif line.startswith(">"):
            quote_lines = []
            while i < len(clean_lines) and clean_lines[i].startswith(">"):
                # Strip the leading '>' and optional single space
                content = clean_lines[i][1:]
                if content.startswith(" "):
                    content = content[1:]
                quote_lines.append(content)
                i += 1
            i -= 1  # compensate for loop increment
            elements.append({"type": "blockquote", "text": " ".join(quote_lines)})

        # Images
        elif line.startswith("!["):
            m = re.match(r'!\[(.*?)\]\((.*?)\)', line)
            if m:
                elements.append({"type": "image", "alt": m.group(1), "src": m.group(2)})

        # Fenced code blocks (```mermaid, ```)
        elif line.startswith("```"):
            lang = line[3:].strip().lower()
            code_lines = []
            i += 1
            while i < len(clean_lines) and not clean_lines[i].startswith("```"):
                code_lines.append(clean_lines[i])
                i += 1
            code = "\n".join(code_lines)
            if lang == "mermaid":
                img_path = _render_mermaid(code)
                if img_path:
                    elements.append({"type": "mermaid_image", "src": img_path, "code": code})
                else:
                    elements.append({"type": "body", "text": f"[Mermaid diagram - render failed]\n{code}"})
            else:
                elements.append({"type": "body", "text": f"```{lang}\n{code}\n```"})

        # Display math \[...\]
        elif line.strip().startswith("\\["):
            math_lines = [line]
            if not line.strip().endswith("\\]"):
                i += 1
                while i < len(clean_lines) and not clean_lines[i].strip().endswith("\\]"):
                    math_lines.append(clean_lines[i])
                    i += 1
                if i < len(clean_lines):
                    math_lines.append(clean_lines[i])
            math_text = " ".join(m.strip() for m in math_lines)
            math_text = re.sub(r'^\\\[|\\\]$', '', math_text).strip()
            # Extract equation number from \tag{...} so OMML rendering can
            # succeed (pandoc can't handle \tag in all contexts).
            tag_match = re.search(r'\\tag\{([^}]*)\}', math_text)
            eq_num = tag_match.group(1) if tag_match else None
            math_text = re.sub(r'\s*\\tag\{[^}]*\}', '', math_text).strip()
            elements.append({"type": "display_math", "text": math_text, "eq_num": eq_num})

        # Tables (simple pipe tables)
        elif "|" in line and i + 1 < len(clean_lines) and _is_table_sep(clean_lines[i + 1], line):
            def _split_pipe(row: str) -> list[str]:
                """Split a pipe-table row, keeping interior empty cells."""
                parts = row.split("|")
                if parts[0].strip() == "" and parts[-1].strip() == "" and len(parts) > 2:
                    return [c.strip() for c in parts[1:-1]]
                return [c.strip() for c in parts if c.strip()]
            headers = _split_pipe(line)
            i += 2  # skip separator
            rows = []
            while i < len(clean_lines) and "|" in clean_lines[i]:
                row = _split_pipe(clean_lines[i])
                if row:
                    rows.append(row)
                i += 1
            i -= 1
            # Normalize: pad every row to header width with empty strings
            ncols = len(headers)
            for row in rows:
                while len(row) < ncols:
                    row.append("")
            elements.append({"type": "table", "headers": headers, "rows": rows})

        # Horizontal rules (skip)
        elif line.strip() in ("---", "***", "___"):
            pass

        # Body text
        else:
            text = line.strip()
            # Check for inline footnote [^n]
            fn_match = re.search(r'\[\^(\d+)\]', text)
            if fn_match:
                footnote_counter += 1
                fn_id = fn_match.group(1)
                markers = "①②③④⑤⑥⑦⑧⑨⑩"
                marker = markers[footnote_counter - 1] if footnote_counter <= 10 else f"({footnote_counter})"
                # Keep visible footnote marker in text
                text = re.sub(r'\[\^\d+\]', marker, text)
                fn_text = footnotes_def.get(fn_id, "")
                elements.append({"type": "body", "text": text, "footnote": fn_text, "fn_num": footnote_counter})
            else:
                # Merge consecutive body lines into a paragraph
                body_lines = [text]
                i += 1
                # Stop at any structural element (headers, blockquotes, display
                # math, tables, fenced code, etc.) so they're not swallowed.
                body_stops = ("#", ">", "!", "|", "-", "\\[", "$$", "```")
                while i < len(clean_lines) and clean_lines[i].strip() and \
                      not any(clean_lines[i].startswith(p) for p in body_stops):
                    if not re.match(r'\[\^\d+\]:', clean_lines[i]):
                        body_lines.append(clean_lines[i].strip())
                    i += 1
                i -= 1
                elements.append({"type": "body", "text": " ".join(body_lines)})

        i += 1

    if thesis_extras:
        elements = _apply_thesis_extras(elements)
    return elements


# ── 论文合并稿解析（23E 模板模式）────────────────────────────
#
# 合并稿＝把论文各部分按顺序拼接成单个 markdown，单元以 <!-- @unit:ID --> 注释分隔。
# 本模块只负责"解析"，单元顺序由合并稿本身决定。

_UNIT_KIND = {
    "abstract_cn": "abstract_cn", "abstract_en": "abstract_en",
    "references": "references", "acknowledgement": "acknowledgement",
}

# 图表题：整段粗体（**图1.1 …**），附表题可带段尾非粗体说明（**附表A.1 …**（请…））
_CAPTION_RE = re.compile(r"^\*\*((?:附表|图|表)[\dA-Z].*?)\*\*(.*)$")
_SOURCE_NOTE_RE = re.compile(r"^(?:资料来源|数据来源)：")
_KEYWORDS_CN_RE = re.compile(r"^\*\*关键词：\*\*\s*(.+)$")
_KEYWORDS_EN_RE = re.compile(r"^\*\*Key ?Words?:?\*\*:?\s*(.+)$", re.IGNORECASE)


def _apply_thesis_extras(elements: list[ParsedElement]) -> list[ParsedElement]:
    """把 body 元素按论文版式惯例再分类（关键词→列表项→图表题→来源注）。"""
    out: list[ParsedElement] = []
    for el in elements:
        if el["type"] != "body":
            out.append(el)
            continue
        text = el["text"]
        m = _KEYWORDS_CN_RE.match(text)
        if m:
            kws = [k.strip() for k in m.group(1).split("；") if k.strip()]
            out.append({"type": "keywords_cn", "keywords": kws})
            continue
        m = _KEYWORDS_EN_RE.match(text)
        if m:
            kws = [k.strip() for k in m.group(1).split(",") if k.strip()]
            out.append({"type": "keywords_en", "keywords": kws})
            continue
        if text.startswith("- "):
            out.append({"type": "list_item", "text": text[2:].strip()})
            continue
        m = _CAPTION_RE.match(text)
        if m and not m.group(2).strip().startswith("所示"):
            label = m.group(1).strip()
            kind = "caption_tab" if ("表" in label[:3]) else "caption_fig"
            out.append({"type": kind, "text": label, "rest": m.group(2).strip()})
            continue
        if _SOURCE_NOTE_RE.match(text):
            out.append({"type": "source_note", "text": text})
            continue
        out.append(el)
    return out


def _classify_unit(uid: str) -> str:
    if uid in _UNIT_KIND:
        return _UNIT_KIND[uid]
    if uid.startswith("appendix_"):
        return "appendix"
    return "chapter"


def _shape_unit_elements(kind: str, elements: list[ParsedElement]) -> list[ParsedElement]:
    """按单元类型整理元素：摘要/参考文献/致谢单元的 `# 标题` 由引擎自行渲染，丢弃；
    章与附录保留 chapter 元素（标题全文照排，含"第N章"/"附录X："前缀）。"""
    if kind in ("abstract_cn", "abstract_en", "references", "acknowledgement"):
        if elements and elements[0]["type"] == "chapter":
            elements = elements[1:]
    return elements


def parse_merged_thesis(filepath: str) -> list[dict]:
    """解析带单元标记的合并稿 → [{id, kind, elements}]。

    kind: abstract_cn / abstract_en / chapter / references / appendix / acknowledgement
    """
    text = Path(filepath).read_text(encoding="utf-8")
    pat = re.compile(r"<!-- @unit:([\w]+) -->\n(.*?)<!-- @/unit:\1 -->", re.S)
    units: list[dict] = []
    for m in pat.finditer(text):
        uid, body = m.group(1), m.group(2)
        body = re.sub(r"(?m)^-{3,}\s*$", "", body)   # 防御：去悬空 hr 线
        elements = parse_markdown_text(body, thesis_extras=True)
        kind = _classify_unit(uid)
        units.append({"id": uid, "kind": kind,
                      "elements": _shape_unit_elements(kind, elements)})
    return units


def parse_docx(filepath: str) -> list[ParsedElement]:
    """Parse a .docx file into structured elements by reading styles and content."""
    from docx import Document

    doc = Document(filepath)
    elements: list[ParsedElement] = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name if para.style else ""

        # Map Word styles to element types
        if "Heading 1" in style_name or "heading 1" in style_name:
            elements.append({"type": "chapter", "text": text})
        elif "Heading 2" in style_name or "heading 2" in style_name:
            elements.append({"type": "h1", "text": text})
        elif "Heading 3" in style_name or "heading 3" in style_name:
            elements.append({"type": "h2", "text": text})
        elif "Heading 4" in style_name or "heading 4" in style_name:
            elements.append({"type": "h3", "text": text})
        else:
            # Check for blockquote style
            if para.paragraph_format.left_indent and para.paragraph_format.left_indent > 0:
                elements.append({"type": "blockquote", "text": text})
            else:
                elements.append({"type": "body", "text": text})

    # 全文没有任何 Heading 样式时按文本模式兜底识别层级：
    # 第X章→章、X.Y→一级节、X.Y.Z→二级节（编号后须跟空白，防正文数字误判）
    if not any(el["type"] in ("chapter", "h1", "h2", "h3") for el in elements):
        for el in elements:
            if el["type"] != "body":
                continue
            t = el["text"]
            if _DOCX_CHAPTER_RE.match(t):
                el["type"] = "chapter"
            elif _DOCX_H2_RE.match(t):
                el["type"] = "h2"
            elif _DOCX_H1_RE.match(t):
                el["type"] = "h1"

    return elements


_DOCX_CHAPTER_RE = re.compile(r"^第[0-9一二三四五六七八九十百]+章[\s　]")
_DOCX_H2_RE = re.compile(r"^\d+\.\d+\.\d+[\s　]")
_DOCX_H1_RE = re.compile(r"^\d+\.\d+[\s　]")


def parse_file(filepath: str) -> tuple[list[ParsedElement], str]:
    """Auto-detect file type and parse.
    Returns (elements, file_type) where file_type is 'md' / 'txt' / 'docx'.
    .txt 按 markdown 规则解析（无 # 标记的纯文本会全部成为正文，
    建议先补标题层级再格式化）。
    """
    path = Path(filepath)
    suffix = path.suffix.lower()
    if suffix in (".md", ".txt"):
        return parse_markdown(filepath), suffix.lstrip(".")
    elif suffix == ".docx":
        return parse_docx(filepath), "docx"
    else:
        raise ValueError(f"Unsupported file type: {suffix}. Expected .md / .txt / .docx")
