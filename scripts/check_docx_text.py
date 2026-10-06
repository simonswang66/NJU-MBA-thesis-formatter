#!/usr/bin/env python3
"""docx ↔ 合并稿文本一致性核查（2026-10-06 装配工作流·能力层）。

逐单元比对 docx 正文文本与合并稿 md 文本，定位不一致处：
- 单元区域＝本单元标题段至下一单元标题段（容忍用户在单元内插入的横排等节）；
- 仅比文本内容：全角/半角空格、中英文间空格、md 粗体标记、<br>、换行、
  "• " 列表前缀一律归一化；表格按"全表内容串"比对（免疫纵向/横向合并差异）；
- 冻结表／横排表配置见 docx 同目录 `.docx-layout.json`；冻结表的内容差异
  单列（用户手工定版，不算错误）。

用法：
    python3 check_docx_text.py <论文.docx> <合并稿.md>
"""

from __future__ import annotations

import json
import re
import sys
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from content_parser import parse_merged_thesis
from section_replacer import SectionReplacer, _unit_title

_UNIT_TITLE_STYLES = {"MBA论文章标题", "MBA论文前置标题"}


def norm(s: str) -> str:
    """内容级归一化（两侧同施）：NFKC、去 md 粗体、<br>→空格、空白折叠、
    非 ASCII 字符邻接空格抹除（用户在单元格内手工换行/断字的伪差异）、
    去行首列表符。"""
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("**", "")
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*/\s*", "/", s)          # 斜杠两侧空格（用户格内换行伪差异）
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s+(?=[^\x00-\x7f])", "", s)
    s = re.sub(r"(?<=[^\x00-\x7f])\s+", "", s)
    s = re.sub(r"^•\s*", "", s)
    return s.strip()


# ── md 侧 ────────────────────────────────────────────────────
def md_unit_items(unit: dict) -> list[tuple[str, str]]:
    """合并稿单元 → [(kind, 归一化文本)]；表格整体一项并打表号标签。"""
    items: list[tuple[str, str]] = []
    for el in unit["elements"]:
        t = el["type"]
        if t == "chapter":
            continue
        if t in ("h1", "h2", "h3", "body", "source_note", "list_item"):
            items.append(("p", norm(el["text"])))
        elif t in ("caption_fig", "caption_tab"):
            items.append(("p", norm(el["text"] + el.get("rest", ""))))
        elif t == "keywords_cn":
            items.append(("p", norm("关键词：" + "；".join(el["keywords"]))))
        elif t == "keywords_en":
            items.append(("p", norm("Key Words: " + ", ".join(el["keywords"]))))
        elif t == "table":
            cells = list(el["headers"]) + [c for row in el["rows"] for c in row]
            cells = [c for c in cells if norm(c)]   # 空单元格不参与（见 docx 侧）
            items.append(("tbl", norm("|".join(cells))))
    return [(k, v) for k, v in items if v]


# ── docx 侧 ──────────────────────────────────────────────────
def docx_unit_items(replacer: SectionReplacer, unit: dict,
                    known_titles: set[str]) -> list[tuple[str, str]]:
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    title = _unit_title(unit)
    anchor, _close = replacer._find_region(title)
    items: list[tuple[str, str]] = []
    cur_table_tag = None
    for el in anchor.itersiblings():
        if el.tag == qn("w:sectPr"):
            break
        if el.tag == qn("w:p"):
            p = Paragraph(el, replacer.doc)
            sname = p.style.name if p.style else ""
            if sname in _UNIT_TITLE_STYLES and p.text.strip():
                # 下一单元标题为界；拆行的长章题首段（如"第5章…SWOT分析"）也是边界
                tnorm = SectionReplacer._norm_title(p.text)
                if tnorm in known_titles or any(
                        len(tnorm) > 4 and kt.startswith(tnorm)
                        for kt in known_titles):
                    break
                continue
            if sname.startswith("toc"):
                continue
            txt = norm(p.text)
            if txt:
                items.append(("p", txt))
                m = re.match(r"^(附?表(?:[A-Z]\d*(?:\.\d+)*|\d+(?:\.\d+)*))", txt.replace(" ", ""))
                cur_table_tag = m.group(1) if m else cur_table_tag
        elif el.tag == qn("w:tbl"):
            tbl = Table(el, replacer.doc)
            vals, seen_tc = [], set()
            for row in tbl.rows:
                for c in row.cells:
                    # 存元素本身而非 id()：lxml 代理是弱引用，id 会被回收复用
                    if c._tc in seen_tc:           # 横/纵合并：锚点格重复出现
                        continue
                    seen_tc.add(c._tc)
                    vals.append(c.text)
            # 空单元格不参与内容比对（分组行/跨年块的占位空格在两侧形态不同）
            vals = [v for v in vals if norm(v)]
            txt = norm("|".join(vals))
            if txt.strip("|"):
                items.append(("tbl", txt, cur_table_tag))
    return [i if len(i) == 3 else (i[0], i[1], None) for i in items]


def load_layout(docx_path: str) -> dict:
    f = Path(docx_path).parent / ".docx-layout.json"
    if f.exists():
        return json.loads(f.read_text(encoding="utf-8"))
    return {}


def main():
    docx_path, md_path = sys.argv[1], sys.argv[2]
    layout = load_layout(docx_path)
    frozen = set(layout.get("frozen", []))
    units = parse_merged_thesis(md_path)
    replacer = SectionReplacer(docx_path,
                               **_layout_for_replacer(layout))
    known_titles = {SectionReplacer._norm_title(_unit_title(u)) for u in units}
    replacer._known_titles = known_titles

    total, frozen_hits = 0, 0
    for unit in units:
        md_items = md_unit_items(unit)
        try:
            dx_items = docx_unit_items(replacer, unit, known_titles)
        except ValueError as e:
            print(f"[{unit['id']}] ⚠️ {e}")
            continue
        sm = SequenceMatcher(None,
                             [t for _, t in md_items],
                             [t for _, t, *_ in dx_items],
                             autojunk=False)
        diffs, fdiffs = [], []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            for k in range(max(i2 - i1, j2 - j1)):
                m = md_items[i1 + k][1] if i1 + k < i2 else "〈无〉"
                d_it = dx_items[j1 + k] if j1 + k < j2 else None
                d = d_it[1] if d_it else "〈无〉"
                tagtbl = d_it[2] if d_it else None
                if tagtbl and any(tagtbl.replace("附表", "表") == f or
                                  tagtbl == f or f.endswith(tagtbl[-3:])
                                  for f in frozen):
                    fdiffs.append((tagtbl, m, d))
                else:
                    diffs.append((m, d))
        if not diffs and not fdiffs:
            print(f"[{unit['id']}] ✅ 一致（{len(md_items)} 项）")
            continue
        for m, d in diffs:
            total += 1
            print(f"[{unit['id']}] ❌")
            print(f"  md  : {m[:100]}")
            print(f"  docx: {d[:100]}")
        for tagtbl, m, d in fdiffs:
            frozen_hits += 1
            print(f"[{unit['id']}] 🧊 冻结表 {tagtbl}（预期内差异）")
            print(f"  md  : {m[:100]}")
            print(f"  docx: {d[:100]}")
    print(f"—— 真实差异 {total} 处；冻结表差异 {frozen_hits} 处 ——")


def _layout_for_replacer(layout: dict) -> dict:
    return {"frozen_tables": layout.get("frozen", []),
            "landscape_tables": layout.get("landscape", [])}


if __name__ == "__main__":
    main()
