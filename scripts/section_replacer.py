#!/usr/bin/env python3
"""章级增量更新器。

把 docx 中**某个单元**的内容替换为合并稿中的最新版本，其余部分（含手工
精调的格式）原样保留。

原理：
- 每个单元＝一个节：单元标题段（MBA论文章标题／MBA论文前置标题）起，
  到本节的收尾 sectPr 段止（sectPr 段保留——它承载页眉与页码设置）；
- 删除区间内全部段落与表格，用 Thesis23EEngine 的同一渲染器在该节内
  重渲染单元内容（图片关系直接在目标文档内创建，无跨文档引用问题）；
- 单元标题文字若变化，同步更新该节页眉；
- 目录是 Word 域，不在合并稿单元内，永不触碰；页码/排版差异不计入一致性。

约定：
- docx 变更前**必须**时间戳备份入 `<docx 同目录>/bak/`（--no-backup 可关）；
- 表格内容随单元重渲染更新，表格格式保留 docx 现有版本。

用法：
    python3 section_replacer.py <论文.docx> <合并稿.md> --unit ch3 [--unit ch4]
    python3 section_replacer.py <论文.docx> <合并稿.md> --unit ch3 --dry-run
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from content_parser import parse_merged_thesis
from thesis23e_engine import Thesis23EEngine, STYLE

# 前置单元在 docx 中的标题（引擎生成，与合并稿单元 id 的映射）
_FRONT_TITLE = {
    "abstract_cn": "摘　要",
    "abstract_en": "Abstract",
    "references": "参考文献",
    "acknowledgement": "致　谢",
}


def _unit_title(unit: dict) -> str:
    """合并稿单元 → docx 中的单元标题文字。"""
    if unit["id"] in _FRONT_TITLE:
        return _FRONT_TITLE[unit["id"]]
    for el in unit["elements"]:
        if el["type"] == "chapter":
            return el["text"]
    raise ValueError(f"单元 {unit['id']} 找不到标题")


class SectionReplacer:
    def __init__(self, docx_path: str, frozen_tables: list[str] | None = None,
                 landscape_tables: list[str] | None = None):
        self.path = docx_path
        self.doc = Document(docx_path)
        # 冻结表（用户手工定版、"以后不要动"的表，如 表5.2/表6.1/表6.2）：
        # 单元替换时按题注锚点保留旧表整体（含格式与回车换行），只更新周边文字
        self.frozen_tables = frozen_tables or []
        # 渲染器复用 Thesis23EEngine 的方法，但作用于现有文档（不做模板预处理）
        self.r = Thesis23EEngine.__new__(Thesis23EEngine)
        self.r.doc = self.doc
        self.r._landscape_tables = set(landscape_tables or [])
        self.r._landscape_open = False

    # ── 区域定位 ─────────────────────────────────────────────
    _TITLE_STYLES = {"MBA论文章标题", "MBA论文前置标题"}

    @staticmethod
    def _norm_title(s: str) -> str:
        """标题匹配归一：空格/全角空格一律抹除（渲染层会把编号后空格转全角，
        合并稿与 docx 的空格形态因此不同，比较时不可直接用原文）。"""
        return s.replace(" ", "").replace("　", "").strip()

    def _find_region(self, title: str):
        """返回 (标题段元素, 收尾sectPr段元素|None)。None＝末节（致谢，body 级 sectPr）。

        区域＝本单元标题段起，至**下一单元标题段**前（用户可能在单元内插入
        横排等自定义节，不能停在第一个 sectPr 段）；收尾段＝区间内最后一个
        带 sectPr 的段（承载本单元节属性，须保留，横排等中段节随内容重建）。
        """
        title_els = [p._p for p in self.doc.paragraphs
                     if p.style and p.style.name in self._TITLE_STYLES and p.text.strip()]
        anchor = None
        want = self._norm_title(title)
        for el in title_els:
            if self._norm_title(self._para_text(el)) == want:
                anchor = el
                break
        if anchor is None:
            # 前缀回退：用户手工把长章题拆成两段时（如第5章），首段命中即可
            for el in title_els:
                t = self._norm_title(self._para_text(el))
                if len(t) > 4 and want.startswith(t):
                    anchor = el
                    break
        if anchor is None:
            raise ValueError(
                f"docx 中找不到单元标题：{title!r}；现存标题：{self.available_titles()}")
        known = getattr(self, "_known_titles", None)
        close = None
        for el in anchor.itersiblings():
            if any(el is t for t in title_els):
                # 下一单元标题为界；拆行的长章题首段（前缀命中已知标题）也是边界
                tnorm = self._norm_title(self._para_text(el))
                if known is None or tnorm in known or any(
                        len(tnorm) > 4 and kt.startswith(tnorm) for kt in known):
                    break
                continue
            if el.tag == qn("w:p"):
                pPr = el.find(qn("w:pPr"))
                if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                    close = el      # 区间内最后一个 sectPr 段＝本单元收尾
            elif el.tag == qn("w:sectPr"):
                break               # body 级 sectPr（末节致谢）
        return anchor, close

    def available_titles(self) -> list[str]:
        return [p.text.strip() for p in self.doc.paragraphs
                if p.style and p.style.name in self._TITLE_STYLES and p.text.strip()]

    # ── 冻结表保留 ───────────────────────────────────────────
    @staticmethod
    def _para_text(el) -> str:
        return "".join(t.text or "" for t in el.iter(qn("w:t")))

    def _caption_matches(self, el, prefix: str) -> bool:
        if el.tag != qn("w:p"):
            return False
        norm = self._para_text(el).replace("　", " ").replace(" ", "")
        return norm.startswith(prefix)

    def _next_table(self, el):
        """caption 段之后最近的 tbl（空段跳过；遇任何非空段落即止）。"""
        for nxt in el.itersiblings():
            if nxt.tag == qn("w:tbl"):
                return nxt
            if nxt.tag == qn("w:p") and self._para_text(nxt).strip():
                return None
        return None

    def _extract_frozen(self, region: list) -> dict:
        """区域内全部表格按题注前缀留存深拷贝（不只冻结表——非冻结表也要
        用旧表骨架做文本原位同步，保住用户手工格式）。"""
        import copy as _copy
        saved = {}
        for el in region:
            if el.tag != qn("w:p"):
                continue
            text = self._para_text(el).replace("　", "").replace(" ", "")
            m = re.match(r"^(附?表(?:[A-Z]\d*(?:\.\d+)*|\d+(?:\.\d+)*))", text)
            if not m:
                continue
            tbl = self._next_table(el)
            if tbl is not None:
                saved[m.group(1)] = _copy.deepcopy(tbl)
        return saved

    def _restore_frozen(self, new_elements: list, saved: dict,
                        unit: dict | None = None) -> tuple[int, list[str]]:
        """在新元素序列中按题注锚点处理旧表：
        - 冻结表：原样回置（文字也不动）；
        - 非冻结表：md 侧文字原位同步进旧表骨架后回置（格式保留、只改文字）；
          结构对不上（增删行列）则保留新渲染表并告警。
        返回 (回置张数, 告警列表)。
        """
        md_tables = _md_tables(unit["elements"]) if unit else {}
        warnings: list[str] = []
        n = 0
        for prefix, old_tbl in saved.items():
            for el in new_elements:
                if not self._caption_matches(el, prefix):
                    continue
                new_tbl = self._next_table(el)
                if new_tbl is None:
                    break
                if prefix in self.frozen_tables:
                    pass                                # 冻结：纯保留
                elif prefix in md_tables:
                    ok, warn = _sync_table_text(
                        old_tbl, md_tables[prefix], prefix)
                    warnings.extend(warn)
                    if not ok:
                        warnings.append(
                            f"{prefix} 结构变化（增删行列），已用新渲染表、格式需手工重做")
                        break
                else:
                    warnings.append(f"{prefix} 在 md 中未找到对应表，保留旧表未动")
                new_tbl.addprevious(old_tbl)
                new_tbl.getparent().remove(new_tbl)
                n += 1
                break
        return n, warnings

    # ── 替换 ─────────────────────────────────────────────────
    def replace_unit(self, unit: dict, dry_run: bool = False) -> dict:
        title = _unit_title(unit)
        anchor_title, anchor_close = self._find_region(title)
        body = self.doc.element.body

        # Word 存盘后节尾 sectPr 常并入本单元末个内容段（不再有空段）：
        # 需剥离到新空段保住节属性，旧内容段纳入删除范围由渲染重建
        close_has_content = False
        if anchor_close is not None:
            close_has_content = bool("".join(
                t.text or "" for t in anchor_close.iter(qn("w:t"))).strip())

        # 区间＝标题段起，至收尾 sectPr 段前（末节则至 body 级 sectPr 前）
        region = []
        for el in anchor_title.itersiblings():
            if el is anchor_close or el.tag == qn("w:sectPr"):
                break
            region.append(el)
        region.insert(0, anchor_title)
        if close_has_content:
            region.append(anchor_close)   # 旧内容段一并删（其文本在合并稿中有对应）
        if dry_run:
            return {"unit": unit["id"], "title": title, "removed": len(region)}

        saved_frozen = self._extract_frozen(region)

        if close_has_content:
            old_close = anchor_close
            pPr = old_close.find(qn("w:pPr"))
            sectPr = pPr.find(qn("w:sectPr"))
            pPr.remove(sectPr)
            new_p = OxmlElement("w:p")
            new_pPr = OxmlElement("w:pPr")
            new_pPr.append(sectPr)
            new_p.append(new_pPr)
            old_close.addnext(new_p)
            anchor_close = new_p

        # 收尾段节属性归一纵排：用户的横排节可能正好落在单元收尾段上
        # （横排由引擎按 .docx-layout.json 配置重建，不依赖保留）
        if anchor_close is not None:
            sp = anchor_close.find(qn("w:pPr")).find(qn("w:sectPr"))
            pgSz = sp.find(qn("w:pgSz")) if sp is not None else None
            if pgSz is not None and pgSz.get(qn("w:orient")) == "landscape":
                pgSz.set(qn("w:w"), "11906")
                pgSz.set(qn("w:h"), "16838")
                if pgSz.get(qn("w:orient")) is not None:
                    del pgSz.attrib[qn("w:orient")]

        for el in region:
            body.remove(el)

        # 在文末渲染单元新内容（python-docx 自动插到 body sectPr 前），
        # 记录新增元素，整体搬到锚点（收尾 sectPr 段）之前；末节无需搬动
        before = set(body)
        if unit["kind"] in ("chapter", "appendix"):
            self.r._add_unit_title(title)
        elif unit["kind"] == "abstract_cn":
            self.r._add_unit_title("摘　要", front=True)
        elif unit["kind"] == "abstract_en":
            para = self.r._add_unit_title("Abstract", front=True)
            for run in para.runs:
                from format_engine import FONT_ABSTRACT_TITLE_EN
                run.font.name = FONT_ABSTRACT_TITLE_EN
        elif unit["kind"] == "references":
            self.r._add_unit_title("参考文献")
        elif unit["kind"] == "acknowledgement":
            self.r._add_unit_title("致　谢")
        elements = unit["elements"]
        if unit["kind"] in ("chapter", "appendix") and elements and \
                elements[0]["type"] == "chapter":
            elements = elements[1:]
        if unit["kind"] == "references":
            for el in elements:
                if el["type"] == "body":
                    self.r._add_reference_entry(el["text"])
        else:
            self.r._appendix = (unit["kind"] == "appendix")  # 附录小节标题去大纲
            self.r._render_elements(elements)
            self.r._appendix = False
        new_elements = [el for el in body if el not in before]
        if anchor_close is not None:
            for el in new_elements:
                body.remove(el)
                anchor_close.addprevious(el)

        restored, tbl_warnings = self._restore_frozen(new_elements, saved_frozen, unit)

        # 标题文字若变化，同步该节页眉
        self._sync_header(anchor_close, title)
        return {"unit": unit["id"], "title": title, "removed": len(region),
                "inserted": len(new_elements), "frozen_restored": restored,
                "table_warnings": tbl_warnings}

    def _sync_header(self, anchor_sectPr_para, title: str):
        """锚点段（或末节 body sectPr）对应节的页眉文字更新为单元标题。"""
        if anchor_sectPr_para is not None:
            pPr = anchor_sectPr_para.find(qn("w:pPr"))
            target = pPr.find(qn("w:sectPr")) if pPr is not None else None
        else:
            target = self.doc.element.body.find(qn("w:sectPr"))
        target_sec = next((s for s in self.doc.sections if s._sectPr is target), None)
        if target_sec is None or target_sec.header.is_linked_to_previous:
            return
        for p in target_sec.header.paragraphs:
            if p.text.strip() and p.runs:
                first, rest = p.runs[0], p.runs[1:]
                first.text = title
                for run in rest:
                    run.text = ""
                return

    def save(self, output_path: str | None = None):
        self.doc.save(output_path or self.path)


def backup(docx_path: str) -> str:
    src = Path(docx_path)
    bak_dir = src.parent / "bak"
    bak_dir.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    dst = bak_dir / f"{src.stem}.bak-{ts}{src.suffix}"
    shutil.copy2(src, dst)
    return str(dst)


def _md_tables(elements: list[dict]) -> dict:
    """单元元素中 表题前缀 → 表数据 的映射（md 侧）。"""
    out = {}
    last_cap = None
    for el in elements:
        if el["type"] == "caption_tab":
            m = re.match(r"^(附?表(?:[A-Z]\d*(?:\.\d+)*|\d+(?:\.\d+)*))", el["text"].replace(" ", "").replace("　", ""))
            last_cap = m.group(1) if m else None
        elif el["type"] == "table":
            if last_cap:
                out[last_cap] = el
            last_cap = None
    return out


def _norm_cell(s: str) -> str:
    """单元格内容级归一（与 check_docx_text 同口径）。"""
    import unicodedata
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("**", "")
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*/\s*", "/", s)
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s+(?=[^\x00-\x7f])", "", s)
    s = re.sub(r"(?<=[^\x00-\x7f])\s+", "", s)
    return s.strip()


def _tc_text(tc) -> str:
    return "".join(t.text or "" for t in tc.iter(qn("w:t")))


def _set_tc_text(tc, text: str):
    """把 md 原文写进单元格：保留首段 pPr 与首 run 的 rPr（格式骨架不动）；
    <br> → w:br 断行；用户的多段落单元格在行数一致时按段分配。"""
    text = text.replace("**", "")
    lines = re.split(r"<br\s*/?>", text, flags=re.IGNORECASE)
    ps = tc.findall(qn("w:p"))

    def _write_para(p, line):
        runs = p.findall(qn("w:r"))
        if runs:
            keep = runs[0]
            for child in list(keep):
                if child.tag != qn("w:rPr"):      # 保留 run 级格式
                    keep.remove(child)
            for r in runs[1:]:
                p.remove(r)
        else:
            keep = OxmlElement("w:r")
            p.append(keep)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = line
        keep.append(t)

    if len(ps) == len(lines):                     # 段落数与行数一致：逐段写
        for p, line in zip(ps, lines):
            _write_para(p, line)
    else:                                         # 否则首段内断行
        _write_para(ps[0], lines[0])
        for line in lines[1:]:
            rbr = OxmlElement("w:r")
            rbr.append(OxmlElement("w:br"))
            ps[0].append(rbr)
            rt = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = line
            rt.append(t)
            ps[0].append(rt)
        for p in ps[1:]:
            for r in p.findall(qn("w:r")):
                p.remove(r)


def _sync_table_text(tbl_el, md_el: dict, prefix: str):
    """md 表格数据原位同步进 docx 表元素（保格式）。返回 (ok, warnings)。

    行/格按"有效单元格"对齐：docx 侧排除纵向合并续格、md 侧排除空占位格。
    结构不一致（增删行列）返回 False，由调用方决定保留新渲染表。
    """
    md_rows = [md_el["headers"]] + md_el["rows"]
    trs = tbl_el.findall(qn("w:tr"))
    warnings: list[str] = []
    if len(trs) != len(md_rows):
        return False, [f"{prefix} 行数不一致（docx {len(trs)} 行 vs md {len(md_rows)} 行）"]
    changed = 0
    for tr, mrow in zip(trs, md_rows):
        tcs = []
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            vm = tcPr.find(qn("w:vMerge")) if tcPr is not None else None
            if vm is not None and vm.get(qn("w:val")) in (None, "continue"):
                continue                          # 纵合并续格：文字在锚点格
            if not _norm_cell(_tc_text(tc)):
                continue                          # 空占位格（分组行/跨年块）不参与
            tcs.append(tc)
        m_vals = [c for c in mrow if _norm_cell(c)]
        if len(tcs) != len(m_vals):
            return False, [f"{prefix} 行结构不一致（docx 有效格 {len(tcs)} vs md {len(m_vals)}）"]
        for tc, val in zip(tcs, m_vals):
            if _norm_cell(_tc_text(tc)) != _norm_cell(val):
                _set_tc_text(tc, val)
                changed += 1
    if changed:
        warnings.append(f"{prefix} 文字原位更新 {changed} 格（格式骨架保留）")
    return True, warnings


def _load_layout(docx_path: str) -> dict:
    """读取 docx 同目录的 .docx-layout.json（冻结表＋横排表配置）。"""
    import json
    f = Path(docx_path).parent / ".docx-layout.json"
    if f.exists():
        return json.loads(f.read_text(encoding="utf-8"))
    return {}


def _replacer_layout(docx_path: str) -> dict:
    lay = _load_layout(docx_path)
    return {"frozen_tables": lay.get("frozen", []),
            "landscape_tables": lay.get("landscape", [])}


def main():
    ap = argparse.ArgumentParser(description="docx 章级增量更新")
    ap.add_argument("docx", help="目标论文 docx")
    ap.add_argument("md", help="带 <!-- @unit:ID --> 单元标记的合并稿")
    ap.add_argument("--unit", action="append", required=True,
                    help="要替换的单元 id（可多次）：ch1…ch7/abstract_cn/abstract_en/"
                         "references/appendix_a…/acknowledgement")
    ap.add_argument("-o", "--output", default=None, help="输出路径（默认原位更新）")
    ap.add_argument("--no-backup", action="store_true", help="跳过变更前时间戳备份")
    ap.add_argument("--no-purge", action="store_true",
                    help="跳过保存后的未引用媒体清理（默认自动清理）")
    ap.add_argument("--no-stamp", action="store_true",
                    help="跳过保存后的作者元数据改写（默认改为系统用户）")
    ap.add_argument("--dry-run", action="store_true", help="只报告将删除的段落数，不写文件")
    args = ap.parse_args()

    units = {u["id"]: u for u in parse_merged_thesis(args.md)}
    for uid in args.unit:
        if uid not in units:
            print(f"!! 合并稿中无单元 {uid}；可选：{sorted(units)}")
            sys.exit(1)

    if not args.dry_run and not args.no_backup and args.output is None:
        print(f"备份 → {backup(args.docx)}")

    replacer = SectionReplacer(args.docx, **_replacer_layout(args.docx))
    replacer._known_titles = {SectionReplacer._norm_title(_unit_title(u))
                              for u in units.values()}
    for uid in args.unit:
        info = replacer.replace_unit(units[uid], dry_run=args.dry_run)
        if args.dry_run:
            print(f"[dry-run] {info['unit']}: 将删除 {info['removed']} 个元素（{info['title']}）")
        else:
            extra = f"，表格保留回置 {info['frozen_restored']} 张" if info.get("frozen_restored") else ""
            print(f"OK {info['unit']}: 删 {info['removed']} 元素 → 插 {info['inserted']} 元素（{info['title']}）{extra}")
            for w in info.get("table_warnings", []):
                print(f"  ⚠️ {w}")
    if not args.dry_run:
        replacer.save(args.output)
        print(f"已保存：{args.output or args.docx}")
        # 后处理（2026-10-06 用户裁定）：清未引用媒体＋作者改写为系统用户
        from docx_postprocess import purge_unreferenced_media, stamp_author
        target = args.output or args.docx
        if not args.no_purge:
            info = purge_unreferenced_media(target)
            if info["dropped_rels"] or info["dropped_media"]:
                print(f"purge: 删关系 {info['dropped_rels']} 项，删文件 {info['dropped_media']}")
        if not args.no_stamp:
            stamp_author(target)


if __name__ == "__main__":
    main()
