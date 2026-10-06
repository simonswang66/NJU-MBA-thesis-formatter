"""
NJU MBA Thesis/Assignment Docx Formatting Engine.

Applies formatting rules from:
《南京大学商学院MBA学位论文写作要求与规范》

Usage:
    from format_engine import FormatEngine
    engine = FormatEngine(mode="thesis")  # or "assignment"
    engine.build(elements, output_path)
"""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Literal

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
from docx.shared import Cm, Pt, Inches, RGBColor
from docx.enum.style import WD_STYLE_TYPE

Mode = Literal["assignment", "thesis", "coursework"]

# ── Constants ──────────────────────────────────────────────────
A4_WIDTH = Cm(21.0)
A4_HEIGHT = Cm(29.7)
MARGIN_TOP = Cm(2.54)
MARGIN_BOTTOM = Cm(2.54)
MARGIN_LEFT = Cm(3.18)
MARGIN_RIGHT = Cm(3.18)
HEADER_DISTANCE = Cm(1.5)
FOOTER_DISTANCE = Cm(1.75)

FONT_BODY_CN = "宋体"
FONT_BODY_EN = "Times New Roman"
FONT_HEADING = "黑体"
FONT_CAPTION = "宋体"
FONT_TOC_SONG = "宋体"
FONT_TOC_FANGSONG = "仿宋"
FONT_ABSTRACT_TITLE_CN = "黑体"
FONT_ABSTRACT_TITLE_EN = "Arial"
FONT_HEADER_FOOTER = "宋体"

SIZE_BODY = Pt(12)          # 小四
SIZE_CAPTION = Pt(10.5)     # 五号
SIZE_CHAPTER = Pt(16)
SIZE_H1 = Pt(14)
SIZE_H2 = Pt(13)
SIZE_H3 = Pt(12)
SIZE_HEADER_FOOTER = Pt(10.5)
SIZE_TOC_LEVEL3 = Pt(10.5)

LINE_SPACING_BODY = Pt(20)
LINE_SPACING_SINGLE = 1.0
LINE_SPACING_1_5 = 1.5

SPACING_CHAPTER_BEFORE = Pt(24)
SPACING_CHAPTER_AFTER = Pt(18)
SPACING_H1_BEFORE = Pt(24)
SPACING_H1_AFTER = Pt(6)
SPACING_H2_BEFORE = Pt(12)
SPACING_H2_AFTER = Pt(6)
SPACING_TOC_ITEM_BEFORE = Pt(6)
SPACING_TOC_ITEM_AFTER = Pt(6)

TWO_CHAR_INDENT = Cm(0.74)  # approx 2 Chinese characters


def _set_cn_font(run, font_name: str):
    """Set East-Asian font on a run via lxml (python-docx only sets ascii)."""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)


def _set_font(run, cn: str, en: str = FONT_BODY_EN, size: Pt = SIZE_BODY, bold: bool = False):
    """Set both CN and EN fonts on a run."""
    run.font.size = size
    run.font.name = en
    run.bold = bold
    _set_cn_font(run, cn)


def _set_paragraph_spacing(para, before: Pt, after: Pt, line_spacing, alignment=None):
    """Configure paragraph spacing and alignment."""
    pf = para.paragraph_format
    pf.space_before = before
    pf.space_after = after
    pf.line_spacing = line_spacing
    if alignment is not None:
        pf.alignment = alignment


def _add_field_code(paragraph, field_code: str):
    """Insert an OOXML field code (for page numbers, TOC, etc.)."""
    run = paragraph.add_run()
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    run._element.append(fld)

    run2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = field_code
    run2._element.append(instr)

    run3 = paragraph.add_run()
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run3._element.append(fld2)


def _add_page_number(paragraph, fmt: str = "ARABIC"):
    """Add page number field to a paragraph."""
    if fmt == "ROMAN":
        _add_field_code(paragraph, " PAGE \\* ROMAN ")
    else:
        _add_field_code(paragraph, " PAGE ")


def _set_header_footer(section, text: str, page_num_fmt: str = "ARABIC"):
    """Configure header (centered section title) and footer (centered page number)."""
    # Header
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.line_spacing = LINE_SPACING_SINGLE
    run = hp.add_run(text)
    _set_font(run, FONT_HEADER_FOOTER, FONT_BODY_EN, SIZE_HEADER_FOOTER)

    # Footer - page number
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.line_spacing = LINE_SPACING_SINGLE
    _add_page_number(fp, page_num_fmt)
    # Set font on the page number field
    for run in fp.runs:
        _set_font(run, FONT_HEADER_FOOTER, FONT_BODY_EN, SIZE_HEADER_FOOTER)


class FormatEngine:
    """Apply NJU MBA formatting to structured content and produce .docx."""

    def __init__(self, mode: Mode = "assignment"):
        self.mode = mode
        self.doc = Document()

        # ── Page setup ──
        section = self.doc.sections[0]
        section.page_width = A4_WIDTH
        section.page_height = A4_HEIGHT
        section.top_margin = MARGIN_TOP
        section.bottom_margin = MARGIN_BOTTOM
        section.left_margin = MARGIN_LEFT
        section.right_margin = MARGIN_RIGHT

        # ── Define styles ──
        self._define_styles()

        # Tracking
        self._chapter_num = 0
        self._figure_num = 0
        self._table_num = 0
        self._current_h1_text = ""  # for header

    # ── Style definitions ──────────────────────────────────────
    def _define_styles(self):
        """Register custom paragraph styles for thesis formatting."""
        # (name, cn_font, size, bold, align, outline_level)
        # outline_level: None = body text, 0 = top-level (like Heading 1), 1 = second, 2 = third
        style_specs = [
            ("NJU Chapter", FONT_HEADING, SIZE_CHAPTER, True, WD_ALIGN_PARAGRAPH.CENTER, 0),
            ("NJU H1", FONT_HEADING, SIZE_H1, False, WD_ALIGN_PARAGRAPH.LEFT, 1),
            ("NJU H2", FONT_HEADING, SIZE_H2, False, WD_ALIGN_PARAGRAPH.LEFT, 2),
            ("NJU Body", FONT_BODY_CN, SIZE_BODY, False, WD_ALIGN_PARAGRAPH.JUSTIFY, None),
            ("NJU Caption", FONT_CAPTION, SIZE_CAPTION, True, WD_ALIGN_PARAGRAPH.CENTER, None),
            ("NJU TOC1", FONT_TOC_SONG, SIZE_BODY, True, WD_ALIGN_PARAGRAPH.LEFT, None),
            ("NJU TOC2", FONT_TOC_SONG, SIZE_BODY, False, WD_ALIGN_PARAGRAPH.LEFT, None),
            ("NJU TOC3", FONT_TOC_FANGSONG, SIZE_TOC_LEVEL3, False, WD_ALIGN_PARAGRAPH.LEFT, None),
            ("NJU Abstract Title", FONT_ABSTRACT_TITLE_CN, SIZE_CHAPTER, True, WD_ALIGN_PARAGRAPH.CENTER, None),
            ("NJU Abstract Body", FONT_BODY_CN, SIZE_BODY, False, WD_ALIGN_PARAGRAPH.JUSTIFY, None),
        ]
        for name, cn_font, size, bold, align, outline_lvl in style_specs:
            style = self.doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.font.size = size
            style.font.name = FONT_BODY_EN
            style.font.bold = bold
            style.paragraph_format.alignment = align
            # Set East-Asian font
            style_xml = style.element
            rPr = style_xml.get_or_add_rPr()
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.insert(0, rFonts)
            rFonts.set(qn("w:eastAsia"), cn_font)
            # Set outline level for navigation pane
            if outline_lvl is not None:
                pPr = style_xml.get_or_add_pPr()
                ol = OxmlElement("w:outlineLvl")
                ol.set(qn("w:val"), str(outline_lvl))
                pPr.append(ol)

        # Default paragraph font
        style = self.doc.styles["Normal"]
        style.font.size = SIZE_BODY
        style.font.name = FONT_BODY_EN
        style.paragraph_format.line_spacing = LINE_SPACING_BODY
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:eastAsia"), FONT_BODY_CN)

    # ── Page helpers ───────────────────────────────────────────
    def _ensure_new_page(self):
        """Ensure the next content starts on a new page."""
        if len(self.doc.paragraphs) > 0:
            last = self.doc.paragraphs[-1]
            run = last.add_run()
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run._element.append(br)

    def _new_section(self, header_text: str = "", page_num_fmt: str = "ARABIC"):
        """Start a new document section (for page number resets, etc.)."""
        new_sec = self.doc.add_section(WD_SECTION_START.NEW_PAGE)
        # Copy page dimensions from first section
        first = self.doc.sections[0]
        new_sec.page_width = first.page_width
        new_sec.page_height = first.page_height
        new_sec.top_margin = first.top_margin
        new_sec.bottom_margin = first.bottom_margin
        new_sec.left_margin = first.left_margin
        new_sec.right_margin = first.right_margin
        if header_text:
            _set_header_footer(new_sec, header_text, page_num_fmt)
        return new_sec

    # ── Element writers ────────────────────────────────────────
    def add_cover_page(self, cover_type: str = "coursework"):
        """Insert a cover page from template at the beginning of the document.

        This works at the OOXML level: the body docx is saved, then unpacked,
        and the cover template content is merged in before repacking. The merge
        is handled by merge_cover.py post-processing step.

        Args:
            cover_type: "thesis" (degree thesis cover + originality statement)
                       or "coursework" (course assignment cover)
        """
        self._has_cover = True
        self._cover_type = cover_type

    def add_chapter(self, title: str, number: int | None = None):
        """Add a chapter heading: '第X章  XXX' (thesis) or just title (assignment)."""
        if number is None:
            self._chapter_num += 1
            number = self._chapter_num
        else:
            self._chapter_num = number

        if self.mode == "thesis":
            text = f"第{number}章  {title}"
        else:
            text = title

        self._ensure_new_page()
        para = self.doc.add_paragraph(text, style="NJU Chapter")
        _set_paragraph_spacing(para, SPACING_CHAPTER_BEFORE, SPACING_CHAPTER_AFTER, LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)

        if self.mode == "thesis":
            self._current_h1_text = text
            self._new_section(text, "ARABIC")
        else:
            self._current_h1_text = title
        self._figure_num = 0
        self._table_num = 0

    def add_h1(self, title: str, number: str = ""):
        """Add a level-1 section heading: 'X.X  XXX'."""
        text = f"{number}  {title}" if number else title
        para = self.doc.add_paragraph(text, style="NJU H1")
        _set_paragraph_spacing(para, SPACING_H1_BEFORE, SPACING_H1_AFTER, LINE_SPACING_SINGLE)
        self._current_h1_text = text

    def add_h2(self, title: str, number: str = ""):
        """Add a level-2 section heading: 'X.X.X  XXX', left-indented."""
        text = f"{number}  {title}" if number else title
        para = self.doc.add_paragraph(text, style="NJU H2")
        _set_paragraph_spacing(para, SPACING_H2_BEFORE, SPACING_H2_AFTER, LINE_SPACING_SINGLE)
        para.paragraph_format.left_indent = TWO_CHAR_INDENT

    def add_h3(self, title: str, number: str = ""):
        """Add a level-3 heading: '(1)XXX', inline with body, left-indented."""
        text = f"({number}){title}" if number else f"{title}"
        para = self.doc.add_paragraph(style="NJU Body")
        para.paragraph_format.left_indent = TWO_CHAR_INDENT
        run = para.add_run(text)
        _set_font(run, FONT_BODY_CN, FONT_BODY_EN, SIZE_BODY, bold=False)

    def _latex_to_omml(self, latex: str, display: bool = False) -> list:
        """Convert LaTeX to OMML elements using pandoc. Returns list of OMML elements."""
        import subprocess, tempfile, zipfile, re
        from lxml import etree
        # pandoc's tex_math_dollars chokes on \tag{...} inside $$...$$ in some contexts;
        # strip it (equation numbering is not needed in docx) before conversion.
        latex_clean = re.sub(r'\s*\\tag\{[^}]*\}', '', latex)
        wrapper = f"$$\n{latex_clean}\n$$\n" if display else f"${latex_clean}$"
        md = tempfile.NamedTemporaryFile(suffix='.md', delete=False, mode='w')
        md.write(wrapper)
        md.close()
        docx_tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
        docx_tmp.close()
        try:
            subprocess.run(['pandoc', md.name, '-o', docx_tmp.name, '--from', 'markdown', '--to', 'docx'],
                         capture_output=True, timeout=10)
            with zipfile.ZipFile(docx_tmp.name) as z:
                xml = z.read('word/document.xml')
            root = etree.fromstring(xml)
            ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                  'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            # Find OMML in the generated docx body
            body = root.find('.//w:body', ns)
            if body is None:
                return []
            # Extract all m:oMath or m:oMathPara elements
            result = []
            for el in body.iter():
                if el.tag.endswith('}oMath') and el.tag.startswith('{http://schemas.openxmlformats.org/officeDocument/2006/math'):
                    result.append(el)
                elif el.tag.endswith('}oMathPara'):
                    result.append(el)
            return result
        except Exception:
            return []
        finally:
            import os
            os.unlink(md.name)
            os.unlink(docx_tmp.name)

    def _add_inline_math(self, para, latex: str):
        """Insert inline LaTeX as proper OMML equation (via pandoc)."""
        import copy
        ommls = self._latex_to_omml(latex, display=False)
        if ommls:
            for omml in ommls:
                para._element.append(copy.deepcopy(omml))
        else:
            run = para.add_run(f"[${latex}$]")
            _set_font(run, FONT_BODY_CN, FONT_BODY_EN, SIZE_BODY)

    def _parse_inline(self, text: str, para, base_font_cn: str = FONT_BODY_CN,
                       base_font_en: str = FONT_BODY_EN, base_size: Pt = SIZE_BODY):
        """Parse **bold**, *italic*, \(LaTeX\) and <br> into docx runs.

        <br> (also <br/>) inserts an in-paragraph line break, so a markdown
        table cell can hold a stacked list within a single source line.
        """
        import re
        for i, seg in enumerate(re.split(r'<br\s*/?>', text, flags=re.IGNORECASE)):
            if i:
                para.add_run().add_break()
            self._parse_segment(seg, para, base_font_cn, base_font_en, base_size)

    def _parse_segment(self, text: str, para, base_font_cn: str = FONT_BODY_CN,
                       base_font_en: str = FONT_BODY_EN, base_size: Pt = SIZE_BODY):
        """Parse **bold**, *italic*, and \(LaTeX\) inline markers into docx runs."""
        import re
        pattern = re.compile(
            r'(\*\*(.+?)\*\*)|'       # **bold**
            r'(\*(.+?)\*)|'            # *italic*
            r'(\\\((.+?)\\\))'         # \(inline LaTeX\)
        )
        last = 0
        for m in pattern.finditer(text):
            if m.start() > last:
                plain = text[last:m.start()]
                run = para.add_run(plain)
                _set_font(run, base_font_cn, base_font_en, base_size)
            if m.group(1):  # **bold**
                run = para.add_run(m.group(2))
                _set_font(run, base_font_cn, base_font_en, base_size, bold=True)
            elif m.group(3):  # *italic*
                run = para.add_run(m.group(4))
                _set_font(run, base_font_cn, base_font_en, base_size)
                run.italic = True
            elif m.group(5):  # \(LaTeX\) → rendered PNG
                self._add_inline_math(para, m.group(6))
            last = m.end()
        if last < len(text):
            run = para.add_run(text[last:])
            _set_font(run, base_font_cn, base_font_en, base_size)

    def add_body(self, text: str):
        """Add body text paragraph with first-line indent. Parses **bold** and \(LaTeX\)."""
        para = self.doc.add_paragraph(style="NJU Body")
        para.paragraph_format.first_line_indent = TWO_CHAR_INDENT
        _set_paragraph_spacing(para, Pt(0), Pt(0), LINE_SPACING_BODY, WD_ALIGN_PARAGRAPH.JUSTIFY)
        # Clear default text run added by style
        for r in para.runs:
            r._element.getparent().remove(r._element)
        self._parse_inline(text, para)

    def add_image(self, image_path: str, width_inches: float = 5.5):
        """Insert an image (for mermaid diagrams, figures, etc.)."""
        from docx.shared import Inches as DocInches
        para = self.doc.add_paragraph(style="NJU Body")
        para.paragraph_format.first_line_indent = Pt(0)
        _set_paragraph_spacing(para, Pt(6), Pt(6), LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)
        run = para.add_run()
        run.add_picture(image_path, width=DocInches(width_inches))

    def add_display_math(self, latex: str, eq_num: str | None = None):
        """Add display LaTeX \[...\] as proper OMML (via pandoc), centered.

        If eq_num is provided (extracted from \\tag{...} by the parser), render
        the equation number right-aligned on the same line as the formula.
        """
        import copy
        ommls = self._latex_to_omml(latex, display=True)
        if ommls:
            para = self.doc.add_paragraph(style="NJU Body")
            para.paragraph_format.first_line_indent = Pt(0)
            _set_paragraph_spacing(para, Pt(6), Pt(6), LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)
            for omml in ommls:
                para._element.append(copy.deepcopy(omml))
            if eq_num:
                # Add a right-aligned tab + equation number on the same line.
                pPr = para._element.get_or_add_pPr()
                ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                tabs_el = pPr.find(qn('w:tabs'))
                if tabs_el is None:
                    tabs_el = OxmlElement('w:tabs')
                    pPr.append(tabs_el)
                tab_el = OxmlElement('w:tab')
                tab_el.set(qn('w:val'), 'right')
                # ~14.6 cm ≈ A4 usable width (21cm - 2×3.18cm margins)
                tab_el.set(qn('w:pos'), '8300')
                tabs_el.append(tab_el)
                # Insert a tab run followed by the equation number in parentheses.
                tab_run = para.add_run()
                tab_run._element.append(OxmlElement('w:tab'))
                num_run = para.add_run(f'({eq_num})')
                _set_font(num_run, FONT_BODY_CN, FONT_BODY_EN, SIZE_BODY)
        else:
            self.add_body(f"[Equation: {latex}]")

    def add_blockquote(self, text: str):
        """Add an indented blockquote-style paragraph. Parses **bold** and \(LaTeX\)."""
        para = self.doc.add_paragraph(style="NJU Body")
        para.paragraph_format.left_indent = TWO_CHAR_INDENT
        para.paragraph_format.right_indent = TWO_CHAR_INDENT
        _set_paragraph_spacing(para, Pt(6), Pt(6), LINE_SPACING_BODY)
        for r in para.runs:
            r._element.getparent().remove(r._element)
        self._parse_inline(text, para)

    def add_figure_caption(self, text: str):
        """Add figure caption below a figure."""
        self._figure_num += 1
        caption = f"图 {self._chapter_num}.{self._figure_num}  {text}"
        para = self.doc.add_paragraph(caption, style="NJU Caption")
        _set_paragraph_spacing(para, Pt(6), Pt(12), LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)
        return caption

    def add_table_caption(self, text: str):
        """Add table caption above a table."""
        self._table_num += 1
        caption = f"表 {self._chapter_num}.{self._table_num}  {text}"
        para = self.doc.add_paragraph(caption, style="NJU Caption")
        _set_paragraph_spacing(para, Pt(6), Pt(6), LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)
        return caption

    def add_three_line_table(self, headers: list[str], rows: list[list[str]], col_widths: list = None):
        """Add a three-line table (三线表) with proper borders."""
        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.autofit = True

        # Header row - use _parse_inline for **bold** support
        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            for run in para.runs:
                run._element.getparent().remove(run._element)
            self._parse_inline(h, para, FONT_BODY_CN, FONT_BODY_EN, SIZE_CAPTION)
            for run in para.runs:
                run.bold = True

        # Data rows - use _parse_inline for **bold** and \(LaTeX\) support
        for r, row_data in enumerate(rows):
            for c, val in enumerate(row_data):
                cell = table.rows[r + 1].cells[c]
                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
                for run in para.runs:
                    run._element.getparent().remove(run._element)
                self._parse_inline(str(val), para, FONT_BODY_CN, FONT_BODY_EN, SIZE_CAPTION)

        # Apply three-line border style via XML
        tbl = table._tbl
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            tbl.insert(0, tblPr)

        borders = OxmlElement("w:tblBorders")
        # Top border (1.5pt)
        top = OxmlElement("w:top")
        top.set(qn("w:val"), "single")
        top.set(qn("w:sz"), "12")  # 1.5pt = 12 eighths of a point
        top.set(qn("w:space"), "0")
        top.set(qn("w:color"), "000000")
        borders.append(top)

        # Bottom border (1.5pt)
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")
        borders.append(bottom)

        # Header separator (0.75pt)
        insideH = OxmlElement("w:insideH")
        insideH.set(qn("w:val"), "single")
        insideH.set(qn("w:sz"), "6")  # 0.75pt
        insideH.set(qn("w:space"), "0")
        insideH.set(qn("w:color"), "000000")
        borders.append(insideH)

        # ECMA-376 CT_TblPr 顺序：tblBorders 须在 shd/tblLayout/tblCellMar/tblLook
        # 之前。原实现 append 到末尾（tblLook 之后），Word 严格校验报"无法读取
        # 的内容"（2026-10-06 全文 docx 29 表全部因此返修）。
        for tag in ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"):
            nxt = tblPr.find(qn(tag))
            if nxt is not None:
                nxt.addprevious(borders)
                break
        else:
            tblPr.append(borders)
        return table

    def add_footnote_marker(self, text: str, footnote_num: int):
        """Add text with a footnote marker ①②③ at the end. The actual footnote content
        is collected for later OOXML post-processing."""
        markers = "①②③④⑤⑥⑦⑧⑨⑩"
        marker = markers[footnote_num - 1] if footnote_num <= 10 else f"({footnote_num})"
        para = self.doc.add_paragraph(style="NJU Body")
        para.paragraph_format.first_line_indent = TWO_CHAR_INDENT
        run = para.add_run(f"{text}{marker}")
        _set_font(run, FONT_BODY_CN, FONT_BODY_EN, SIZE_BODY)
        return marker

    # ── Thesis mode elements ───────────────────────────────────
    def add_abstract_cn(self, text: str, keywords: list[str]):
        """Add Chinese abstract page."""
        self._ensure_new_page()
        para = self.doc.add_paragraph("摘  要", style="NJU Abstract Title")
        _set_paragraph_spacing(para, SPACING_CHAPTER_BEFORE, SPACING_CHAPTER_AFTER, LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)
        # Body
        for para_text in text.split("\n"):
            if para_text.strip():
                p = self.doc.add_paragraph(para_text.strip(), style="NJU Abstract Body")
                p.paragraph_format.first_line_indent = TWO_CHAR_INDENT
                _set_paragraph_spacing(p, Pt(0), Pt(0), LINE_SPACING_BODY, WD_ALIGN_PARAGRAPH.JUSTIFY)
        # Keywords
        kw_para = self.doc.add_paragraph(style="NJU Abstract Body")
        kw_para.paragraph_format.first_line_indent = TWO_CHAR_INDENT
        run = kw_para.add_run("关键词：")
        _set_font(run, FONT_BODY_CN, FONT_BODY_EN, SIZE_BODY, bold=True)
        run2 = kw_para.add_run("；".join(keywords))
        _set_font(run2, FONT_BODY_CN, FONT_BODY_EN, SIZE_BODY)

        if self.mode == "thesis":
            self._new_section("摘要", "ROMAN")

    def add_abstract_en(self, text: str, keywords: list[str]):
        """Add English abstract page."""
        self._ensure_new_page()
        para = self.doc.add_paragraph("Abstract", style="NJU Abstract Title")
        _set_paragraph_spacing(para, SPACING_CHAPTER_BEFORE, SPACING_CHAPTER_AFTER, LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)
        # Override title font to Arial
        for run in para.runs:
            run.font.name = FONT_ABSTRACT_TITLE_EN
            run.font.size = SIZE_CHAPTER
            run.bold = True

        for para_text in text.split("\n"):
            if para_text.strip():
                p = self.doc.add_paragraph(para_text.strip(), style="NJU Body")
                p.paragraph_format.first_line_indent = TWO_CHAR_INDENT
                _set_paragraph_spacing(p, Pt(0), Pt(0), LINE_SPACING_BODY, WD_ALIGN_PARAGRAPH.JUSTIFY)
                for run in p.runs:
                    run.font.name = FONT_BODY_EN

        kw_para = self.doc.add_paragraph(style="NJU Body")
        run = kw_para.add_run("Key Words: ")
        _set_font(run, FONT_ABSTRACT_TITLE_EN, FONT_BODY_EN, SIZE_BODY, bold=True)
        run2 = kw_para.add_run("; ".join(keywords))
        _set_font(run2, FONT_BODY_EN, FONT_BODY_EN, SIZE_BODY)

        if self.mode == "thesis":
            self._new_section("Abstract", "ROMAN")

    def add_toc(self, entries: list[dict]):
        """Add table of contents entries.
        entries: list of {level: 1|2|3, text: str, page: int}
        """
        self._ensure_new_page()
        para = self.doc.add_paragraph("目  录", style="NJU Abstract Title")
        _set_paragraph_spacing(para, SPACING_CHAPTER_BEFORE, SPACING_CHAPTER_AFTER, LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)

        for entry in entries:
            level = entry.get("level", 1)
            text = entry["text"]
            p = self.doc.add_paragraph(text, style=f"NJU TOC{level}")
            if level == 1:
                _set_paragraph_spacing(p, SPACING_TOC_ITEM_BEFORE, SPACING_TOC_ITEM_AFTER, LINE_SPACING_SINGLE)
            elif level == 2:
                _set_paragraph_spacing(p, Pt(0), Pt(0), LINE_SPACING_SINGLE)
                p.paragraph_format.left_indent = TWO_CHAR_INDENT
            elif level == 3:
                _set_paragraph_spacing(p, Pt(0), Pt(0), LINE_SPACING_SINGLE)
                p.paragraph_format.left_indent = TWO_CHAR_INDENT * 2

        if self.mode == "thesis":
            self._new_section("目录", "ROMAN")

    def add_references(self, items: list[str]):
        """Add references section. items are already formatted reference strings."""
        self._ensure_new_page()
        para = self.doc.add_paragraph("参考文献", style="NJU Chapter")
        _set_paragraph_spacing(para, SPACING_CHAPTER_BEFORE, SPACING_CHAPTER_AFTER, LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)

        for i, ref in enumerate(items, 1):
            p = self.doc.add_paragraph(f"[{i}]  {ref}", style="NJU Body")
            _set_paragraph_spacing(p, Pt(0), Pt(3), LINE_SPACING_BODY)

    def add_appendix(self, label: str, content: list[dict]):
        """Add an appendix section.
        content: list of {type: 'body'|'figure'|'table', text: str}
        """
        self._ensure_new_page()
        title = f"附录 {label}" if label else "附录"
        para = self.doc.add_paragraph(title, style="NJU Chapter")
        _set_paragraph_spacing(para, SPACING_CHAPTER_BEFORE, SPACING_CHAPTER_AFTER, LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)

        for item in content:
            if item["type"] == "body":
                self.add_body(item["text"])
            elif item["type"] == "figure":
                self.add_figure_caption(item["text"])
            elif item["type"] == "table":
                self.add_table_caption(item["text"])

    def add_acknowledgement(self):
        """Add acknowledgement/致谢 section."""
        self._ensure_new_page()
        para = self.doc.add_paragraph("致  谢", style="NJU Chapter")
        _set_paragraph_spacing(para, SPACING_CHAPTER_BEFORE, SPACING_CHAPTER_AFTER, LINE_SPACING_SINGLE, WD_ALIGN_PARAGRAPH.CENTER)

    # ── Headers/Footers ────────────────────────────────────────
    def setup_assignment_headers(self):
        """For assignment mode: simple header with document title, no special page numbering."""
        section = self.doc.sections[0]
        _set_header_footer(section, "", "ARABIC")

    # ── Save ────────────────────────────────────────────────────
    def save(self, output_path: str) -> str:
        """Save document to .docx and return the path."""
        path = Path(output_path)
        self.doc.save(str(path))
        return str(path.resolve())


def create_engine(mode: Mode = "assignment") -> FormatEngine:
    """Factory function to create a FormatEngine."""
    return FormatEngine(mode=mode)
