"""
OOXML post-processing: insert footnotes (页下注) into .docx files.

python-docx does not natively support footnotes. This script:
1. Unpacks the .docx to an OOXML directory
2. Creates the footnotes part with proper XML
3. Modifies document.xml to insert footnote references
4. Repacks to a valid .docx

unpack/pack 由本包 ooxml_zip.py 提供（自实现，无外部依赖）。

Usage:
    from footnote_inserter import insert_footnotes
    insert_footnotes("input.docx", "output.docx", [...])
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from ooxml_zip import unpack, pack
from defusedxml.minidom import parseString as safe_parse
from xml.dom.minidom import Document, Element


FOOTNOTE_MARKERS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def _make_footnotes_xml(footnotes: list[str]) -> str:
    """Generate w:footnotes XML content."""
    doc = Document()
    footnotes_el = doc.createElement("w:footnotes")
    footnotes_el.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    footnotes_el.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    # Footnote 0 (separator - required)
    fn0 = doc.createElement("w:footnote")
    fn0.setAttribute("w:id", "0")
    p0 = doc.createElement("w:p")
    pPr0 = doc.createElement("w:pPr")
    spacing0 = doc.createElement("w:spacing")
    spacing0.setAttribute("w:after", "0")
    spacing0.setAttribute("w:line", "240")
    spacing0.setAttribute("w:lineRule", "auto")
    pPr0.appendChild(spacing0)
    p0.appendChild(pPr0)
    r0 = doc.createElement("w:r")
    sep = doc.createElement("w:separator")
    r0.appendChild(sep)
    p0.appendChild(r0)
    fn0.appendChild(p0)
    footnotes_el.appendChild(fn0)

    # Footnote -1 (continuation separator - required)
    fn_minus = doc.createElement("w:footnote")
    fn_minus.setAttribute("w:id", "-1")
    p_minus = doc.createElement("w:p")
    pPr_minus = doc.createElement("w:pPr")
    spacing_minus = doc.createElement("w:spacing")
    spacing_minus.setAttribute("w:after", "0")
    spacing_minus.setAttribute("w:line", "240")
    spacing_minus.setAttribute("w:lineRule", "auto")
    pPr_minus.appendChild(spacing_minus)
    p_minus.appendChild(pPr_minus)
    r_minus = doc.createElement("w:r")
    cont_sep = doc.createElement("w:continuationSeparator")
    r_minus.appendChild(cont_sep)
    p_minus.appendChild(r_minus)
    fn_minus.appendChild(p_minus)
    footnotes_el.appendChild(fn_minus)

    # Actual footnotes
    for idx, fn_text in enumerate(footnotes, 1):
        fn = doc.createElement("w:footnote")
        fn.setAttribute("w:id", str(idx))

        p = doc.createElement("w:p")
        pPr = doc.createElement("w:pPr")
        pStyle = doc.createElement("w:pStyle")
        pStyle.setAttribute("w:val", "FootnoteText")
        pPr.appendChild(pStyle)
        spacing = doc.createElement("w:spacing")
        spacing.setAttribute("w:after", "0")
        spacing.setAttribute("w:line", "240")
        spacing.setAttribute("w:lineRule", "auto")
        pPr.appendChild(spacing)
        p.appendChild(pPr)

        # Footnote reference run
        r_ref = doc.createElement("w:r")
        rPr_ref = doc.createElement("w:rPr")
        style_ref = doc.createElement("w:rStyle")
        style_ref.setAttribute("w:val", "FootnoteReference")
        rPr_ref.appendChild(style_ref)
        r_ref.appendChild(rPr_ref)
        ref = doc.createElement("w:footnoteRef")
        r_ref.appendChild(ref)
        p.appendChild(r_ref)

        # Space after marker
        r_space = doc.createElement("w:r")
        t_space = doc.createElement("w:t")
        t_space.setAttribute("xml:space", "preserve")
        t_space.appendChild(doc.createTextNode(" "))
        r_space.appendChild(t_space)
        p.appendChild(r_space)

        # Footnote text
        r_text = doc.createElement("w:r")
        rPr_text = doc.createElement("w:rPr")
        sz = doc.createElement("w:sz")
        sz.setAttribute("w:val", "18")  # 9pt
        rPr_text.appendChild(sz)
        rFonts = doc.createElement("w:rFonts")
        rFonts.setAttribute("w:eastAsia", "宋体")
        rFonts.setAttribute("w:ascii", "Times New Roman")
        rPr_text.appendChild(rFonts)
        r_text.appendChild(rPr_text)
        t = doc.createElement("w:t")
        t.setAttribute("xml:space", "preserve")
        t.appendChild(doc.createTextNode(fn_text))
        r_text.appendChild(t)
        p.appendChild(r_text)

        fn.appendChild(p)
        footnotes_el.appendChild(fn)

    return footnotes_el.toxml(encoding="utf-8").decode("utf-8")


def _make_footnotes_rels_xml() -> str:
    """Generate relationships for the footnotes part."""
    doc = Document()
    rels = doc.createElement("Relationships")
    rels.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
    # No external relationships needed for basic footnotes
    return rels.toxml(encoding="utf-8").decode("utf-8")


def insert_footnotes(
    input_path: str,
    output_path: str,
    footnotes: list[str],
    footnote_positions: list[str] | None = None,
) -> str:
    """
    Insert footnotes into a .docx file via OOXML manipulation.

    Args:
        input_path: Path to the .docx file (generated by python-docx).
        output_path: Where to save the modified .docx.
        footnotes: List of footnote text strings.
        footnote_positions: List of marker texts (①②③...) to find in document.xml.
                           If None, appends all footnotes at end of document.

    Returns:
        Path to the output .docx file.
    """
    tmpdir = tempfile.mkdtemp(prefix="docx_fn_")
    unpacked = Path(tmpdir) / "unpacked"
    unpacked.mkdir(parents=True, exist_ok=True)

    try:
        # Unpack the docx
        _, msg = unpack(str(Path(input_path).resolve()), str(unpacked), merge_runs=False)
        if msg.startswith("Error:"):
            raise RuntimeError(f"Unpack failed: {msg}")

        # Find document.xml
        word_dir = unpacked / "word"
        doc_xml_path = word_dir / "document.xml"
        if not doc_xml_path.exists():
            raise FileNotFoundError(f"document.xml not found in {word_dir}")

        # Create footnotes XML
        fn_xml = _make_footnotes_xml(footnotes)
        (word_dir / "footnotes.xml").write_text(fn_xml, encoding="utf-8")

        # Create footnotes relationships
        (word_dir / "_rels" / "footnotes.xml.rels").write_text(_make_footnotes_rels_xml(), encoding="utf-8")

        # Update [Content_Types].xml to include footnotes
        content_types_path = unpacked / "[Content_Types].xml"
        ct_xml = content_types_path.read_text(encoding="utf-8")
        if "footnotes" not in ct_xml:
            fn_ct = '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
            ct_xml = ct_xml.replace("</Types>", f"  {fn_ct}\n</Types>")
            content_types_path.write_text(ct_xml, encoding="utf-8")

        # Update word/_rels/document.xml.rels to include footnotes relationship
        doc_rels_path = word_dir / "_rels" / "document.xml.rels"
        rels_xml = doc_rels_path.read_text(encoding="utf-8")
        if "footnotes" not in rels_xml:
            fn_rel = '<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
            rels_xml = rels_xml.replace("</Relationships>", f"  {fn_rel}\n</Relationships>")
            doc_rels_path.write_text(rels_xml, encoding="utf-8")

        # Modify document.xml to insert footnote references
        doc_xml = doc_xml_path.read_text(encoding="utf-8")
        if footnote_positions:
            for idx, marker in enumerate(footnote_positions, 1):
                fn_ref_xml = (
                    f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                    f'<w:footnoteReference w:id="{idx}"/></w:r>'
                )
                # Replace the marker text in document body
                if marker in doc_xml:
                    doc_xml = doc_xml.replace(
                        f'<w:t xml:space="preserve">{marker}</w:t>',
                        f'<w:t xml:space="preserve">{marker}</w:t>{fn_ref_xml}'
                    )
                elif marker in doc_xml:
                    doc_xml = doc_xml.replace(
                        f'<w:t>{marker}</w:t>',
                        f'<w:t>{marker}</w:t>{fn_ref_xml}'
                    )

        doc_xml_path.write_text(doc_xml, encoding="utf-8")

        # Repack
        _, msg = pack(str(unpacked), str(Path(output_path).resolve()), validate=False)
        if msg.startswith("Error:"):
            raise RuntimeError(f"Pack failed: {msg}")

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    return output_path
