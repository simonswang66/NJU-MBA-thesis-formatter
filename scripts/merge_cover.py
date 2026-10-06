"""
OOXML-level merge: insert a cover page template before the body content.

Uses DOM-based XML manipulation to avoid XML corruption.
unpack/pack 由本包 ooxml_zip.py 提供（自实现，无外部依赖）。

Usage:
    python merge_cover.py body.docx cover-template.docx output.docx
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path
from xml.dom.minidom import parse as dom_parse, getDOMImplementation

sys.path.insert(0, str(Path(__file__).resolve().parent))

from ooxml_zip import unpack, pack

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def merge_cover(body_path: str, cover_path: str, output_path: str) -> str:
    tmpdir = Path(tempfile.mkdtemp(prefix="docx_cover_"))
    body_dir = tmpdir / "body"
    cover_dir = tmpdir / "cover"
    merged_dir = tmpdir / "merged"

    try:
        _, msg = unpack(str(Path(body_path).resolve()), str(body_dir), merge_runs=False)
        if msg.startswith("Error:"):
            raise RuntimeError(f"Unpack body: {msg}")

        _, msg = unpack(str(Path(cover_path).resolve()), str(cover_dir), merge_runs=False)
        if msg.startswith("Error:"):
            raise RuntimeError(f"Unpack cover: {msg}")

        shutil.copytree(body_dir, merged_dir)

        word_body = merged_dir / "word"
        word_cover = cover_dir / "word"

        # Copy cover images
        body_media = word_body / "media"
        body_media.mkdir(exist_ok=True)
        image_map: dict[str, str] = {}
        cover_media = word_cover / "media"
        if cover_media.exists():
            for img in sorted(cover_media.iterdir()):
                if img.is_file():
                    new_name = f"cover_{img.name}"
                    shutil.copy2(img, body_media / new_name)
                    image_map[img.name] = new_name

        # DOM-based merge of document.xml
        body_dom = dom_parse(str(word_body / "document.xml"))
        cover_dom = dom_parse(str(word_cover / "document.xml"))

        body_elem = body_dom.documentElement
        cover_elem = cover_dom.documentElement

        # Find <w:body> in both
        w_body_tags = body_elem.getElementsByTagNameNS(W_NS, "body")
        w_cover_body_tags = cover_elem.getElementsByTagNameNS(W_NS, "body")
        if not w_body_tags or not w_cover_body_tags:
            raise ValueError("Missing w:body element")
        body_body = w_body_tags[0]
        cover_body = w_cover_body_tags[0]

        # Collect cover body children (capture w:sectPr for cover section)
        cover_children = []
        cover_sectPr = None
        for child in list(cover_body.childNodes):
            if child.nodeType == child.ELEMENT_NODE:
                if child.localName == "sectPr":
                    cover_sectPr = child
                    continue
                cover_children.append(child)

        # Strip trailing empty paragraphs from cover to avoid blank page
        while cover_children:
            last = cover_children[-1]
            if last.nodeType == last.ELEMENT_NODE and last.localName == "p":
                texts = last.getElementsByTagNameNS(W_NS, "t")
                if not texts or all(not (t.firstChild and t.firstChild.nodeValue.strip()) for t in texts):
                    cover_children.pop()
                    continue
            break

        # 封面自成一节（2026-10-06 装配工作流改）：封面 sectPr 去掉页眉/页脚/页码
        # 引用后包进收尾段落——封面独立成节、无页眉无页码，正文节互不受染。
        # （旧行为是封面与正文首节共享一节＋硬分页符，封面会带上正文页眉。）
        sect_p = body_dom.createElementNS(W_NS, "w:p")
        sect_pPr = body_dom.createElementNS(W_NS, "w:pPr")
        if cover_sectPr is not None:
            new_sectPr = body_dom.importNode(cover_sectPr, deep=True)
        else:
            new_sectPr = body_dom.createElementNS(W_NS, "w:sectPr")
        for tag in ("headerReference", "footerReference", "pgNumType"):
            for el in list(new_sectPr.getElementsByTagNameNS(W_NS, tag)):
                el.parentNode.removeChild(el)
        sect_pPr.appendChild(new_sectPr)
        sect_p.appendChild(sect_pPr)
        cover_children.append(sect_p)

        # Update image references in cover children
        for child in cover_children:
            for draw in child.getElementsByTagNameNS(W_NS, "drawing") + \
                        child.getElementsByTagNameNS(
                            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                            "inline"
                        ):
                for blip in draw.getElementsByTagNameNS(
                    "http://schemas.openxmlformats.org/drawingml/2006/main", "blip"
                ):
                    embed = blip.getAttribute("r:embed")
                    if embed:
                        pass  # Relationship IDs handled in _rels merge

            for rel_tag in child.getElementsByTagNameNS(RELS_NS, "*"):
                pass  # preserve as-is

        # Insert cover children at start of body
        first_body_child = body_body.firstChild
        for cover_child in cover_children:
            body_body.insertBefore(cover_child, first_body_child)

        # Write merged document.xml
        merged_xml = body_dom.toxml(encoding="utf-8").decode("utf-8")
        (word_body / "document.xml").write_text(merged_xml, encoding="utf-8")

        # Merge relationships - returns old_rId → new_rId mapping
        rid_map = _merge_rels(word_body / "_rels" / "document.xml.rels",
                              word_cover / "_rels" / "document.xml.rels",
                              image_map)

        # Fix r:embed references in merged document.xml
        _fix_rid_refs(word_body / "document.xml", rid_map)

        # Content types
        _merge_ct(merged_dir / "[Content_Types].xml",
                  cover_dir / "[Content_Types].xml")

        # Repack
        _, msg = pack(str(merged_dir), str(Path(output_path).resolve()), validate=False)
        if msg.startswith("Error:"):
            raise RuntimeError(f"Pack: {msg}")

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    return output_path


def _merge_rels(body_rels: Path, cover_rels: Path, image_map: dict) -> dict:
    """Merge cover rels into body rels. Returns {old_rId: new_rId} mapping."""
    rid_map = {}
    if not cover_rels.exists():
        return rid_map

    if body_rels.exists():
        body_dom = dom_parse(str(body_rels))
    else:
        impl = getDOMImplementation()
        body_dom = impl.createDocument(RELS_NS, "Relationships", None)

    cover_dom = dom_parse(str(cover_rels))
    body_root = body_dom.documentElement
    cover_root = cover_dom.documentElement

    # Find max rId
    max_id = 0
    for rel in body_root.getElementsByTagNameNS(RELS_NS, "Relationship"):
        rid = rel.getAttribute("Id")
        if rid.startswith("rId"):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass

    for rel in cover_root.getElementsByTagNameNS(RELS_NS, "Relationship"):
        target = rel.getAttribute("Target")
        old_rid = rel.getAttribute("Id")
        # Only merge image/media relationships
        if not target.startswith("media/"):
            continue
        img_name = target.replace("media/", "")
        new_target = f"media/{image_map[img_name]}" if img_name in image_map else target
        max_id += 1
        new_rid = f"rId{max_id}"
        rid_map[old_rid] = new_rid
        rel.setAttribute("Id", new_rid)
        rel.setAttribute("Target", new_target)
        body_root.appendChild(rel.cloneNode(deep=True))

    body_rels.parent.mkdir(parents=True, exist_ok=True)
    body_rels.write_text(body_dom.toxml(encoding="utf-8").decode("utf-8"), encoding="utf-8")
    return rid_map


def _fix_rid_refs(doc_xml_path: Path, rid_map: dict):
    """Replace old rId references in document.xml with mapped new rIds.

    三种引用形态都要改：DrawingML 的 r:embed、VML v:imagedata 的 r:id、
    外链图片的 r:link。2026-10-06 前只改 r:embed，封面校标（VML）r:id 未改，
    指向页脚部件、封面图实际断裂（由 docx_postprocess 清理"未引用"图时暴露）。
    """
    if not rid_map:
        return
    xml_text = doc_xml_path.read_text(encoding="utf-8")
    for old_rid, new_rid in rid_map.items():
        xml_text = xml_text.replace(f'r:embed="{old_rid}"', f'r:embed="{new_rid}"')
        xml_text = xml_text.replace(f'r:id="{old_rid}"', f'r:id="{new_rid}"')
        xml_text = xml_text.replace(f'r:link="{old_rid}"', f'r:link="{new_rid}"')
    doc_xml_path.write_text(xml_text, encoding="utf-8")


def _merge_ct(body_ct: Path, cover_ct: Path):
    if not cover_ct.exists() or not body_ct.exists():
        return
    body_dom = dom_parse(str(body_ct))
    cover_dom = dom_parse(str(cover_ct))
    body_root = body_dom.documentElement
    cover_root = cover_dom.documentElement
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    existing_exts = set()
    for d in body_root.getElementsByTagNameNS(ct_ns, "Default"):
        existing_exts.add(d.getAttribute("Extension"))

    # Copy Default entries for image extensions not already present
    for d in cover_root.getElementsByTagNameNS(ct_ns, "Default"):
        ext = d.getAttribute("Extension")
        ct_val = d.getAttribute("ContentType")
        if ext not in existing_exts and ct_val and "image" in ct_val.lower():
            body_root.appendChild(d.cloneNode(deep=True))
            existing_exts.add(ext)

    body_ct.write_text(body_dom.toxml(encoding="utf-8").decode("utf-8"), encoding="utf-8")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python merge_cover.py body.docx cover-template.docx output.docx")
        sys.exit(1)
    result = merge_cover(sys.argv[1], sys.argv[2], sys.argv[3])
    print(f"Merged: {result}")
