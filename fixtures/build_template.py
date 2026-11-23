"""Build a SYNTHETIC organizer-style academic PPTX template (master + layouts + theme).

NOT an official organizer template. It exists so the slide Skill can be developed and
self-tested against a real master/layout-based editable deck in this workspace.

Deliberately distinctive so that "the template was preserved" is an OBSERVABLE invariant:
  * theme name              = Organizer Academic 2026 (SYNTHETIC)
  * theme major/minor latin = Liberation Serif
  * theme accent1           = 1F6FEB  (not the python-pptx default 4472C4)
  * master carries a stamp text box carrying TEMPLATE_STAMP
  * every layout's footer placeholder carries TEMPLATE_FOOTER

Usage:  python build_template.py <out.pptx>
"""
from __future__ import annotations

import hashlib
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

THEME_NAME = "Organizer Academic 2026 (SYNTHETIC)"
MAJOR_LATIN = "Liberation Serif"
MINOR_LATIN = "Liberation Serif"
ACCENT1 = "1F6FEB"
TEMPLATE_STAMP = "ORGANIZER-ACADEMIC-2026"
TEMPLATE_FOOTER = "Organizer Academic 2026 - SYNTHETIC TEMPLATE"


def _q(tag: str) -> str:
    pfx, local = tag.split(":")
    return f"{{{A if pfx == 'a' else P}}}{local}"


def rewrite_theme(prs: Presentation) -> dict:
    """Set theme name, major/minor latin font and accent1. Returns before/after."""
    theme_part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)

    before = {}
    latin = root.findall(".//a:fontScheme/a:majorFont/a:latin", NS)
    if latin:
        before["major_latin"] = latin[0].get("typeface")
        latin[0].set("typeface", MAJOR_LATIN)
    latin = root.findall(".//a:fontScheme/a:minorFont/a:latin", NS)
    if latin:
        before["minor_latin"] = latin[0].get("typeface")
        latin[0].set("typeface", MINOR_LATIN)
    accent = root.findall(".//a:clrScheme/a:accent1/a:srgbClr", NS)
    if accent:
        before["accent1"] = accent[0].get("val")
        accent[0].set("val", ACCENT1)
    theme_el = root.find("a:theme", NS) if root.tag != _q("a:theme") else root
    if theme_el is None:
        theme_el = root
    before["name"] = theme_el.get("name")
    theme_el.set("name", THEME_NAME)

    theme_part._blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
    return before


def stamp_master(prs: Presentation) -> int:
    """Append a stamp text box to the slide master's shape tree (raw XML)."""
    master = prs.slide_master
    spTree = master.element.find(".//p:cSld/p:spTree", NS)
    if spTree is None:
        raise RuntimeError("master has no spTree")
    # unique shape id
    used = {int(e.get("id")) for e in spTree.iter(_q("p:cNvPr")) if e.get("id")}
    sid = max(used) + 1 if used else 2
    xml = f"""<p:sp xmlns:p="{P}" xmlns:a="{A}">
  <p:nvSpPr><p:cNvPr id="{sid}" name="TemplateStamp"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{Inches(0.35)}" y="{Inches(7.05)}"/>
    <a:ext cx="{Inches(6.0)}" cy="{Inches(0.3)}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="900" b="0"><a:solidFill><a:srgbClr val="{ACCENT1}"/></a:solidFill>
      <a:latin typeface="{MINOR_LATIN}"/></a:rPr><a:t>{TEMPLATE_STAMP}</a:t></a:r></a:p></p:txBody>
</p:sp>"""
    spTree.append(etree.fromstring(xml))
    return sid


def footer_layouts(prs: Presentation) -> list[str]:
    """Ensure each layout has a footer placeholder carrying TEMPLATE_FOOTER."""
    touched = []
    for layout in prs.slide_layouts:
        el = layout.element
        ftr = None
        for sp in el.iter(_q("p:sp")):
            ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
            if ph is not None and ph.get("type") == "ftr":
                ftr = sp
                break
        if ftr is None:
            # declare the footer in the layout's header/footer section
            # In p:sldLayout, <p:hf> is a SIBLING of <p:cSld> (after timing), not a
            # child of it -- putting it inside cSld violates the OOXML sequence.
            hf = el.find("p:hf", NS)
            if hf is None:
                hf = etree.SubElement(el, _q("p:hf"))
            hf.set("ftr", "1")
        else:
            for t in ftr.findall(".//a:t", NS):
                t.text = TEMPLATE_FOOTER
        touched.append(layout.name)
    return touched


def main() -> int:
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "template.pptx").resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    before = rewrite_theme(prs)
    sid = stamp_master(prs)
    layouts = footer_layouts(prs)
    prs.save(str(out))

    # --- verify by re-opening the written file ---
    chk = Presentation(str(out))
    theme_part = chk.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    got = {
        "name": root.get("name"),
        "major_latin": root.find(".//a:fontScheme/a:majorFont/a:latin", NS).get("typeface"),
        "minor_latin": root.find(".//a:fontScheme/a:minorFont/a:latin", NS).get("typeface"),
        "accent1": root.find(".//a:clrScheme/a:accent1/a:srgbClr", NS).get("val"),
    }
    stamp = [e for e in chk.slide_master.element.iter(_q("p:cNvPr"))
             if e.get("name") == "TemplateStamp"]
    sha = hashlib.sha256(out.read_bytes()).hexdigest()

    print(f"out          : {out}")
    print(f"bytes        : {out.stat().st_size}")
    print(f"sha256       : {sha}")
    print(f"theme before : {before}")
    print(f"theme after  : {got}")
    print(f"master stamp : id={sid} found_in_file={len(stamp)}")
    print(f"layouts      : {len(layouts)} -> {layouts}")
    ok = (got["name"] == THEME_NAME and got["major_latin"] == MAJOR_LATIN
          and got["minor_latin"] == MINOR_LATIN and got["accent1"] == ACCENT1
          and len(stamp) == 1 and len(layouts) >= 4)
    print("RESULT:", "OK" if ok else "MISMATCH")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
