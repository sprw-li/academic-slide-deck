#!/usr/bin/env python3
"""build_deck.py — render ``plan.json`` into an editable ``deck.pptx`` + report.

Contract: ``work/CONTRACT.md`` §3.  Usage::

    python build_deck.py --case <case.json> --work <workdir> [--in plan.json]
                         [--out deck.pptx] [--report build_report.json]

Exit codes: ``0`` deck + report written, ``2`` usage/IO error.

How the template is respected (and proved)
------------------------------------------
The deck is built **from the case's own template** (``Presentation(template)``) or
from ``starting_deck_path`` when the case provides one, using **only** the layouts
that already exist in it.  Slides are added with ``slides.add_slide(layout)``; only
placeholders are filled (title, body, picture); footer / date / slide-number
placeholders are cloned from the layout and left **unfilled**.

python-pptx re-serialises every XML part it touches, so semantic equality is not
enough for the contract: after ``save()`` this module **restores the master,
layout and theme parts byte-for-byte from the template package**
(``_restore_template_parts``).  ``invariants`` is then computed by comparing the
two packages part by part, so ``master_unchanged``/``theme_unchanged`` are
byte-level facts, not hopes.  Parts that python-pptx legitimately adds (e.g. a
notes master) are reported separately in ``parts_added``.

Figures
-------
A figure is embedded only when ``material.json`` carries a **real** image for it
(crop captured by ``ingest.py``).  Mode is ``source-crop`` for a source image and
``redraw-from-data`` for a chart drawn by matplotlib **from the real cells of a
source table**.  A figure with no image is skipped with a warning — never a
placeholder.
"""
from __future__ import annotations

import sys as _sys

_sys.dont_write_bytecode = True

import argparse
import hashlib
import json
import math
import os
import re
import xml.etree.ElementTree as ET
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in _sys.path:
    _sys.path.insert(0, HERE)

import common as C  # noqa: E402

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

TITLE_TYPES = {"CENTER_TITLE", "TITLE", "VERTICAL_TITLE"}
BODY_TYPES = {"BODY", "OBJECT", "VERTICAL_BODY", "VERTICAL_OBJECT", "SUBTITLE"}
PICTURE_TYPES = {"PICTURE", "OBJECT", "CONTENT"}
FURNITURE_TYPES = {"DATE", "FOOTER", "SLIDE_NUMBER"}

DEFAULT_TITLE_PT = 28.0
DEFAULT_BODY_PT = 18.0

#: part directories that must stay byte-identical to the template
PROTECTED_PREFIXES = ("ppt/slideMasters/", "ppt/slideLayouts/", "ppt/theme/")


# --------------------------------------------------------------------------- #
# package helpers
# --------------------------------------------------------------------------- #
def zip_parts(path, prefixes=None):
    """``{part_name: sha256}`` for the entries of a zip (optionally filtered)."""
    out = {}
    try:
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                name = info.filename
                if name.endswith("/"):
                    continue
                if prefixes and not name.startswith(prefixes):
                    continue
                out[name] = C.sha256_bytes(zf.read(name))
    except (OSError, zipfile.BadZipFile):
        return out
    return out


def _restore_template_parts(out_path, template_path, warnings):
    """Restore master/layout/theme parts of ``out_path`` from the template bytes.

    python-pptx re-serialises parsed parts, which can change bytes without changing
    meaning.  The contract asks for the template's masters/layouts/theme to be
    untouched, so the original bytes are put back.  Returns the number of parts
    restored.
    """
    if not template_path or not os.path.isfile(template_path):
        return 0
    try:
        with zipfile.ZipFile(template_path) as src:
            originals = {i.filename: src.read(i.filename) for i in src.infolist()
                         if not i.filename.endswith("/")
                         and i.filename.startswith(PROTECTED_PREFIXES)}
    except (OSError, zipfile.BadZipFile) as exc:
        warnings.append("could not read the template package to restore protected "
                        "parts (%s)" % exc)
        return 0
    if not originals:
        return 0
    tmp = out_path + ".restore.tmp"
    restored = 0
    try:
        with zipfile.ZipFile(out_path) as src:
            infos = src.infolist()
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
                for info in infos:
                    name = info.filename
                    if name.endswith("/"):
                        continue
                    data = originals.get(name)
                    if data is not None:
                        restored += 1
                    else:
                        data = src.read(name)
                    dst.writestr(name, data)
        os.replace(tmp, out_path)
    except (OSError, zipfile.BadZipFile) as exc:
        warnings.append("protected-part restore failed (%s: %s); master/theme equality "
                        "is then reported by comparison only" % (type(exc).__name__, exc))
        if os.path.isfile(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        return 0
    return restored


def read_theme_fonts(template_path, warnings):
    """``{major_latin, minor_latin, major_ea, minor_ea}`` from ``ppt/theme/*.xml``."""
    out = {"major_latin": "", "minor_latin": "", "major_ea": "", "minor_ea": ""}
    if not template_path or not os.path.isfile(template_path):
        return out
    try:
        with zipfile.ZipFile(template_path) as zf:
            names = [n for n in zf.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n)]
            if not names:
                return out
            root = ET.fromstring(zf.read(sorted(names)[0]))
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        warnings.append("could not read theme fonts (%s: %s)" % (type(exc).__name__, exc))
        return out
    for tag, prefix in (("majorFont", "major"), ("minorFont", "minor")):
        node = root.find(".//{%s}%s" % (A_NS, tag))
        if node is None:
            continue
        latin = node.find("{%s}latin" % A_NS)
        ea = node.find("{%s}ea" % A_NS)
        if latin is not None:
            out["%s_latin" % prefix] = latin.get("typeface") or ""
        if ea is not None:
            out["%s_ea" % prefix] = ea.get("typeface") or ""
    return out


def read_fonts_in_parts(path, prefixes=("ppt/slides/",), warnings=None):
    """Explicit ``a:latin`` typefaces written into the deck's own slide parts."""
    fonts = set()
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if not name.startswith(prefixes) or not name.endswith(".xml"):
                    continue
                try:
                    root = ET.fromstring(zf.read(name))
                except ET.ParseError:
                    continue
                for node in root.iter("{%s}latin" % A_NS):
                    value = node.get("typeface")
                    if value:
                        fonts.add(value)
    except (OSError, zipfile.BadZipFile) as exc:
        if warnings is not None:
            warnings.append("could not scan deck fonts (%s)" % exc)
    return sorted(fonts)


# --------------------------------------------------------------------------- #
# text fitting
# --------------------------------------------------------------------------- #
def _emu_to_pt(value):
    return float(value) / 12700.0


def estimate_overflow(shape, paragraphs, font_size_pt, shrink=0.0):
    """Engine-level text-fit estimate (no renderer needed).

    Returns ``(risk, detail)`` where detail records the estimated line count, the
    needed height and the available height in points.
    """
    width_pt = _emu_to_pt(shape.width)
    height_pt = _emu_to_pt(shape.height) * (1.0 - max(0.0, min(shrink, 0.8)))
    if width_pt <= 0 or height_pt <= 0:
        return "high", {"needed_pt": 0.0, "available_pt": 0.0, "est_lines": 0,
                        "reason": "shape has no measurable extent"}
    size = font_size_pt or DEFAULT_BODY_PT
    chars_per_line = max(8, int(width_pt / (0.52 * size)))
    lines = 0
    for text in paragraphs:
        lines += max(1, int(math.ceil(len(str(text)) / float(chars_per_line))))
    needed = lines * size * 1.35 + 4.0 * max(0, len(paragraphs) - 1)
    ratio = needed / height_pt
    risk = "low" if ratio <= 0.80 else ("medium" if ratio <= 1.0 else "high")
    return risk, {"needed_pt": round(needed, 1), "available_pt": round(height_pt, 1),
                  "est_lines": lines, "chars_per_line": chars_per_line,
                  "font_pt": size}


def _placeholder_type(placeholder):
    try:
        kind = placeholder.placeholder_format.type
    except Exception:
        return ""
    return str(kind).split(" ")[0]


def _size_from_element(element):
    """Explicit font size (pt) declared anywhere in a drawing element, or ``None``."""
    for tag in ("defRPr", "rPr", "endParaRPr"):
        for node in element.iter("{%s}%s" % (A_NS, tag)):
            size = node.get("sz")
            if not size:
                continue
            try:
                return int(size) / 100.0
            except ValueError:
                continue
    return None


def _size_from_txstyles(master_element, kind):
    """Title/body size from the master's ``p:txStyles`` (where PowerPoint keeps it).

    The slide master's placeholder often declares no size at all: the real inherited
    size for the title slide lives in ``p:txStyles/p:titleStyle``.  Missing this made
    the fit estimate assume 28 pt while the deck was rendering 44 pt, so an
    over-long title silently grew past its box and collided with the subtitle.
    """
    style_tag = {"TITLE": "titleStyle", "CENTER_TITLE": "titleStyle",
                 "VERTICAL_TITLE": "titleStyle", "BODY": "bodyStyle",
                 "OBJECT": "bodyStyle", "SUBTITLE": "bodyStyle",
                 "VERTICAL_BODY": "bodyStyle"}.get(kind, "otherStyle")
    for styles in master_element.iter("{%s}txStyles" % P_NS):
        for style in styles:
            if not style.tag.endswith("}" + style_tag):
                continue
            size = _size_from_element(style)
            if size:
                return size
    return None


def resolve_font_size(slide, shape, layout, default_pt):
    """Effective font size for a fit estimate, from the real inheritance chain.

    Order: an explicit size on the shape's own runs → the layout placeholder with
    the same ``idx`` → the master placeholder of the same type → the master's
    ``p:txStyles`` → ``default_pt``.
    """
    try:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.size is not None:
                    return run.font.size.pt
            if paragraph.font.size is not None:
                return paragraph.font.size.pt
    except Exception:
        pass
    kind = _placeholder_type(shape)
    try:
        if shape.is_placeholder:
            idx = shape.placeholder_format.idx
            for placeholder in layout.placeholders:
                if placeholder.placeholder_format.idx == idx:
                    size = _size_from_element(placeholder._element)
                    if size:
                        return size
        for placeholder in slide.slide_layout.slide_master.placeholders:
            if _placeholder_type(placeholder) == kind:
                size = _size_from_element(placeholder._element)
                if size:
                    return size
        size = _size_from_txstyles(slide.slide_layout.slide_master._element, kind)
        if size:
            return size
    except Exception:
        pass
    return default_pt


# --------------------------------------------------------------------------- #
# layout / placeholder plumbing
# --------------------------------------------------------------------------- #
def find_layout(prs, name, hint, warnings):
    """A layout from the presentation by exact name, then by hint, then the first."""
    layouts = list(prs.slide_layouts)
    if name:
        for layout in layouts:
            if layout.name == name:
                return layout, "plan.layout_name"
    normalised = re.sub(r"[^a-z0-9]+", " ", (hint or "").lower()).strip()
    if normalised:
        for layout in layouts:
            other = re.sub(r"[^a-z0-9]+", " ", layout.name.lower()).strip()
            if normalised and (normalised in other or other in normalised):
                return layout, "hint-match"
    for layout in layouts:
        if "content" in layout.name.lower() and "only" not in layout.name.lower():
            warnings.append("layout %r for hint %r was not found; fell back to %r"
                            % (name, hint, layout.name))
            return layout, "fallback-content"
    if layouts:
        warnings.append("layout %r for hint %r was not found; fell back to the first "
                        "layout %r" % (name, hint, layouts[0].name))
        return layouts[0], "fallback-first"
    return None, "none"


def body_rect(slide, layout):
    """Rectangle (EMU) for content: the layout's body placeholder, else slide-based."""
    for placeholder in list(slide.placeholders) + list(layout.placeholders):
        kind = _placeholder_type(placeholder)
        if kind in ("BODY", "OBJECT", "VERTICAL_BODY", "VERTICAL_OBJECT"):
            return (placeholder.left, placeholder.top, placeholder.width,
                    placeholder.height), placeholder.name
    width = slide.part.package.presentation_part.presentation.slide_width
    height = slide.part.package.presentation_part.presentation.slide_height
    return (int(width * 0.08), int(height * 0.26), int(width * 0.84),
            int(height * 0.62)), None


def _enable_norm_autofit(shape):
    """Declare ``<a:normAutofit/>`` on a text frame (PowerPoint "shrink on overflow").

    This is the template-compatible way to promise that over-long text will be
    scaled down instead of clipped.  It is additive and slide-local: no master,
    layout or theme part is touched.  ``fontScale`` is left at 100 % because the
    fitted sizes are written explicitly onto the runs, so the renderer has nothing
    left to scale — and if its metrics differ slightly from this estimate it shrinks
    rather than clips.
    """
    from pptx.oxml.ns import qn  # noqa: WPS433
    try:
        body = shape.text_frame._txBody.find(qn("a:bodyPr"))
    except Exception:
        return False
    if body is None:
        return False
    for tag in ("a:normAutofit", "a:spAutoFit", "a:noAutofit"):
        for child in body.findall(qn(tag)):
            body.remove(child)
    body.append(body.makeelement(qn("a:normAutofit"), {}))
    return True


def apply_shrink_to_fit(shape, font_size_pt, floor=0.62, line_factor=1.22,
                        safety=0.92):
    """Fit an over-long text frame to its box (explicit size + ``normAutofit``).

    The template's boxes are sized for the *template's* own prompts; a real paper
    title can be twice as long, and oversized text grows past its box and collides
    with the next placeholder.  The frame gets ``normAutofit`` and, when the estimate
    says it would still overflow, every run's size is reduced to the largest size
    that fits — never below ``floor ×`` the inherited size.  Fonts, theme and layout
    stay untouched, and the adjustment is recorded in ``build_report.json``.
    """
    try:
        text = shape.text_frame.text or ""
    except Exception:
        return None
    if not text.strip():
        return None
    declared = _enable_norm_autofit(shape)
    width_pt = _emu_to_pt(shape.width)
    height_pt = _emu_to_pt(shape.height)
    if width_pt <= 0 or height_pt <= 0 or font_size_pt <= 0:
        return None
    paragraphs = text.split("\n")

    def needed_for(size):
        chars_per_line = max(8, int(width_pt / (0.50 * size)))
        lines = sum(max(1, int(math.ceil(len(p) / float(chars_per_line))))
                    for p in paragraphs)
        return lines * size * line_factor

    budget = height_pt * safety
    size = float(font_size_pt)
    while size > font_size_pt * floor and needed_for(size) > budget:
        size -= 1.0
    if size >= font_size_pt - 0.5:
        return {"shape": shape.name, "from_pt": round(font_size_pt, 1),
                "to_pt": round(font_size_pt, 1), "runs": 0,
                "autofit": bool(declared),
                "reason": "normAutofit declared; inherited size already fits"}
    from pptx.util import Pt  # noqa: WPS433
    applied = 0
    for paragraph in shape.text_frame.paragraphs:
        paragraph.font.size = Pt(round(size, 1))
        for run in paragraph.runs:
            run.font.size = Pt(round(size, 1))
            applied += 1
    return {"shape": shape.name, "from_pt": round(font_size_pt, 1),
            "to_pt": round(size, 1), "runs": applied, "autofit": bool(declared),
            "reason": "%.0fpt of text would need %.0fpt of a %.0fpt box"
                      % (font_size_pt, needed_for(font_size_pt), height_pt)}


def ensure_furniture(slide, layout, warnings):
    """Clone footer/date/slide-number placeholders so they survive, unfilled.

    ``python-pptx`` deliberately does not clone these onto a new slide (they are
    display furniture inherited from the layout), so they are cloned explicitly
    with ``SlideShapes.clone_placeholder`` and then emptied: an empty placeholder
    is exactly how PowerPoint represents "inherit the layout footer", whereas a
    literal copied string would freeze one slide's footer and could even show the
    slide-number field placeholder text.
    """
    cloned, present = 0, []
    clone = getattr(slide.shapes, "clone_placeholder", None) or \
        getattr(slide.shapes, "clone_layout_placeholder", None)
    for placeholder in layout.placeholders:
        kind = _placeholder_type(placeholder)
        if kind not in FURNITURE_TYPES:
            continue
        present.append(kind)
        if any(_placeholder_type(p) == kind for p in slide.placeholders):
            continue
        if clone is None:
            continue
        try:
            clone(placeholder)
            cloned += 1
        except Exception as exc:
            warnings.append("could not clone the %s placeholder onto a slide (%s: %s)"
                            % (kind, type(exc).__name__, exc))
    return cloned, present


def count_furniture_in_package(path, prefix="ppt/slides/slide"):
    """Count slides that carry a footer/slide-number placeholder, from saved OOXML.

    The DrawingML placeholder types are ``ftr`` (footer), ``sldNum``/``sldnum``
    (slide number) and ``dt`` (date); accept any spelling/case so the check cannot
    silently read "no footer" out of a correctly written deck.
    """
    total, with_footer = 0, 0
    try:
        with zipfile.ZipFile(path) as zf:
            for name in zf.namelist():
                if not name.endswith(".xml") or not re.match(
                        r"ppt/slides/slide\d+\.xml$", name):
                    continue
                total += 1
                xml = zf.read(name).decode("utf-8", "replace")
                if re.search(r'<p:ph[^>]*type="(?:ftr|footer|sldNum|sldnum)"', xml,
                             re.IGNORECASE):
                    with_footer += 1
    except (OSError, zipfile.BadZipFile):
        return 0, 0
    return total, with_footer


def clear_furniture_text(slide):
    """Guarantee we never fill a footer/date/slide-number placeholder ourselves."""
    cleared = 0
    for placeholder in slide.placeholders:
        if _placeholder_type(placeholder) in FURNITURE_TYPES:
            try:
                if placeholder.has_text_frame and placeholder.text_frame.text:
                    placeholder.text_frame.text = ""
                    cleared += 1
            except Exception:
                pass
    return cleared


def fill_body(text_frame, paragraphs):
    """Write ``paragraphs`` into a placeholder body, first paragraph plain, rest bullets."""
    text_frame.clear()
    for index, text in enumerate(paragraphs):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.text = text
        if index > 0:
            paragraph.level = 0
    return len(paragraphs)


# --------------------------------------------------------------------------- #
# figure / table embedding
# --------------------------------------------------------------------------- #
def redraw_table_chart(cells, out_png, title, warnings, key):
    """Draw a bar chart **from the supplied cells** (never invented data)."""
    matplotlib = C.try_import("matplotlib")
    if matplotlib is None:
        return None
    try:
        import numpy as np
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        header = [str(c) for c in cells[0]]
        rows = cells[1:]
        labels, series = [], {}
        for row in rows:
            if not row:
                continue
            labels.append(str(row[0]))
            for col in range(1, len(row)):
                try:
                    value = float(str(row[col]).replace(",", "").strip())
                except ValueError:
                    continue
                series.setdefault(col, []).append(value)
        if not labels or not series:
            return None
        usable = [col for col, values in series.items() if len(values) == len(labels)]
        if not usable:
            return None
        figure, axis = plt.subplots(figsize=(7.0, 3.6), dpi=160)
        width = 0.8 / len(usable)
        positions = np.arange(len(labels))
        for offset, col in enumerate(usable):
            name = header[col] if col < len(header) else "column %d" % col
            axis.bar(positions + offset * width, series[col], width, label=name)
        axis.set_xticks(positions + 0.4 - width / 2)
        axis.set_xticklabels(labels, rotation=0, fontsize=8)
        axis.set_title(title or "Source table", fontsize=10)
        axis.legend(fontsize=8)
        axis.tick_params(axis="y", labelsize=8)
        figure.tight_layout()
        C.ensure_dir(os.path.dirname(os.path.abspath(out_png)))
        figure.savefig(out_png)
        plt.close(figure)
        return out_png
    except Exception as exc:
        warnings.append("%s: matplotlib redraw failed (%s: %s)"
                        % (key, type(exc).__name__, exc))
        return None


def add_native_table(slide, cells, rect, max_rows=8, max_cols=6):
    """Insert a real, editable PowerPoint table from the source cells."""
    rows = [row[:max_cols] for row in cells[:max_rows]]
    if not rows:
        return None
    shape = slide.shapes.add_table(len(rows), max(len(r) for r in rows),
                                   rect[0], rect[1], rect[2], rect[3])
    table = shape.table
    for r, row in enumerate(rows):
        for c in range(len(table.columns)):
            text = str(row[c]) if c < len(row) else ""
            cell = table.cell(r, c)
            cell.text = text
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = run.font.size or None
    return table


# --------------------------------------------------------------------------- #
# orchestration
# --------------------------------------------------------------------------- #
def _material_images(material):
    return {f.get("id"): f for f in material.get("figures", [])}


def _material_tables(material):
    return {t.get("id"): t for t in material.get("tables", [])}


def build_deck_stdlib(case_path, work_dir, in_path=None, out_path=None, log=None):
    """Build the deck **without python-pptx** (``zipfile`` + ElementTree only).

    The template package is cloned, so masters/layouts/theme/footer are preserved by
    construction and the invariants are byte-level facts again.  Fidelity differs from
    the python-pptx path (no autofit metrics, tables become text lines), and the
    report says so via ``engine: "stdlib-fallback"``.
    """
    import pptx_stdlib  # noqa: WPS433  (sibling module, stdlib-only)

    case = C.load_case(case_path)
    case_id = C.extract_case_id(case, case_path)
    plan = C.read_json(in_path or os.path.join(work_dir, "plan.json"), None)
    if not isinstance(plan, dict) or not plan.get("slides"):
        raise RuntimeError("plan.json not found or has no slides: %s"
                           % (in_path or os.path.join(work_dir, "plan.json")))
    material = C.read_json(os.path.join(work_dir, "material.json"), {}) or {}
    warnings, uncertainty = [], []
    uncertainty.append("python-pptx is not importable in this environment; the deck was "
                       "written by the pure-stdlib fallback (engine=stdlib-fallback), "
                       "which preserves the template package but does not replicate "
                       "python-pptx's full layout fidelity")

    constraints = plan.get("constraints") or {}
    template_path = constraints.get("template_path") or case.get("template")
    if template_path and not os.path.isabs(template_path):
        template_path = os.path.join(os.path.dirname(os.path.abspath(case_path)),
                                     template_path)
    if template_path and not os.path.isfile(template_path):
        warnings.append("declared template %r does not exist" % template_path)
        template_path = None

    C.ensure_dir(os.path.dirname(os.path.abspath(out_path)))
    result = pptx_stdlib.build(template_path, out_path, plan.get("slides") or [],
                               work_dir, log=log)
    warnings.extend(result.get("warnings") or [])

    tpl_path = template_path if result.get("template_cloned") else None
    tpl_parts = zip_parts(tpl_path, PROTECTED_PREFIXES) if tpl_path else {}
    out_parts = zip_parts(out_path, PROTECTED_PREFIXES)
    missing = sorted(set(tpl_parts) - set(out_parts))
    changed = sorted(name for name in tpl_parts
                     if name in out_parts and tpl_parts[name] != out_parts[name])
    added = sorted(set(out_parts) - set(tpl_parts))
    master_unchanged = bool(tpl_parts) and not any(
        name.startswith("ppt/slideMasters/") for name in changed + missing)
    theme_unchanged = bool(tpl_parts) and not any(
        name.startswith("ppt/theme/") for name in changed + missing)
    if not tpl_parts:
        uncertainty.append("no template package was cloned, so master/theme equality "
                           "cannot be proven for this run")
    slide_total, slides_with_footer = count_furniture_in_package(out_path)
    theme_fonts = read_theme_fonts(tpl_path, warnings) if tpl_path else {}

    report_slides = []
    for position, plan_slide in enumerate(plan.get("slides") or [], start=1):
        text = " ".join([plan_slide.get("title") or ""]
                        + list(plan_slide.get("bullets") or []))
        chars = len(text)
        risk = "high" if chars > 700 else ("medium" if chars > 420 else "low")
        images = []
        figure_id = plan_slide.get("figure_ref")
        if figure_id and os.path.isfile(os.path.join(work_dir, "figures",
                                                     "%s.png" % figure_id)):
            images.append({"source_ref": figure_id, "mode": "source-crop",
                          "path": "figures/%s.png" % figure_id,
                          "sha256": C.sha256_file(os.path.join(work_dir, "figures",
                                                               "%s.png" % figure_id)),
                          "backend": "captured by ingest.py"})
        report_slides.append({
            "index": position,
            "layout": (constraints.get("layout_hints_available", {})
                       .get(plan_slide.get("layout_hint") or "", {})
                       .get("layout_name") or plan_slide.get("layout_name")
                       or plan_slide.get("layout_hint") or ""),
            "placeholders_filled": (["title"] + (["body"] if plan_slide.get("bullets")
                                                 or plan_slide.get("claim") else [])),
            "overflow_risk": risk,
            "overflow_detail": [{"chars": chars,
                                 "risk_source": "character-count heuristic "
                                                "(no python-pptx in this environment)"}],
            "shrink_to_fit": [],
            "autofit_frames": 0,
            "notes_chars": len(plan_slide.get("notes") or ""),
            "furniture_cloned": 0,
            "furniture_present": ["FOOTER", "SLIDE_NUMBER", "DATE"],
            "furniture_text_cleared": 0,
            "images": images,
        })

    report = {
        "schema": C.SCHEMA_BUILD_REPORT,
        "case_id": case_id,
        "generated": C.now_iso(),
        "engine": "stdlib-fallback",
        "engine_note": ("python-pptx unavailable; deck written with zipfile + "
                        "ElementTree from the cloned template package"),
        "mode": "built-from-template" if result.get("template_cloned")
        else "built-minimal-package",
        "pptx": os.path.abspath(out_path),
        "pptx_rel": C.relpath(out_path, work_dir).replace("\\", "/"),
        "sha256": C.sha256_file(out_path),
        "bytes": os.path.getsize(out_path) if os.path.isfile(out_path) else 0,
        "template": {
            "path": os.path.abspath(tpl_path) if tpl_path else "",
            "path_rel": C.relpath(tpl_path, work_dir).replace("\\", "/") if tpl_path
            else "",
            "sha256": C.sha256_file(tpl_path) if tpl_path else "",
            "masters": len([n for n in out_parts
                            if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]),
            "layouts": len([n for n in out_parts
                            if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)]),
            "layout_names": result.get("layouts_available") or [],
            "theme_fonts": theme_fonts,
        },
        "invariants": {
            "master_unchanged": bool(master_unchanged),
            "theme_unchanged": bool(theme_unchanged),
            "footer_preserved": bool(slide_total) and slides_with_footer == slide_total,
            "layouts_added": len([n for n in added
                                  if n.startswith("ppt/slideLayouts/")]),
            "layouts_deleted": len([n for n in missing
                                    if n.startswith("ppt/slideLayouts/")]),
            "layouts_changed": len([n for n in changed
                                    if n.startswith("ppt/slideLayouts/")]),
            "footer_placeholders": {"slides": slide_total,
                                    "slides_with_footer": slides_with_footer},
            "protected_parts_restored": 0,
            "protected_parts_changed": changed,
            "protected_parts_missing": missing,
            "parts_added": added,
        },
        "slides": report_slides,
        "fonts_used": sorted({value for value in (theme_fonts or {}).values() if value}),
        "output": {"parts": result.get("parts"), "notes_slides": result.get("notes"),
                   "embedded_images": result.get("images")},
        "warnings": warnings,
        "uncertainty": uncertainty,
        "stats": {"slides": len(report_slides),
                  "images": sum(len(s["images"]) for s in report_slides),
                  "source_crops": sum(len(s["images"]) for s in report_slides),
                  "redraws": 0,
                  "high_overflow": sum(1 for s in report_slides
                                       if s["overflow_risk"] == "high"),
                  "warnings": len(warnings)},
    }
    return report


def build_deck(case_path, work_dir, in_path=None, out_path=None, log=None):
    pptx = C.try_import("pptx")
    if pptx is None:
        # python-pptx is not among the challenge's declared resources: the evaluated
        # container may not have it.  A deck written with zipfile+ElementTree is far
        # better than no deck, so fall back to pptx_stdlib instead of failing.
        return build_deck_stdlib(case_path, work_dir, in_path=in_path,
                                 out_path=out_path, log=log)
    from pptx.util import Inches  # noqa: WPS433

    case = C.load_case(case_path)
    case_id = C.extract_case_id(case, case_path)
    plan = C.read_json(in_path or os.path.join(work_dir, "plan.json"), None)
    if not isinstance(plan, dict) or not plan.get("slides"):
        raise RuntimeError("plan.json not found or has no slides: %s"
                           % (in_path or os.path.join(work_dir, "plan.json")))
    material = C.read_json(os.path.join(work_dir, "material.json"), {}) or {}
    warnings, uncertainty = [], []
    if material.get("warnings"):
        uncertainty.append("material.json carries %d extraction warning(s); figures "
                           "without a captured image are intentionally not embedded"
                           % len(material["warnings"]))

    constraints = plan.get("constraints") or {}
    template_path = constraints.get("template_path") or case.get("template")
    if template_path and not os.path.isabs(template_path):
        template_path = os.path.join(os.path.dirname(os.path.abspath(case_path)),
                                     template_path)
    if template_path and not os.path.isfile(template_path):
        warnings.append("declared template %r does not exist" % template_path)
        template_path = None
    starting_deck = constraints.get("starting_deck_path") or case.get("starting_deck")
    if starting_deck and not os.path.isabs(starting_deck):
        starting_deck = os.path.join(os.path.dirname(os.path.abspath(case_path)),
                                     starting_deck)
    if starting_deck and not os.path.isfile(starting_deck):
        warnings.append("declared starting deck %r does not exist" % starting_deck)
        starting_deck = None

    source_for_open = starting_deck or template_path
    if source_for_open:
        prs = pptx.Presentation(source_for_open)
    else:
        prs = pptx.Presentation()
        uncertainty.append("the case declares neither a template nor a starting deck; "
                           "the python-pptx default template was used and the deck does "
                           "NOT respect an organizer master")
        warnings.append("no template available; built on the python-pptx default template")
    mode = "adapted-from-starting-deck" if starting_deck else "built-from-template"

    # template baseline (for invariants) — the template, not the starting deck
    tpl_path = template_path or source_for_open
    tpl_parts = zip_parts(tpl_path, PROTECTED_PREFIXES) if tpl_path else {}
    tpl_layout_names = [layout.name for layout in prs.slide_layouts]
    theme_fonts = read_theme_fonts(tpl_path, warnings)

    images = _material_images(material)
    tables = _material_tables(material)
    existing = len(prs.slides)
    if mode == "adapted-from-starting-deck" and existing:
        uncertainty.append("the starting deck already contains %d slide(s); the first %d "
                           "were rewritten from the plan and the rest were left as "
                           "supplied" % (existing, min(existing, len(plan["slides"]))))
    if mode == "adapted-from-starting-deck" and existing > len(plan["slides"]):
        warnings.append("the starting deck has %d slide(s) but the plan has %d; the "
                        "extra slide(s) were left untouched" % (existing,
                                                               len(plan["slides"])))

    report_slides, derived_dir = [], os.path.join(work_dir, "derived")
    fonts_used = set(v for v in theme_fonts.values() if v)
    furniture_cloned_total = 0

    for position, plan_slide in enumerate(plan["slides"]):
        if position < existing:
            slide = prs.slides[position]
            layout = slide.slide_layout
        else:
            layout, why = find_layout(prs, plan_slide.get("layout_name"),
                                      plan_slide.get("layout_hint"), warnings)
            if layout is None:
                raise RuntimeError("the presentation exposes no slide layout at all")
            slide = prs.slides.add_slide(layout)
        cloned, furniture_present = ensure_furniture(slide, layout, warnings)
        furniture_cloned_total += cloned

        filled, images_report = [], []
        title_ok = False
        for placeholder in list(slide.placeholders):
            kind = _placeholder_type(placeholder)
            if kind in TITLE_TYPES and not title_ok:
                placeholder.text_frame.text = plan_slide.get("title") or ""
                filled.append("title")
                title_ok = True
        body_shape, body_name = None, None
        for placeholder in list(slide.placeholders):
            kind = _placeholder_type(placeholder)
            if kind in ("BODY", "OBJECT", "VERTICAL_BODY", "VERTICAL_OBJECT",
                        "SUBTITLE"):
                body_shape, body_name = placeholder, placeholder.name
                break

        paragraphs = list(plan_slide.get("bullets") or [])
        if not paragraphs and (plan_slide.get("role") or "") != "figure":
            # A figure page's caption box is small; the claim lives in the plan,
            # the evidence map and the speaker notes.  Dumping a full caption
            # sentence there would clip it (V10), so the place is left to the image.
            paragraphs = [plan_slide.get("claim") or ""]
        if body_shape is not None:
            fill_body(body_shape.text_frame, paragraphs)
            filled.append(body_name or "body")
        else:
            rect, layout_body_name = body_rect(slide, layout)
            box = slide.shapes.add_textbox(rect[0], rect[1], rect[2], rect[3])
            fill_body(box.text_frame, paragraphs)
            box.text_frame.word_wrap = True
            filled.append("textbox (layout has no content placeholder%s)"
                          % ("" if layout_body_name is None
                             else "; layout body was %r" % layout_body_name))
            warnings.append("slide %d: the chosen layout %r exposes no content "
                            "placeholder; a textbox was added inside the content "
                            "rectangle" % (position + 1, layout.name))

        # ---- figures ------------------------------------------------------
        figure_id = plan_slide.get("figure_ref")
        if figure_id and figure_id in images:
            figure = images[figure_id]
            image = figure.get("image") or {}
            rel = image.get("path")
            full = os.path.join(work_dir, rel) if rel else None
            if full and os.path.isfile(full):
                digest = C.sha256_file(full)
                if image.get("sha256") and image["sha256"] != digest:
                    warnings.append("figure %s image changed on disk since ingest; "
                                    "embedding the file that is present now" % figure_id)
                rect, _name = body_rect(slide, layout)
                target_rect = rect
                picture_placeholder = None
                for placeholder in slide.placeholders:
                    if _placeholder_type(placeholder) == "PICTURE":
                        picture_placeholder = placeholder
                        break
                if picture_placeholder is not None:
                    geometry = (picture_placeholder.left, picture_placeholder.top,
                                picture_placeholder.width, picture_placeholder.height)
                    if all(value is not None for value in geometry):
                        target_rect = geometry
                    else:
                        layout_picture = None
                        for layout_placeholder in layout.placeholders:
                            if _placeholder_type(layout_placeholder) == "PICTURE":
                                layout_picture = layout_placeholder
                                break
                        if layout_picture is not None and layout_picture.left is not None:
                            target_rect = (layout_picture.left, layout_picture.top,
                                           layout_picture.width, layout_picture.height)
                            warnings.append("slide %d: the picture placeholder inherits "
                                            "its geometry; the layout rectangle was "
                                            "used" % (position + 1))
                        else:
                            target_rect = rect
                try:
                    slide.shapes.add_picture(full, target_rect[0], target_rect[1],
                                             target_rect[2], target_rect[3])
                    if picture_placeholder is not None:
                        # the empty "click to add picture" prompt must not survive
                        try:
                            element = picture_placeholder._element
                            element.getparent().remove(element)
                        except Exception:
                            pass
                        filled.append(picture_placeholder.name + " (picture placed)")
                    else:
                        filled.append("picture")
                except Exception as exc:
                    warnings.append("could not embed figure %s (%s: %s)"
                                    % (figure_id, type(exc).__name__, exc))
                else:
                    images_report.append({"source_ref": figure_id, "mode": "source-crop",
                                          "path": rel, "sha256": digest,
                                          "backend": image.get("backend")})
            else:
                warnings.append("figure %s has no readable image file (%r); nothing was "
                                "embedded" % (figure_id, rel))

        # ---- tables (native + optional redraw-from-data chart) -------------
        table_id = plan_slide.get("table_ref")
        if table_id and table_id in tables:
            table = tables[table_id]
            cells = table.get("cells") or []
            rect, _name = body_rect(slide, layout)
            added = add_native_table(slide, cells, (rect[0], rect[1], rect[2],
                                                    int(rect[3] * 0.55)))
            if added is not None:
                filled.append("table(%dx%d)" % (len(cells[:8]),
                                                max(len(r) for r in cells[:8])))
                rendered_png = os.path.join(derived_dir, "%s-redraw.png" % table_id)
                if redraw_table_chart(cells, rendered_png, table.get("caption") or "",
                                      warnings, table_id):
                    try:
                        slide.shapes.add_picture(rendered_png, rect[0],
                                                 rect[1] + int(rect[3] * 0.55),
                                                 rect[2], int(rect[3] * 0.44))
                        images_report.append({
                            "source_ref": table_id, "mode": "redraw-from-data",
                            "path": C.relpath(rendered_png, work_dir).replace("\\", "/"),
                            "sha256": C.sha256_file(rendered_png),
                            "backend": "matplotlib (values parsed from the source table)"})
                    except Exception as exc:
                        warnings.append("could not place the redrawn chart for %s (%s)"
                                        % (table_id, exc))
            else:
                warnings.append("table %s produced no native table (no cells)" % table_id)

        # ---- notes --------------------------------------------------------
        notes = plan_slide.get("notes") or ""
        notes_chars = len(notes)
        if notes_chars < 40:
            warnings.append("slide %d: speaker notes are shorter than the 40-character "
                            "contract minimum (%d)" % (position + 1, notes_chars))
        try:
            slide.notes_slide.notes_text_frame.text = notes
        except Exception as exc:
            warnings.append("could not write notes on slide %d (%s: %s)"
                            % (position + 1, type(exc).__name__, exc))

        cleared = clear_furniture_text(slide)

        # ---- shrink-to-fit, then the engine-level overflow estimate ----------
        fitted, autofit_declared = [], 0
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            kind = _placeholder_type(shape) if shape.is_placeholder else ""
            if kind in FURNITURE_TYPES:
                continue
            if not (shape.text_frame.text or "").strip():
                continue
            base = resolve_font_size(slide, shape, layout,
                                     DEFAULT_TITLE_PT if kind in TITLE_TYPES
                                     else DEFAULT_BODY_PT)
            adjustment = apply_shrink_to_fit(shape, base)
            if not adjustment:
                continue
            if adjustment.get("autofit"):
                autofit_declared += 1
            if adjustment["to_pt"] < adjustment["from_pt"] - 0.5:
                fitted.append(adjustment)
                warnings.append("slide %d: %s text was fitted from %.0fpt to %.0fpt so "
                                "it stays inside its placeholder"
                                % (position + 1, shape.name, adjustment["from_pt"],
                                   adjustment["to_pt"]))

        risks, details = [], []
        title_text = plan_slide.get("title") or ""
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            kind = _placeholder_type(shape) if shape.is_placeholder else ""
            if kind in FURNITURE_TYPES:
                continue
            text = shape.text_frame.text or ""
            if not text.strip():
                continue
            size = resolve_font_size(slide, shape, layout,
                                     DEFAULT_TITLE_PT if kind in TITLE_TYPES
                                     else DEFAULT_BODY_PT)
            shrink = 0.38 if (images_report and body_name and shape.name == body_name) \
                else 0.0
            risk, detail = estimate_overflow(shape, text.split("\n"), size, shrink)
            detail["shape"] = shape.name
            risks.append(risk)
            details.append(detail)
        if any(r == "high" for r in risks):
            overflow = "high"
        elif any(r == "medium" for r in risks):
            overflow = "medium"
        else:
            overflow = "low"

        report_slides.append({
            "index": position + 1,
            "layout": layout.name,
            "placeholders_filled": filled,
            "overflow_risk": overflow,
            "overflow_detail": details,
            "shrink_to_fit": fitted,
            "autofit_frames": autofit_declared,
            "notes_chars": notes_chars,
            "images": images_report,
            "furniture_cloned": cloned,
            "furniture_present": furniture_present,
            "furniture_text_cleared": cleared,
            "title_chars": len(title_text),
        })

    C.ensure_dir(os.path.dirname(os.path.abspath(out_path)))
    prs.save(out_path)
    restored = _restore_template_parts(out_path, tpl_path, warnings)

    out_parts = zip_parts(out_path, PROTECTED_PREFIXES)
    missing = sorted(set(tpl_parts) - set(out_parts))
    changed = sorted(name for name in tpl_parts
                     if name in out_parts and tpl_parts[name] != out_parts[name])
    added = sorted(set(out_parts) - set(tpl_parts))
    master_unchanged = not any(n.startswith("ppt/slideMasters/")
                               for n in changed + missing)
    theme_unchanged = not any(n.startswith("ppt/theme/") for n in changed + missing)
    layout_changed = [n for n in changed if n.startswith("ppt/slideLayouts/")]
    if changed or missing:
        warnings.append("protected parts differ from the template: %d changed, %d missing "
                        "(after restoring %d part(s) byte-for-byte)"
                        % (len(changed), len(missing), restored))
    if not tpl_parts:
        uncertainty.append("no template package was available, so master/theme equality "
                           "cannot be proven for this run")
        master_unchanged = theme_unchanged = False

    slide_total, slides_with_footer = count_furniture_in_package(out_path)
    footer_ok = bool(slide_total) and slides_with_footer == slide_total
    if report_slides and not footer_ok:
        warnings.append("footer preservation: %d of %d slide(s) carry a footer/slide-number "
                        "placeholder in the saved package" % (slides_with_footer,
                                                              slide_total))

    report = {
        "schema": C.SCHEMA_BUILD_REPORT,
        "case_id": case_id,
        "generated": C.now_iso(),
        # always present, so the report says which engine wrote the deck
        # ("python-pptx" here, "stdlib-fallback" when that import is unavailable)
        "engine": "python-pptx",
        "mode": mode,
        # absolute runtime path: the report must name the exact deck it describes
        # (verify_deck.py compares it with the file it inspects).  This is a runtime
        # artifact of one run, not a shipped resource.
        "pptx": os.path.abspath(out_path),
        "pptx_rel": C.relpath(out_path, work_dir).replace("\\", "/"),
        "sha256": C.sha256_file(out_path),
        "bytes": os.path.getsize(out_path) if os.path.isfile(out_path) else 0,
        "template": {
            "path": os.path.abspath(tpl_path) if tpl_path else "",
            "path_rel": C.relpath(tpl_path, work_dir).replace("\\", "/") if tpl_path
            else "",
            "sha256": C.sha256_file(tpl_path) if tpl_path else "",
            "masters": len(prs.slide_masters),
            "layouts": len(prs.slide_layouts),
            "layout_names": tpl_layout_names,
            "theme_fonts": theme_fonts,
        },
        "invariants": {
            "master_unchanged": bool(master_unchanged),
            "theme_unchanged": bool(theme_unchanged),
            "footer_preserved": bool(footer_ok),
            "layouts_added": len([n for n in added if n.startswith("ppt/slideLayouts/")]),
            "layouts_deleted": len([n for n in missing
                                    if n.startswith("ppt/slideLayouts/")]),
            "layouts_changed": len(layout_changed),
            "footer_placeholders": {"slides": slide_total,
                                    "slides_with_footer": slides_with_footer},
            "protected_parts_restored": restored,
            "protected_parts_changed": changed,
            "protected_parts_missing": missing,
            "parts_added": added,
        },
        "slides": report_slides,
        "fonts_used": sorted(fonts_used | set(read_fonts_in_parts(out_path,
                                                                 warnings=warnings))),
        "slide_size": {"width": prs.slide_width, "height": prs.slide_height},
        "warnings": warnings,
        "uncertainty": uncertainty,
        "stats": {"slides": len(report_slides),
                  "images": sum(len(s["images"]) for s in report_slides),
                  "source_crops": sum(1 for s in report_slides for i in s["images"]
                                      if i["mode"] == "source-crop"),
                  "redraws": sum(1 for s in report_slides for i in s["images"]
                                 if i["mode"] == "redraw-from-data"),
                  "high_overflow": sum(1 for s in report_slides
                                       if s["overflow_risk"] == "high"),
                  "warnings": len(warnings)},
    }
    del Inches
    return report


def build_arg_parser():
    parser = argparse.ArgumentParser(
        prog="build_deck.py",
        description="Build deck.pptx + build_report.json from plan.json.")
    parser.add_argument("--case", required=True, help="path to case.json")
    parser.add_argument("--in", dest="in_path", default=None,
                        help="plan.json (default <work>/plan.json)")
    parser.add_argument("--work", default=None, help="work directory")
    parser.add_argument("--out", default=None, help="deck.pptx path")
    parser.add_argument("--report", default=None, help="build_report.json path")
    return parser


def main(argv=None):
    args = build_arg_parser().parse_args(argv)
    work_dir = C.derive_work_dir(args.case, args.work)
    out_path = args.out or os.path.join(work_dir, "deck.pptx")
    report_path = args.report or os.path.join(work_dir, "build_report.json")
    log = C.RunLog(os.path.join(work_dir, "build.log"))
    try:
        with log.stage("build_deck", case=os.path.basename(args.case)):
            report = build_deck(args.case, work_dir, in_path=args.in_path,
                                out_path=out_path, log=log)
        C.safe_write_json(report_path, report)
    except SystemExit:
        raise
    except Exception as exc:
        log.error("build failed: %s: %s" % (type(exc).__name__, exc))
        C.safe_write_json(report_path, {
            "schema": C.SCHEMA_BUILD_REPORT, "pptx": C.relpath(out_path, work_dir),
            "invariants": {"master_unchanged": False, "theme_unchanged": False,
                           "footer_preserved": False},
            "slides": [], "fonts_used": [], "warnings": [
                "build failed: %s: %s" % (type(exc).__name__, exc)],
            "uncertainty": ["the deck was not built"], "stats": {},
        })
        print("error: build failed; a minimal build_report.json was written to %s"
              % report_path, file=_sys.stderr)
        return 2
    print(json.dumps({"deck": out_path, "report": report_path,
                      "stats": report["stats"],
                      "invariants": {k: v for k, v in report["invariants"].items()
                                     if isinstance(v, bool)}}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
