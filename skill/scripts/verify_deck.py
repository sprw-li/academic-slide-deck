#!/usr/bin/env python3
"""verify_deck.py -- the skill's adversarial self-check gate.

Implements the mandatory V01-V15 list from CONTRACT.md section 3 and writes
``verify_report.json`` (schema ``verify_report/v1``).  **Exit 0 iff verdict == "PASS"**;
``PASS`` requires zero ``FAIL`` and zero ``UNVERIFIED``.  ``SKIP`` is only legal with a
recorded reason, and this script downgrades a reasonless ``SKIP`` to ``UNVERIFIED`` so the
gate cannot be weakened by accident.

Why it is deliberately *not* the builder's self-report
------------------------------------------------------

The builder already writes ``build_report.json`` with an ``invariants`` block claiming
``master_unchanged``, ``theme_unchanged`` etc.  A gate that trusted those booleans would pass
anything.  So this script re-derives the properties from the artefacts themselves:

* V07/V08/V09 parse the **pptx OOXML directly** (zip + lxml) and compare against the template's
  own parts -- the builder's own XML, not its claims.
* V04 re-derives number traceability from ``material.json`` and the **text actually on the
  slides**, never from ``plan.json``'s ``numbers_used``.
* V13 reads ``case.json.output_contract`` and ``result.json`` and compares the key sets.
* V02/V03/V10/V11 read geometry and text from the deck file.
* V12 cross-checks the renderer's reported ``missingFonts`` against the fonts the deck names.
* V10 additionally cross-checks the rendered PDF with ``pdfplumber`` when a render exists.

Honesty rule: a check that cannot inspect its evidence reports ``UNVERIFIED``, which fails the
gate by design (CONTRACT.md:196).  Nothing here is allowed to report ``PASS`` on something it
did not read.

CLI
---

    verify_deck.py --case case.json --work WORK [--out verify_report.json] [--quiet]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import zipfile
from pathlib import Path

SCHEMA = "verify_report/v1"

# --------------------------------------------------------------------------- #
# optional imports -- every one of them degrades rather than crashing
# --------------------------------------------------------------------------- #

try:
    from lxml import etree  # type: ignore
except Exception:  # noqa: BLE001
    etree = None

try:
    from pptx import Presentation  # type: ignore
except Exception:  # noqa: BLE001
    Presentation = None

try:
    from pptx.util import Emu  # type: ignore
except Exception:  # noqa: BLE001
    Emu = None

try:
    from PIL import Image  # type: ignore
except Exception:  # noqa: BLE001
    Image = None

try:
    from fontTools.ttLib import TTFont  # type: ignore
except Exception:  # noqa: BLE001
    TTFont = None

try:
    import pdfplumber  # type: ignore
except Exception:  # noqa: BLE001
    pdfplumber = None

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

EMU_PER_PT = 12700
EMU_PER_INCH = 914400

# --------------------------------------------------------------------------- #
# tiny io helpers (common.py is deck-core's; import when present, else local)
# --------------------------------------------------------------------------- #

try:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from common import dump_json as _dump_json  # type: ignore
    from common import load_json as _common_load_json  # type: ignore
except Exception:  # noqa: BLE001
    _dump_json = None
    _common_load_json = None


def load_json(path, default=None):
    if _common_load_json is not None:
        try:
            value = _common_load_json(path)
            return default if value is None else value
        except Exception:  # noqa: BLE001
            pass
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def dump_json(path, obj):
    if _dump_json is not None:
        try:
            _dump_json(path, obj)
            return
        except Exception:  # noqa: BLE001
            pass
    path = Path(path)
    if str(path.parent) not in ("", "."):
        path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def sha256_file(path):
    try:
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return ""


def file_line(path, needle, default=1):
    """Return ``"<path>:<line>"`` for the first line containing *needle*.

    This is the real evidence pointer form the contract asks for; when the needle is not
    found the pointer still names the file, at line 1, rather than inventing a line.
    """
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            for i, line in enumerate(fh, 1):
                if needle in line:
                    return f"{path}:{i}"
    except OSError:
        pass
    return f"{path}:{default}"


def xml_line(xml_text, needle, container, part):
    """Pointer into an XML part of a container file, e.g. ``deck.pptx!ppt/theme/theme1.xml:41``."""
    if not xml_text:
        return f"{container}!{part}:1"
    for i, line in enumerate(xml_text.splitlines(), 1):
        if needle in line:
            return f"{container}!{part}:{i}"
    return f"{container}!{part}:1"


def find_key_line(path, key):
    return file_line(path, f'"{key}"')


# --------------------------------------------------------------------------- #
# checks bookkeeping
# --------------------------------------------------------------------------- #

class Gate:
    """Accumulates checks and defects; enforces the PASS semantics.

    Three distinct outcomes, deliberately kept apart (see references/verification_gates.md):

    * **N/A**            -- recorded as ``SKIP`` *with a reason*.  Legal under CONTRACT.md:196,
                            which allows SKIP only when a reason is recorded.  Does not fail the
                            gate because there is genuinely nothing to inspect.
    * **WARN**           -- advisory only: appended to ``self.warnings`` and to the check's
                            detail.  Never changes a status, so it cannot fail the gate.
    * **BLOCKING**       -- ``FAIL`` (inspected and wrong) or ``UNVERIFIED`` (could not inspect).
                            Both fail the gate, by design.
    """

    def __init__(self):
        self.checks = []
        self.defects = []
        self.warnings = []

    def warn(self, cid, detail):
        self.warnings.append({"check": cid, "detail": detail})

    def add(self, cid, name, status, detail, evidence, slide=None, defects=None):
        if status == "SKIP" and not (detail or "").strip():
            status = "UNVERIFIED"
            detail = "SKIP without a recorded reason is not allowed; treated as UNVERIFIED."
        self.checks.append({"id": cid, "name": name, "status": status,
                            "detail": detail, "evidence": evidence})
        for d in (defects or []):
            entry = {"severity": d.get("severity", "medium"), "check": cid,
                     "detail": d.get("detail", "")}
            if d.get("slide") is not None:
                entry["slide"] = d["slide"]
            elif slide is not None:
                entry["slide"] = slide
            self.defects.append(entry)

    def counts(self):
        out = {"PASS": 0, "FAIL": 0, "SKIP": 0, "UNVERIFIED": 0}
        for c in self.checks:
            out[c["status"]] = out.get(c["status"], 0) + 1
        return out

    def verdict(self):
        c = self.counts()
        return "PASS" if c["FAIL"] == 0 and c["UNVERIFIED"] == 0 else "FAIL"


# --------------------------------------------------------------------------- #
# pptx / OOXML access
# --------------------------------------------------------------------------- #

def zip_parts(path):
    """``{part_name: bytes}`` for a pptx/docx/xlsx container, or ``{}`` on failure."""
    try:
        with zipfile.ZipFile(path) as zf:
            return {name: zf.read(name) for name in zf.namelist()}
    except (OSError, zipfile.BadZipFile):
        return {}


def part_text(parts, name):
    blob = parts.get(name)
    if not blob:
        return ""
    return blob.decode("utf-8", "replace")


def slide_numbers(parts):
    nums = []
    for name in parts:
        m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
        if m:
            nums.append(int(m.group(1)))
    return sorted(nums)


def layout_names(parts):
    return sorted(n for n in parts if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n))


def master_names(parts):
    return sorted(n for n in parts if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))


def theme_names(parts):
    return sorted(n for n in parts if re.match(r"ppt/theme/theme\d+\.xml$", n))


def canonical_xml(blob):
    """Namespace-normalised XML text, so prefix churn is not mistaken for a change."""
    if etree is None or not blob:
        return blob or b""
    try:
        parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False)
        tree = etree.fromstring(blob, parser)
        return etree.tostring(tree, method="c14n2", strip_text=True)
    except Exception:  # noqa: BLE001
        try:
            return etree.tostring(etree.fromstring(blob), method="c14n")
        except Exception:  # noqa: BLE001
            return blob


def theme_fonts(parts):
    """``{major_latin, minor_latin, major_ea, minor_ea}`` from the first theme part."""
    out = {"major_latin": "", "minor_latin": "", "major_ea": "", "minor_ea": ""}
    for name in theme_names(parts):
        text = part_text(parts, name)
        if not text or etree is None:
            continue
        try:
            root = etree.fromstring(text.encode("utf-8"))
        except Exception:  # noqa: BLE001
            continue
        for tag, latin_key, ea_key in (("majorFont", "major_latin", "major_ea"),
                                       ("minorFont", "minor_latin", "minor_ea")):
            node = root.find(f".//a:{tag}", NS)
            if node is None:
                continue
            latin = node.find("a:latin", NS)
            ea = node.find("a:ea", NS)
            if latin is not None:
                out[latin_key] = latin.get("typeface", "") or ""
            if ea is not None:
                out[ea_key] = ea.get("typeface", "") or ""
        if any(out.values()):
            break
    return out


def fonts_named_in_deck(parts):
    """Every ``typeface="..."`` used anywhere in the deck's own XML parts."""
    names = set()
    for name, blob in parts.items():
        if not name.endswith(".xml"):
            continue
        text = blob.decode("utf-8", "replace")
        for m in re.finditer(r'typeface="([^"]+)"', text):
            value = m.group(1).strip()
            if value:
                names.add(value)
    return sorted(names)


_THEME_REF_RE = re.compile(r"^\+(?:mj|mn)-(?:lt|ea|cs)$", re.I)
_TEXT_PART_RE = re.compile(
    r"ppt/(?:slides/slide\d+|notesSlides/notesSlide\d+|slideLayouts/slideLayout\d+"
    r"|slideMasters/slideMaster\d+|presentation)\.xml$")


def fonts_used_in_deck(parts):
    """Fonts the deck can actually put glyphs in.

    Two deliberate exclusions, both learned from a false positive on the first run
    (2026-11-23): the Office theme's ``<a:fontScheme>`` carries a large *supplemental*
    script-font table (Angsana New, Cordia New, Batang, ...) that is a fallback map, not a
    used font; and ``+mj-lt``/``+mn-ea`` style tokens are references *to* the theme fonts
    rather than font names.  Counting either makes V12 fail on a perfectly good deck.
    """
    used = set()
    for name, blob in parts.items():
        if not _TEXT_PART_RE.search(name):
            continue
        text = blob.decode("utf-8", "replace")
        for m in re.finditer(r'typeface="([^"]+)"', text):
            value = m.group(1).strip()
            if value and not value.startswith("+") and not _THEME_REF_RE.match(value):
                used.add(value)
    for key, value in theme_fonts(parts).items():
        if value and not value.startswith("+"):
            used.add(value)
    return sorted(used)


# --------------------------------------------------------------------------- #
# font resolution + text metrics (for V10/V12)
# --------------------------------------------------------------------------- #

_FONT_DIR_CACHE = None
_FONT_INDEX = {}


def font_dirs():
    global _FONT_DIR_CACHE
    if _FONT_DIR_CACHE is not None:
        return _FONT_DIR_CACHE
    dirs = []
    if os.name == "nt":
        win = os.environ.get("WINDIR")
        if win:
            dirs.append(Path(win) / "Fonts")
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    dirs += [Path.home() / ".fonts", Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
             Path("/Library/Fonts"), Path("/System/Library/Fonts")]
    try:
        import matplotlib  # type: ignore
        dirs.append(Path(matplotlib.get_data_path()) / "fonts" / "ttf")
    except Exception:  # noqa: BLE001
        pass
    _FONT_DIR_CACHE = [d for d in dirs if d.is_dir()]
    return _FONT_DIR_CACHE


def _norm_family(name):
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def resolve_font(family):
    """Best-effort font file for *family*; ``None`` when nothing plausible is installed."""
    key = _norm_family(family)
    if not key:
        return None
    if key in _FONT_INDEX:
        return _FONT_INDEX[key]
    best = None
    for directory in font_dirs():
        try:
            for ext in ("*.ttf", "*.otf", "*.ttc", "*.TTF", "*.OTF"):
                for path in directory.rglob(ext):
                    stem = _norm_family(path.stem)
                    if stem.startswith(key) or key.startswith(stem) or key in stem:
                        best = str(path)
                        break
                if best:
                    break
        except OSError:
            continue
        if best:
            break
    _FONT_INDEX[key] = best
    return best


def _fallback_font():
    for family in ("DejaVu Sans", "Liberation Sans", "Arial"):
        path = resolve_font(family)
        if path:
            return path
    return None


_WIDTH_CACHE = {}


def text_width_pt(text, font_path, size_pt):
    """Advance width of *text* in points, from real font metrics when possible."""
    if not text:
        return 0.0
    if font_path and TTFont is not None:
        cache_key = (font_path, id(None))
        try:
            if cache_key not in _WIDTH_CACHE:
                f = TTFont(font_path, lazy=True)
                _WIDTH_CACHE[cache_key] = f
            font = _WIDTH_CACHE[cache_key]
            upem = font["head"].unitsPerEm or 1000
            cmap = font.getBestCmap()
            hmtx = font["hmtx"]
            total = 0
            for ch in text:
                gname = cmap.get(ord(ch))
                if gname is None:
                    total += int(upem * (1.0 if ord(ch) > 0x2E80 else 0.5))
                    continue
                total += hmtx[gname][0]
            return total * size_pt / float(upem)
        except Exception:  # noqa: BLE001
            pass
    # Heuristic fallback: CJK is full-width, Latin about half an em.
    width = 0.0
    for ch in text:
        width += (1.0 if ord(ch) > 0x2E80 else 0.5) * size_pt
    return width


def estimate_paragraph_lines(text, font_path, size_pt, avail_pt):
    """Greedy wrap: how many lines does *text* need inside *avail_pt* points?"""
    if not text:
        return 0
    if avail_pt <= 1:
        return 1
    lines = 0
    for hard_line in text.split("\n"):
        if not hard_line.strip():
            lines += 1
            continue
        # Break CJK everywhere, Latin on spaces: build greedy tokens.
        tokens = re.findall(r"[A-Za-z0-9\u00C0-\u024F'’\-\.]+|\s+|[^\sA-Za-z0-9]", hard_line)
        cur = 0.0
        used = 1
        for tok in tokens:
            w = text_width_pt(tok, font_path, size_pt)
            if cur + w > avail_pt and cur > 0:
                used += 1
                cur = w if not tok.isspace() else 0.0
            else:
                cur += w
        lines += used
    return lines


def shape_text_extent(shape):
    """Estimate whether a shape's text fits, using real metrics where available.

    Returns ``(overflow, detail, needed_emu, avail_emu)``.  ``overflow`` is ``None`` when the
    shape could not be measured (no text frame, no text), which the caller must treat as
    "not inspected" rather than "fits".
    """
    if Emu is None:
        return None, "python-pptx unavailable", 0, 0
    try:
        tf = shape.text_frame
    except Exception:  # noqa: BLE001
        return None, "no text frame", 0, 0
    text = tf.text or ""
    if not text.strip():
        return None, "empty text frame", 0, 0

    try:
        width_emu = int(shape.width or 0)
        height_emu = int(shape.height or 0)
    except Exception:  # noqa: BLE001
        return None, "no geometry", 0, 0
    if width_emu <= 0 or height_emu <= 0:
        return None, "zero-size shape", 0, 0

    def margin(name, default):
        try:
            value = getattr(tf, name, None)
            return int(value) if value is not None else default
        except Exception:  # noqa: BLE001
            return default

    ml, mr = margin("margin_left", 91440), margin("margin_right", 91440)
    mt, mb = margin("margin_top", 45720), margin("margin_bottom", 45720)
    avail_w = max(1, width_emu - ml - mr)
    avail_h = max(1, height_emu - mt - mb)
    avail_pt = avail_w / EMU_PER_PT

    total_lines = 0
    sizes = []
    for para in tf.paragraphs:
        size_pt = None
        indent = 0
        try:
            pPr = para._p.find(f"{{{NS['a']}}}pPr")
            if pPr is not None and pPr.get("marL"):
                indent = int(pPr.get("marL") or 0)
        except Exception:  # noqa: BLE001
            indent = 0
        for run in para.runs:
            try:
                if run.font.size is not None:
                    size_pt = run.font.size.pt
                    break
            except Exception:  # noqa: BLE001
                continue
        if size_pt is None:
            # fall back to the paragraph endParaRPr or the placeholder default
            m = re.search(r'sz="(\d+)"', (para._p.xml if hasattr(para, "_p") else ""))
            size_pt = int(m.group(1)) / 100.0 if m else 18.0
        size_pt = float(size_pt)
        sizes.append(size_pt)
        font_path = _fallback_font()
        total_lines += estimate_paragraph_lines(para.text, font_path,
                                                size_pt, max(1.0, avail_pt - indent / EMU_PER_PT))

    # line spacing ~1.2em; the last line still needs its full height
    needed_pt = 0.0
    per_para_lines = []
    idx = 0
    for para in tf.paragraphs:
        size_pt = sizes[idx] if idx < len(sizes) else 18.0
        idx += 1
        n = estimate_paragraph_lines(para.text, _fallback_font(), size_pt, avail_pt)
        per_para_lines.append(n)
        needed_pt += max(1, n) * size_pt * 1.2
    needed_emu = int(needed_pt * EMU_PER_PT) + mt + mb

    # PowerPoint will shrink text when an autofit block is present, so overflow is a softer
    # signal there; the caller records that in the detail string.
    return (needed_emu > avail_h), \
        f"needs~{needed_emu / EMU_PER_PT:.1f}pt vs box {avail_h / EMU_PER_PT:.1f}pt, " \
        f"lines={sum(per_para_lines)}", needed_emu, avail_h


def shape_autofit_kind(shape):
    """Which autofit the shape carries, and what it actually means.

    ``normAutofit`` shrinks the *text* to fit the box ("shrink text on overflow"), so the
    content is not clipped.  ``spAutoFit`` resizes the *shape* to fit the text, which can push
    the shape off the slide -- a different failure that must be checked against the slide box,
    not against the shape box.  Conflating the two is how a gate stops noticing real defects.
    """
    try:
        xml = shape._element.xml
    except Exception:  # noqa: BLE001
        return None
    if "normAutofit" in xml:
        return "norm"
    if "spAutoFit" in xml:
        return "sp"
    return None


# --------------------------------------------------------------------------- #
# number extraction (V04)
# --------------------------------------------------------------------------- #

#: figure/table/section labels are numbering machinery, not scientific values
_LABEL_RE = re.compile(
    r"\b(?:fig(?:ure)?|tab(?:le)?|sec(?:tion)?|eq(?:uation)?|appendix|step|page|slide|ref|no)"
    r"\s*\.?\s*\d+", re.I)
_NUM_RE = re.compile(r"\d+(?:[.,]\d+)?")
_UNIT_RE = re.compile(r"^\s*(%|\u2030|\u00b1|\u00d7)?")


def numbers_on_text(text):
    """Numeric tokens in *text*, with caption labels and dates removed."""
    if not text:
        return []
    stripped = _LABEL_RE.sub(" ", text)
    return _NUM_RE.findall(stripped)


def allowed_numbers(material):
    """Every representation of a number that material.json actually attests."""
    allowed = set()
    for item in (material or {}).get("numbers", []) or []:
        raw = str(item.get("raw", "")).strip()
        if raw:
            allowed.add(raw)
            allowed.add(raw.replace(" ", ""))
        value = item.get("value")
        if isinstance(value, (int, float)):
            allowed.add(str(value))
            allowed.add(f"{value:g}")
            if float(value).is_integer():
                allowed.add(str(int(value)))
    return allowed


# --------------------------------------------------------------------------- #
# the checks
# --------------------------------------------------------------------------- #

def resolve_existing(candidate, bases):
    """Resolve *candidate* against *bases* (case dir, work dir, cwd) and return it if real.

    Case fixtures declare ``template_path`` and ``sources`` **relative to the case file**
    (e.g. ``materials/template-organizer.pptx``), so a plain ``Path(x).is_file()`` from the
    process cwd silently fails to find the template and turns V07/V08/V09 into UNVERIFIED.
    """
    if not candidate:
        return None
    raw = Path(str(candidate))
    tries = [raw]
    for base in bases:
        if base:
            tries.append(Path(base) / raw)
    for path in tries:
        try:
            if path.is_file():
                return Path(path)
        except OSError:
            continue
    return None


def resolve_template_path(case, plan, build, bases=()):
    for candidate in (
        (((plan or {}).get("constraints") or {}).get("template_path")),
        (((plan or {}).get("constraints") or {}).get("template")),
        ((((build or {}).get("template") or {}).get("path"))),
        ((((case or {}).get("constraints") or {}).get("template_path"))),
        ((case or {}).get("template")),
    ):
        found = resolve_existing(candidate, bases)
        if found:
            return found
    return None


def collect_case_constraints(case):
    """Slide-count constraints, tolerant about where the case puts them."""
    found = {}
    if not isinstance(case, dict):
        return found
    pools = [case, case.get("constraints") if isinstance(case.get("constraints"), dict) else {},
             case.get("output_contract") if isinstance(case.get("output_contract"), dict) else {},
             case.get("task") if isinstance(case.get("task"), dict) else {}]
    for pool in pools:
        for key in ("slide_count", "slide_count_expected", "n_slides", "num_slides"):
            if isinstance(pool.get(key), int):
                found.setdefault("slide_count", pool[key])
        rng = pool.get("slide_count_range")
        if isinstance(rng, (list, tuple)) and len(rng) == 2:
            found.setdefault("slide_count_range", [int(rng[0]), int(rng[1])])
        for key in ("duration_min", "duration_minutes"):
            if isinstance(pool.get(key), int):
                found.setdefault("duration_min", pool[key])
    return found


def collect_constraint_metadata(case, plan):
    """Audience / duration / template / language, for the leakage check (V17).

    These are *instructions to the skill*, not content for the audience.  A slide that repeats
    them verbatim is telling the room what the brief said.
    """
    meta = {}
    pools = []
    if isinstance(case, dict):
        pools += [case.get("constraints"), case]
    if isinstance(plan, dict):
        pools += [plan.get("constraints"), plan]
    for pool in pools:
        if not isinstance(pool, dict):
            continue
        for key in ("audience", "duration_min", "duration_minutes", "template_path",
                    "template", "language", "lang"):
            if pool.get(key) and key not in meta:
                meta[key] = pool[key]
    return meta


# --------------------------------------------------------------------------- #
# content hygiene helpers (V16 / V17)
# --------------------------------------------------------------------------- #

_PUNCT_RE = re.compile(r"[^\w\s%.\-]+", re.UNICODE)

#: leading "Label:" / "Label：" prefix on a bullet, e.g. "Key result: …", "方法：…"
_LABEL_PREFIX_RE = re.compile(r"^\s*[^:：\n]{1,24}[:：]\s*")
_TRAIL_PUNCT_RE = re.compile(r"[\s.,;:!?。，、；：！？—\-]+$", re.UNICODE)

#: function words ignored by the content-word overlap test, so "the same sentence with a
#: different label" cannot hide behind shared filler
_STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "been", "but", "by", "can", "did", "do",
    "does", "for", "from", "had", "has", "have", "how", "in", "into", "is", "it", "its",
    "may", "might", "must", "no", "not", "of", "on", "or", "our", "out", "over", "should",
    "so", "some", "such", "than", "that", "the", "their", "them", "then", "there", "these",
    "they", "this", "those", "to", "up", "was", "we", "were", "what", "when", "where",
    "which", "while", "who", "will", "with", "would", "you", "your",
}


def _norm_text(text):
    lowered = str(text or "").lower()
    return re.sub(r"\s+", " ", _PUNCT_RE.sub(" ", lowered)).strip()


def bullet_fingerprint(text):
    """Normalise a bullet so that boilerplate repeats regardless of its role prefix.

    Strips a leading ``"Label:"`` prefix, lowercases, collapses whitespace and drops trailing
    punctuation.  ``"Problem: what this page has to establish."`` and
    ``"Conclusion: what this page has to establish."`` therefore collapse to the same
    fingerprint -- which is the whole point: the shared tail is the defect, not the label.
    """
    stripped = _LABEL_PREFIX_RE.sub("", str(text or ""))
    normalised = _norm_text(stripped)
    return _TRAIL_PUNCT_RE.sub("", normalised).strip()


def content_words(normalised_text):
    """Content words of a normalised string, with stopwords removed."""
    return {w for w in normalised_text.split() if len(w) > 1 and w not in _STOPWORDS}


#: a repeated bullet must carry this much text before repetition is called a defect, so short
#: repeated navigation phrases do not become false positives
MIN_REPEAT_CHARS = 20
#: content-word overlap that counts as the same sentence written twice
CONTENT_OVERLAP = 0.70
#: both sides need this many content words before the overlap test is meaningful
MIN_CONTENT_WORDS = 3


def slide_text_units(plan_slides, slides):
    """``[(slide_no, kind, normalised_text, raw_text)]`` for titles, bullets and notes."""
    units = []
    if plan_slides:
        for i, ps in enumerate(plan_slides, 1):
            if not isinstance(ps, dict):
                continue
            for kind, raw in (("title", ps.get("title")),
                              ("notes", ps.get("notes"))):
                n = _norm_text(raw)
                if n:
                    units.append((i, kind, n, str(raw)))
            for b in ps.get("bullets") or []:
                n = _norm_text(b)
                if n:
                    units.append((i, "bullet", n, str(b)))
    else:
        for i, s in enumerate(slides, 1):
            try:
                if s.shapes.title is not None and (s.shapes.title.text or "").strip():
                    raw = s.shapes.title.text
                    units.append((i, "title", _norm_text(raw), raw))
            except Exception:  # noqa: BLE001
                pass
            for sh in s.shapes:
                try:
                    if sh.has_text_frame and sh.text_frame.text.strip():
                        raw = sh.text_frame.text
                        units.append((i, "text", _norm_text(raw), raw))
                except Exception:  # noqa: BLE001
                    continue
    return units


def duplicate_content_issues(units, min_chars=20, threshold=None):
    """Pairs of units on *different* slides that are the same or nearly the same.

    Four detections, in order of how directly they prove duplication:

    1. **repeated bullet** -- the same bullet fingerprint (leading ``"Label:"`` prefix stripped)
       on >= 2 slides.  Boilerplate repeats by construction: ``"Problem: what this page has to
       establish."`` … ``"Conclusion: what this page has to establish."`` all collapse to one
       fingerprint even though their labels differ.
    2. identical strings.
    3. one string contained in the other.
    4. **content-word overlap** >= :data:`CONTENT_OVERLAP` over the shorter side, stopwords
       ignored.  A character-level percentage lets "the same sentence with a different label"
       slip under the line (measured on a real deck: ~77% shared characters, just under 80%);
       content words are what actually repeat.

    Notes are compared for **exact** equality only, and never against a title/bullet: boilerplate
    speaker notes differing by slide number are legitimate and a fuzzy test on them would fail
    every deck built from a house template.
    """
    issues = []
    seen = set()

    def remember(si, sj, reason):
        key = (min(si, sj), max(si, sj), reason.split(":")[0])
        if key in seen:
            return
        seen.add(key)
        issues.append((si, sj, reason))

    # ---- 1. repeated bullet fingerprints ----
    # Counted regardless of *which* slide carries them.  The real defect that motivated this
    # rule had five stubs on a SINGLE slide ("Problem: … Method: … Conclusion: what this page
    # has to establish."), so a "same bullet on >= 2 slides" test would have missed it; two
    # occurrences anywhere in the deck is already boilerplate, and repeating a 20+ character
    # bullet twice on one page is a defect on its own.
    by_fingerprint = {}
    for si, kind, _norm_value, raw in units:
        if kind != "bullet":
            continue
        fp = bullet_fingerprint(raw)
        if len(fp) < MIN_REPEAT_CHARS:
            continue
        by_fingerprint.setdefault(fp, []).append(si)
    for fp, slides_list in by_fingerprint.items():
        if len(slides_list) < 2:
            continue
        where = sorted(set(slides_list))
        descriptor = (f"{len(slides_list)} times on slide {where[0]}" if len(where) == 1
                      else f"on slides {where}")
        remember(where[0], where[-1],
                 f"the same bullet text appears {descriptor} (after stripping any 'Label:' "
                 f"prefix), so the page carries boilerplate rather than content: {fp[:70]!r}")

    # ---- 2-4. pairwise comparison ----
    comparable = []
    for si, kind, norm, raw in units:
        cmp_text = bullet_fingerprint(raw) if kind == "bullet" else norm
        comparable.append((si, kind, norm, cmp_text))

    n = len(comparable)
    for a in range(n):
        si, ki, ti, ci = comparable[a]
        if len(ti) < min_chars:
            continue
        for b in range(a + 1, n):
            sj, kj, tj, cj = comparable[b]
            if si == sj or len(tj) < min_chars:
                continue
            if "notes" in (ki, kj):
                if ki == "notes" and kj == "notes" and ti == tj:
                    remember(si, sj, f"notes of slide {si} are identical to notes of slide "
                                     f"{sj}: {ti[:60]!r}")
                continue
            if ki == kj == "bullet" and ci == cj and len(ci) >= MIN_REPEAT_CHARS:
                continue                      # already reported by rule 1
            reason = None
            if ti == tj:
                reason = f"{ki} of slide {si} is identical to {kj} of slide {sj}"
            elif ti in tj or tj in ti:
                reason = (f"{ki} of slide {si} is contained in {kj} of slide {sj} "
                          f"(same text reused on another page)")
            else:
                A, B = content_words(ci), content_words(cj)
                if len(A) >= MIN_CONTENT_WORDS and len(B) >= MIN_CONTENT_WORDS:
                    ratio = len(A & B) / float(min(len(A), len(B)))
                    if ratio >= CONTENT_OVERLAP:
                        reason = (f"{ki} of slide {si} shares {ratio:.0%} of its content words "
                                  f"(stopwords ignored) with {kj} of slide {sj}")
            if reason:
                remember(si, sj, f"{reason}: {ti[:70]!r}")
    return issues


#: a *labelled* brief field, e.g. "Audience: ...", "Duration: 10.0", "slide_count = 8"
_LEAK_LABEL_RE = re.compile(
    r"\b(audience|duration|duration_min|slide[_ ]?count|number[_ ]?of[_ ]?slides|template"
    r"|template[_ ]?path|starting[_ ]?deck|constraints?)\b\s*[:：=]", re.I)


def constraint_leak_issues(units, meta):
    """Slides that repeat the brief's own metadata (V17).

    Matches a *labelled* field (`Audience: ...`), the declared audience string, the template
    file name, and a "<duration> min" phrasing.  A bare mention of the word "duration" is not
    a leak -- the label separator is required -- so ordinary scientific prose is not flagged.
    """
    audience = str(meta.get("audience") or "").strip()
    template = str(meta.get("template_path") or meta.get("template") or "").strip()
    template_base = Path(template).name if template else ""
    duration = meta.get("duration_min", meta.get("duration_minutes"))
    issues = []
    for slide_no, kind, norm, raw in units:
        text = str(raw)
        m = _LEAK_LABEL_RE.search(text)
        if m:
            issues.append((slide_no, f"{kind} repeats the brief's own {m.group(1)!r} field: "
                                     f"{text.strip()[:80]!r}"))
            continue
        if len(audience) >= 6 and audience.lower() in text.lower():
            issues.append((slide_no, f"{kind} quotes the declared audience text verbatim: "
                                     f"{text.strip()[:80]!r}"))
            continue
        if template_base and len(template_base) > 5 and template_base.lower() in text.lower():
            issues.append((slide_no, f"{kind} names the template file {template_base!r}: "
                                     f"{text.strip()[:80]!r}"))
            continue
        if duration is not None and re.search(rf"\b{re.escape(str(duration))}\b\s*(min|minutes)\b",
                                             text, re.I):
            issues.append((slide_no, f"{kind} states the brief's duration ({duration}): "
                                     f"{text.strip()[:80]!r}"))
    return issues


def run_checks(case_path, work, out=None, case_id=None):
    work = Path(work)
    case = load_json(case_path) if case_path else {}
    case = case if isinstance(case, dict) else {}

    plan_path = work / "plan.json"
    material_path = work / "material.json"
    build_path = work / "build_report.json"
    render_path = work / "render_report.json"
    evidence_path = work / "evidence_map.json"
    result_path = (Path(out) / "result.json") if out else work / "result.json"
    if not result_path.is_file() and (work / "result.json").is_file():
        result_path = work / "result.json"
    deck_path = work / "deck.pptx"

    plan = load_json(plan_path, {}) or {}
    material = load_json(material_path, {}) or {}
    build = load_json(build_path, {}) or {}
    render = load_json(render_path, {}) or {}
    emap = load_json(evidence_path, {}) or {}
    result = load_json(result_path, {}) or {}

    # Which deck to inspect?  The documented location wins (CONTRACT.md:76 says
    # build_deck.py writes ``W/deck.pptx``) because build_report.json may hold a stale or
    # foreign absolute path.  A disagreement is itself worth recording: a gate that silently
    # followed build_report's path would verify a *different* file than the work dir contains.
    deck_source = "work-dir"
    work_deck = work / "deck.pptx"
    reported_deck = ((build or {}).get("pptx") or "") if isinstance(build, dict) else ""
    deck_mismatch = None
    if work_deck.is_file():
        deck_path = work_deck
        try:
            if reported_deck and Path(reported_deck).resolve() != work_deck.resolve():
                deck_mismatch = reported_deck
        except OSError:
            deck_mismatch = reported_deck
    elif reported_deck and Path(reported_deck).is_file():
        deck_path = Path(reported_deck)
        deck_source = "build_report"
    elif not deck_path.is_file():
        deck_path = work_deck

    # Same reasoning for the rendered PDF used by V10's cross-check.
    pdf_from_render = ((render or {}).get("pdf") or {}).get("path") or ""
    work_pdf = work / "render" / "deck.pdf"
    pdf_for_check = str(work_pdf) if work_pdf.is_file() else pdf_from_render
    if pdf_for_check:
        render = dict(render)
        render["pdf"] = dict((render.get("pdf") or {}))
        render["pdf"]["path"] = pdf_for_check

    g = Gate()

    # ---------------- read the artefacts independently ---------------- #
    prs = None
    slides = []
    deck_parts = {}
    deck_slides_xml = []
    if Presentation is not None and deck_path.is_file():
        try:
            prs = Presentation(str(deck_path))
            slides = list(prs.slides)
        except Exception:  # noqa: BLE001
            prs = None
    if deck_path.is_file():
        deck_parts = zip_parts(deck_path)
        deck_slides_xml = [(n, part_text(deck_parts, f"ppt/slides/slide{n}.xml"))
                           for n in slide_numbers(deck_parts)]

    deck_slide_count = len(slides) if slides else len(slide_numbers(deck_parts))
    plan_slides = (plan or {}).get("slides") or []

    # ---------------- V01 slide count ---------------- #
    constraint = collect_case_constraints(case)
    if not constraint:
        pc = (plan or {}).get("constraints") or {}
        if isinstance(pc.get("slide_count"), int):
            constraint["slide_count"] = pc["slide_count"]
        if isinstance(pc.get("slide_count_range"), (list, tuple)):
            constraint["slide_count_range"] = list(pc["slide_count_range"])
    if deck_slide_count == 0:
        g.add("V01", "slide count within the case constraint", "UNVERIFIED",
              "No deck was readable, so the slide count could not be inspected.",
              f"{deck_path}:1")
    elif not constraint:
        g.add("V01", "slide count within the case constraint", "UNVERIFIED",
              f"Deck has {deck_slide_count} slides but no slide-count constraint could be "
              f"found in case.json (keys checked: slide_count, slide_count_expected, n_slides, "
              f"num_slides, slide_count_range) or plan.constraints. Cannot verify compliance.",
              f"{case_path}:1" if case_path else f"{plan_path}:1")
    else:
        ok, why = True, ""
        if "slide_count" in constraint and deck_slide_count != constraint["slide_count"]:
            ok, why = False, (f"deck has {deck_slide_count} slides, case requires exactly "
                              f"{constraint['slide_count']}")
        if "slide_count_range" in constraint:
            lo, hi = constraint["slide_count_range"]
            if not (lo <= deck_slide_count <= hi):
                ok, why = False, (f"deck has {deck_slide_count} slides, case requires "
                                  f"{lo}-{hi}")
        if ok:
            g.add("V01", "slide count within the case constraint", "PASS",
                  f"deck has {deck_slide_count} slides; constraint {constraint} satisfied "
                  f"(counted from the deck file itself)",
                  f"{deck_path}:1" if deck_path.is_file() else f"{plan_path}:1")
        else:
            g.add("V01", "slide count within the case constraint", "FAIL", why,
                  f"{deck_path}:1", defects=[{"severity": "high", "detail": why}])

    # ---------------- V02 titles + claims ---------------- #
    if not plan_slides and not slides:
        g.add("V02", "every slide has non-empty title and claim", "UNVERIFIED",
              "Neither plan.json slides nor a readable deck were available.",
              f"{plan_path}:1")
    else:
        bad = []
        n = max(len(plan_slides), len(slides))
        for i in range(n):
            ps = plan_slides[i] if i < len(plan_slides) else {}
            s = slides[i] if i < len(slides) else None
            title = (ps.get("title") or "").strip()
            claim = (ps.get("claim") or "").strip()
            if s is not None:
                deck_title = ""
                try:
                    if s.shapes.title is not None:
                        deck_title = (s.shapes.title.text or "").strip()
                except Exception:  # noqa: BLE001
                    deck_title = ""
                if not deck_title and not title:
                    bad.append((i + 1, "no title in deck or plan"))
                elif title and not deck_title:
                    bad.append((i + 1, f"plan title {title!r} not present as deck title"))
            elif not title:
                bad.append((i + 1, "plan slide has no title"))
            if not claim:
                bad.append((i + 1, "no claim in plan"))
        if bad:
            g.add("V02", "every slide has non-empty title and claim", "FAIL",
                  "; ".join(f"slide {s}: {d}" for s, d in bad[:12]),
                  find_key_line(plan_path, "claim") if plan_path.is_file() else f"{deck_path}:1",
                  defects=[{"severity": "high", "detail": d, "slide": s} for s, d in bad])
        else:
            g.add("V02", "every slide has non-empty title and claim", "PASS",
                  f"all {n} slides carry a non-empty deck title and a non-empty plan claim",
                  find_key_line(plan_path, "claim") if plan_path.is_file() else f"{deck_path}:1")

    # ---------------- V03 speaker notes >= 40 chars ---------------- #
    MIN_NOTES = 40
    if not slides and not plan_slides:
        g.add("V03", "every slide has notes >= 40 chars", "UNVERIFIED",
              "No deck and no plan slides to read notes from.", f"{work}:1")
    else:
        bad = []
        for i, s in enumerate(slides, 1):
            notes = ""
            try:
                if s.has_notes_slide:
                    notes = s.notes_slide.notes_text_frame.text or ""
            except Exception:  # noqa: BLE001
                notes = ""
            if len(notes.strip()) < MIN_NOTES:
                ps = plan_slides[i - 1] if i - 1 < len(plan_slides) else {}
                plan_notes = (ps.get("notes") or "")
                extra = (f"; plan carries {len(plan_notes.strip())} chars but the deck's notes "
                         f"slide does not") if len(plan_notes.strip()) >= MIN_NOTES else ""
                bad.append((i, f"deck notes {len(notes.strip())} chars{extra}"))
        if not slides and plan_slides:
            for i, ps in enumerate(plan_slides, 1):
                if len((ps.get("notes") or "").strip()) < MIN_NOTES:
                    bad.append((i, f"plan notes {len((ps.get('notes') or '').strip())} chars"))
        if bad:
            g.add("V03", "every slide has notes >= 40 chars", "FAIL",
                  "; ".join(f"slide {s}: {d}" for s, d in bad[:12]), f"{deck_path}:1",
                  defects=[{"severity": "high", "detail": d, "slide": s} for s, d in bad])
        else:
            g.add("V03", "every slide has notes >= 40 chars", "PASS",
                  f"all {len(slides) or len(plan_slides)} slides carry >= {MIN_NOTES} chars of "
                  f"notes, measured from the deck's notes slides",
                  f"{deck_path}:1")

    # ---------------- V04 number traceability (from material + slide text) ---------------- #
    if not (material or {}).get("numbers") and not material:
        g.add("V04", "every number on a slide is traceable to material.numbers", "UNVERIFIED",
              "material.json is missing, so traceability cannot be re-derived.",
              f"{material_path}:1")
    else:
        allowed = allowed_numbers(material)
        offenders = []
        # authoritative text: what is actually written on the slides
        slide_texts = []
        for i, s in enumerate(slides, 1):
            chunks = []
            for sh in s.shapes:
                try:
                    if sh.has_text_frame and sh.text_frame.text.strip():
                        chunks.append(sh.text_frame.text)
                except Exception:  # noqa: BLE001
                    continue
            slide_texts.append((i, "\n".join(chunks)))
        if not slide_texts:
            for i, ps in enumerate(plan_slides, 1):
                pieces = [ps.get("title") or "", ps.get("claim") or ""] + list(ps.get("bullets") or [])
                slide_texts.append((i, "\n".join(str(p) for p in pieces)))
        for i, text in slide_texts:
            for token in numbers_on_text(text):
                if token in allowed:
                    continue
                # tolerant: punctuation-normalised forms
                norm = token.replace(",", ".")
                if norm in allowed or token.rstrip("0").rstrip(".") in allowed:
                    continue
                offenders.append((i, token))
        if not allowed:
            g.add("V04", "every number on a slide is traceable to material.numbers", "UNVERIFIED",
                  "material.numbers is empty, so no slide number can be attested; the check "
                  "cannot pass and is not allowed to pass by default.",
                  f"{material_path}:1")
        elif offenders:
            g.add("V04", "every number on a slide is traceable to material.numbers", "FAIL",
                  f"{len(offenders)} numeric token(s) on slides have no verbatim match in "
                  f"material.numbers: "
                  + ", ".join(f"slide {s}:{t}" for s, t in offenders[:15]),
                  find_key_line(material_path, "numbers"), defects=[
                      {"severity": "high", "slide": s,
                       "detail": f"number {t!r} on the slide is not attested in material.numbers"}
                      for s, t in offenders])
        else:
            g.add("V04", "every number on a slide is traceable to material.numbers", "PASS",
                  f"every numeric token on the slides matches one of the "
                  f"{len(allowed)} attested representations in material.numbers "
                  f"(caption labels stripped; source = deck text, not plan.json)",
                  find_key_line(material_path, "numbers"))

    # ---------------- V05 figure/table refs resolve ---------------- #
    fig_ids = {f.get("id") for f in (material.get("figures") or []) if isinstance(f, dict)}
    tab_ids = {t.get("id") for t in (material.get("tables") or []) if isinstance(t, dict)}
    if not material:
        g.add("V05", "every figure/table reference resolves in material.json", "UNVERIFIED",
              "material.json is missing.", f"{material_path}:1")
    else:
        missing = []
        for i, ps in enumerate(plan_slides, 1):
            fr, tr = ps.get("figure_ref"), ps.get("table_ref")
            if fr and fr not in fig_ids:
                missing.append((i, f"figure_ref {fr!r} not in material.figures {sorted(fig_ids)}"))
            if tr and tr not in tab_ids:
                missing.append((i, f"table_ref {tr!r} not in material.tables {sorted(tab_ids)}"))
            for ev in ps.get("evidence") or []:
                if not isinstance(ev, dict):
                    continue
                ref = ev.get("ref")
                kind = ev.get("type")
                if kind == "figure" and ref and ref not in fig_ids:
                    missing.append((i, f"evidence figure {ref!r} not in material.figures"))
                if kind == "table" and ref and ref not in tab_ids:
                    missing.append((i, f"evidence table {ref!r} not in material.tables"))
        if missing:
            g.add("V05", "every figure/table reference resolves in material.json", "FAIL",
                  "; ".join(f"slide {s}: {d}" for s, d in missing[:12]),
                  find_key_line(material_path, "figures"), defects=[
                      {"severity": "high", "slide": s, "detail": d} for s, d in missing])
        else:
            g.add("V05", "every figure/table reference resolves in material.json", "PASS",
                  f"all figure/table references resolve against {len(fig_ids)} figures and "
                  f"{len(tab_ids)} tables in material.json", find_key_line(material_path, "figures"))

    # ---------------- V06 embedded images have real provenance ---------------- #
    ALLOWED_MODES = {"source-crop", "redraw-from-data"}
    declared = {}
    for entry in (build.get("slides") or []):
        if not isinstance(entry, dict):
            continue
        declared[int(entry.get("index", 0) or 0)] = entry.get("images") or []
    embedded_counts = {}
    for i, s in enumerate(slides, 1):
        count = 0
        for sh in s.shapes:
            try:
                if sh.shape_type is not None and "PICTURE" in str(sh.shape_type):
                    count += 1
                elif hasattr(sh, "image"):
                    count += 1
            except Exception:  # noqa: BLE001
                continue
        embedded_counts[i] = count
    if not slides and not build:
        g.add("V06", "every embedded image is source-crop or redraw-from-data", "UNVERIFIED",
              "Neither the deck nor build_report.json was readable.", f"{build_path}:1")
    else:
        problems = []
        for i, count in embedded_counts.items():
            modes = [str(d.get("mode", "")) for d in declared.get(i, []) if isinstance(d, dict)]
            real_modes = [m for m in modes if m in ALLOWED_MODES]
            if count > len(real_modes):
                problems.append((i, f"{count} embedded image(s) but only {len(real_modes)} "
                                    f"declared with an allowed provenance mode "
                                    f"(declared modes: {modes or 'none'})"))
        for i, imgs in declared.items():
            for d in imgs:
                if not isinstance(d, dict):
                    continue
                mode = str(d.get("mode", ""))
                if mode not in ALLOWED_MODES:
                    # reference-only / placeholder: only a defect if it really is embedded
                    if embedded_counts.get(i, 0) > 0:
                        problems.append((i, f"declared mode {mode!r} for {d.get('source_ref')} "
                                            f"is not an allowed provenance mode"))
        if problems:
            g.add("V06", "every embedded image is source-crop or redraw-from-data", "FAIL",
                  "; ".join(f"slide {s}: {d}" for s, d in problems[:12]),
                  find_key_line(build_path, "images"), defects=[
                      {"severity": "high", "slide": s, "detail": d} for s, d in problems])
        else:
            total = sum(embedded_counts.values())
            g.add("V06", "every embedded image is source-crop or redraw-from-data", "PASS",
                  f"{total} embedded image(s) counted in the deck; every one has a declared "
                  f"provenance mode in {sorted(ALLOWED_MODES)}. Placeholder/invented graphics "
                  f"are excluded by construction; note this proves provenance *mode*, not that "
                  f"the pixels equal the paper's figure.",
                  find_key_line(build_path, "images") if build_path.is_file() else f"{deck_path}:1")

    # ---------------- template-derived checks: V07/V08/V09 ---------------- #
    case_dir = Path(case_path).parent if case_path else None
    template = resolve_template_path(case, plan, build, (case_dir, work))
    if not deck_parts:
        for cid, nm in (("V07", "template masters/layouts not modified"),
                        ("V08", "theme fonts unchanged from template"),
                        ("V09", "footer/slide-number placeholders still present and unfilled")):
            g.add(cid, nm, "UNVERIFIED", "The deck's OOXML parts could not be read.",
                  f"{deck_path}:1")
    elif template is None:
        for cid, nm in (("V07", "template masters/layouts not modified"),
                        ("V08", "theme fonts unchanged from template"),
                        ("V09", "footer/slide-number placeholders still present and unfilled")):
            g.add(cid, nm, "UNVERIFIED",
                  "The template .pptx could not be located (checked plan.constraints."
                  "template_path, build_report.template.path, case.constraints.template_path and "
                  "case.template, each as given and relative to the case file's directory and "
                  "to the work dir), so the deck's parts cannot be compared against their "
                  "source.",
                  find_key_line(build_path, "template") if build_path.is_file() else f"{plan_path}:1")
    else:
        tparts = zip_parts(template)
        # ---- V07 ----
        if not tparts:
            g.add("V07", "template masters/layouts not modified", "UNVERIFIED",
                  f"template parts unreadable: {template}", f"{template}:1")
        else:
            changed, missing_parts, compared = [], [], 0
            for group in (master_names(tparts), layout_names(tparts)):
                for name in group:
                    if name not in deck_parts:
                        missing_parts.append(name)
                        continue
                    compared += 1
                    if canonical_xml(deck_parts[name]) != canonical_xml(tparts[name]):
                        changed.append(name)
            ev = (xml_line(part_text(deck_parts, changed[0]), "<?xml", deck_path, changed[0])
                  if changed else f"{deck_path}:1")
            if missing_parts:
                g.add("V07", "template masters/layouts not modified", "FAIL",
                      f"parts present in the template are missing from the deck: "
                      f"{missing_parts}", ev,
                      defects=[{"severity": "high",
                                "detail": f"deleted template part {p}"} for p in missing_parts])
            elif changed:
                g.add("V07", "template masters/layouts not modified", "FAIL",
                      f"{len(changed)} template part(s) differ from the template's own XML: "
                      f"{changed[:8]}", ev,
                      defects=[{"severity": "high",
                                "detail": f"modified template part {p}"} for p in changed])
            else:
                g.add("V07", "template masters/layouts not modified", "PASS",
                      f"{compared} master/layout part(s) byte-normalised-identical (C14N) "
                      f"between {template.name} and the deck's own OOXML; this is an "
                      f"independent XML comparison, not the builder's invariants flag",
                      f"{deck_path}:1")

        # ---- V08 ----
        t_fonts = theme_fonts(tparts)
        d_fonts = theme_fonts(deck_parts)
        if not any(t_fonts.values()):
            g.add("V08", "theme fonts unchanged from template", "UNVERIFIED",
                  "The template's theme declares no fontScheme that could be read.",
                  f"{template}:1")
        elif t_fonts != d_fonts:
            diff = {k: (t_fonts.get(k), d_fonts.get(k)) for k in t_fonts
                    if t_fonts.get(k) != d_fonts.get(k)}
            g.add("V08", "theme fonts unchanged from template", "FAIL",
                  f"theme fontScheme differs (template -> deck): {diff}",
                  xml_line(part_text(deck_parts, theme_names(deck_parts)[0] if theme_names(deck_parts) else ""),
                           "fontScheme", deck_path,
                           theme_names(deck_parts)[0] if theme_names(deck_parts) else "ppt/theme/theme1.xml"),
                  defects=[{"severity": "high",
                            "detail": f"theme font {k} changed from {v[0]!r} to {v[1]!r}"}
                           for k, v in diff.items()])
        else:
            g.add("V08", "theme fonts unchanged from template", "PASS",
                  f"major/minor latin+ea typefaces identical to the template: {t_fonts}",
                  xml_line(part_text(deck_parts, theme_names(deck_parts)[0] if theme_names(deck_parts) else ""),
                           "fontScheme", deck_path,
                           theme_names(deck_parts)[0] if theme_names(deck_parts) else "ppt/theme/theme1.xml"))

        # ---- V09 ----
        footer_types = ('type="ftr"', 'type="sldNum"', 'type="dt"')
        template_has = any(any(t in part_text(tparts, n) for t in footer_types)
                           for n in master_names(tparts) + layout_names(tparts))
        if not template_has:
            g.add("V09", "footer/slide-number placeholders still present and unfilled", "SKIP",
                  "The template itself declares no footer, slide-number or date placeholder, so "
                  "there is nothing to preserve.", f"{template}:1")
        else:
            still_present = any(any(t in part_text(deck_parts, n) for t in footer_types)
                               for n in master_names(deck_parts) + layout_names(deck_parts))
            filled = []
            for n, xml in deck_slides_xml:
                for ph in re.finditer(r'<p:ph[^>]*type="(ftr|sldNum|dt)"[^>]*/>', xml):
                    after = xml[ph.end():ph.end() + 900]
                    text = "".join(re.findall(r"<a:t>([^<]*)</a:t>", after.split("</p:sp>")[0]))
                    if text.strip():
                        filled.append((n, ph.group(1), text.strip()[:60]))
            if not still_present:
                g.add("V09", "footer/slide-number placeholders still present and unfilled", "FAIL",
                      "the template carries footer/slide-number placeholders but they are gone "
                      "from the deck's masters/layouts", f"{deck_path}:1",
                      defects=[{"severity": "medium",
                                "detail": "footer/slideNumber placeholders removed"}])
            elif filled:
                g.add("V09", "footer/slide-number placeholders still present and unfilled", "FAIL",
                      f"the builder wrote text into footer/slide-number placeholders: {filled[:6]}",
                      f"{deck_path}:1",
                      defects=[{"severity": "medium", "slide": n,
                                "detail": f"placeholder {t} filled with {txt!r}"}
                               for n, t, txt in filled])
            else:
                g.add("V09", "footer/slide-number placeholders still present and unfilled", "PASS",
                      "footer/slide-number/date placeholders survive in the deck's masters and "
                      "layouts and no slide writes text into them; the deck inherits them from "
                      "the template rather than replacing them", f"{deck_path}:1")

    # ---------------- V10 no text overflow ---------------- #
    if not slides:
        g.add("V10", "no text overflow (estimated extent <= shape extent)", "UNVERIFIED",
              "No readable deck, so no shape geometry could be measured.", f"{deck_path}:1")
    else:
        offenders = []
        shrunk = []
        measured = 0
        slide_h = int(getattr(prs, "slide_height", 0) or 0)
        slide_w = int(getattr(prs, "slide_width", 0) or 0)
        for i, s in enumerate(slides, 1):
            for sh in s.shapes:
                overflow, detail, needed, avail = shape_text_extent(sh)
                if overflow is None:
                    continue
                measured += 1
                if not overflow:
                    continue
                kind = shape_autofit_kind(sh)
                name = ""
                try:
                    name = (sh.name or "")[:40]
                except Exception:  # noqa: BLE001
                    name = ""
                if kind == "norm":
                    # text is shrunk to fit: not clipped, but possibly tiny
                    shrunk.append((i, f"{name or 'shape'}: {detail} (normAutofit shrinks text)"))
                    continue
                if kind == "sp":
                    # the shape grows: the real question is whether it leaves the slide
                    try:
                        top = int(sh.top or 0)
                        left = int(sh.left or 0)
                        width = int(sh.width or 0)
                    except Exception:  # noqa: BLE001
                        top = left = width = 0
                    grown_bottom = top + needed
                    if (slide_h and grown_bottom > slide_h + 1) or \
                            (slide_w and left + width > slide_w + 1):
                        offenders.append(
                            (i, f"{name or 'shape'}: spAutoFit grows to "
                                f"{grown_bottom / EMU_PER_INCH:.2f}in bottom / slide is "
                                f"{slide_h / EMU_PER_INCH:.2f}in tall -- the content runs off "
                                f"the slide ({detail})"))
                    continue
                offenders.append((i, f"{name or 'shape'}: {detail} (no autofit, so text is "
                                     f"clipped rather than shrunk)"))
        # rendered cross-check: no word may fall outside the page box
        render_note = "render cross-check not available (no PDF)"
        render_bad = []
        pdf_path = ((render.get("pdf") or {}).get("path") or "")
        if pdf_path and Path(pdf_path).is_file() and pdfplumber is not None:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    n_pages = len(pdf.pages)
                    for pi, page in enumerate(pdf.pages, 1):
                        for w in page.extract_words():
                            if w.get("bottom", 0) > page.height + 1 or \
                                    w.get("x1", 0) > page.width + 1 or \
                                    w.get("x0", 0) < -1:
                                render_bad.append((pi, w.get("text", "")))
                    render_note = f"rendered PDF cross-checked: {n_pages} page(s), " \
                                  f"{len(render_bad)} word(s) outside the page box"
            except Exception as exc:  # noqa: BLE001
                render_note = f"rendered PDF could not be parsed for the cross-check: {exc}"
        elif pdf_path:
            render_note = "rendered PDF present but pdfplumber is unavailable"
        evidence = f"{deck_path}:1"
        if offenders or render_bad:
            detail = "; ".join(f"slide {s}: {d}" for s, d in offenders[:10])
            if render_bad:
                detail += (f" | rendered words outside the page box: {render_bad[:5]}")
            g.add("V10", "no text overflow (estimated extent <= shape extent)", "FAIL",
                  detail + f" | {render_note}", evidence,
                  defects=[{"severity": "medium", "slide": s, "detail": d}
                           for s, d in offenders] +
                          [{"severity": "high", "slide": p,
                            "detail": f"rendered word {t!r} falls outside the page box"}
                           for p, t in render_bad])
        else:
            note = f"{measured} text shape(s) measured from deck geometry with real font " \
                   f"advances; none exceed their box or leave the slide. Shapes carrying " \
                   f"normAutofit shrink their text (not clipped) and spAutoFit shapes were " \
                   f"checked against the slide box instead of their own. {render_note}"
            if shrunk:
                note += f" | {len(shrunk)} shape(s) rely on normAutofit shrink-to-fit: " \
                        + "; ".join(f"slide {s}: {d}" for s, d in shrunk[:5])
            g.add("V10", "no text overflow (estimated extent <= shape extent)", "PASS",
                  note, evidence)

    # ---------------- V11 images inside their shape bounds ---------------- #
    if not slides:
        g.add("V11", "images inside their shape bounds (not cropped)", "UNVERIFIED",
              "No readable deck.", f"{deck_path}:1")
    else:
        problems = []
        checked = 0
        for i, s in enumerate(slides, 1):
            for sh in s.shapes:
                try:
                    is_pic = (sh.shape_type is not None and "PICTURE" in str(sh.shape_type)) \
                        or hasattr(sh, "image")
                except Exception:  # noqa: BLE001
                    is_pic = False
                if not is_pic:
                    continue
                checked += 1
                try:
                    xml = sh._element.xml
                except Exception:  # noqa: BLE001
                    xml = ""
                # A picture's fill is <p:blipFill> while the crop is <a:srcRect>; match the
                # tag name namespace-agnostically so both spellings are caught.
                m = re.search(r"<(?:\w+:)?srcRect([^/>]*)/?>", xml)
                if m and re.search(r'(l|t|r|b)="(?!0")', m.group(1)):
                    problems.append((i, f"cropped via srcRect{m.group(1).strip()}"))
                if prs is not None:
                    try:
                        left, top = int(sh.left or 0), int(sh.top or 0)
                        right = left + int(sh.width or 0)
                        bottom = top + int(sh.height or 0)
                        if left < -1 or top < -1 or right > int(prs.slide_width or 0) + 1 or \
                                bottom > int(prs.slide_height or 0) + 1:
                            problems.append((i, f"frame extends outside the slide "
                                                f"(l={left},t={top},r={right},b={bottom})"))
                    except Exception:  # noqa: BLE001
                        pass
                if Image is not None:
                    try:
                        with sh.image as img:
                            pass
                    except Exception:  # noqa: BLE001
                        pass
        if not checked:
            g.add("V11", "images inside their shape bounds (not cropped)", "PASS",
                  "The deck embeds no picture shapes, so no image can be cropped; vacuously "
                  "satisfied over 0 images (this proves nothing about figures that were not "
                  "embedded).", f"{deck_path}:1")
        elif problems:
            g.add("V11", "images inside their shape bounds (not cropped)", "FAIL",
                  "; ".join(f"slide {s}: {d}" for s, d in problems[:12]), f"{deck_path}:1",
                  defects=[{"severity": "medium", "slide": s, "detail": d}
                           for s, d in problems])
        else:
            g.add("V11", "images inside their shape bounds (not cropped)", "PASS",
                  f"{checked} picture shape(s) checked: no non-zero srcRect crop and every "
                  f"frame lies inside the slide box, read from the shapes' own OOXML",
                  f"{deck_path}:1")

    # ---------------- V12 fonts resolvable + substitution reported ---------------- #
    named = fonts_used_in_deck(deck_parts) if deck_parts else []
    if not named:
        g.add("V12", "fonts used are resolvable and substitution is reported", "UNVERIFIED",
              "No typeface could be attributed to the deck (neither an explicit typeface nor a "
              "theme major/minor face), so resolvability cannot be inspected.",
              f"{deck_path}:1")
    else:
        unresolvable = [f for f in named if resolve_font(f) is None]
        reported = set()
        for f in (render.get("missing_fonts") or []):
            reported.add(str(f))
        for f in ((render.get("uncertainty") or [])):
            for name in named:
                if name in str(f):
                    reported.add(name)
        for f in (build.get("fonts_used") or []):
            if isinstance(f, dict) and f.get("missing"):
                reported.add(str(f.get("name", "")))
        for w in (build.get("warnings") or []):
            for name in named:
                if name in str(w):
                    reported.add(name)
        hidden = [f for f in unresolvable if f not in reported]
        if not unresolvable:
            g.add("V12", "fonts used are resolvable and substitution is reported", "PASS",
                  f"all {len(named)} declared typefaces resolve to a font file on this host: "
                  f"{named[:8]}", f"{deck_path}:1")
        elif hidden:
            g.add("V12", "fonts used are resolvable and substitution is reported", "FAIL",
                  f"unresolvable typeface(s) {hidden} are not recorded as substituted anywhere "
                  f"(render_report.missing_fonts currently reports {sorted(reported) or 'nothing'})",
                  f"{deck_path}:1",
                  defects=[{"severity": "medium", "detail": f"font {f!r} substituted silently"}
                           for f in hidden])
        else:
            g.add("V12", "fonts used are resolvable and substitution is reported", "PASS",
                  f"{len(unresolvable)} declared typeface(s) are not installed on this host "
                  f"({sorted(unresolvable)}) and every one of them is openly reported as "
                  f"substituted, so the substitution is visible rather than hidden",
                  f"{render_path}:1" if render_path.is_file() else f"{deck_path}:1")

    # ---------------- V13 result.json fills every output_contract key ---------------- #
    contract = {}
    if isinstance(case, dict):
        oc = case.get("output_contract")
        if isinstance(oc, dict):
            contract = oc
    if not contract:
        g.add("V13", "result.json contains every key declared in case.json.output_contract",
              "UNVERIFIED",
              "case.json declares no output_contract dict, so there is nothing to compare "
              "(this is a case-shape problem, not a pass).",
              find_key_line(case_path, "output_contract") if case_path else f"{result_path}:1")
    elif not result:
        # Not the same failure as a result.json that exists and is wrong.  A missing file means
        # the producing stage (result_contract.py) has not run yet or was not wired in; that is
        # "we could not inspect it", i.e. UNVERIFIED -- not a FAIL against the skill's output.
        # Both still fail the gate (CONTRACT.md:196), so this is a semantics fix, not a waiver.
        searched = [str(p) for p in ((Path(out) / "result.json") if out else [],
                                     work / "result.json") if str(p)]
        g.add("V13", "result.json contains every key declared in case.json.output_contract",
              "UNVERIFIED",
              f"result.json was not found (looked at: {searched}); the producing stage "
              f"result_contract.py has not run yet, so the {len(contract)} declared key(s) "
              f"{sorted(contract)} could not be inspected. Run result_contract.py before "
              f"verify_deck.py, or pass --out so the gate looks in the output directory.",
              f"{result_path}:1",
              defects=[{"severity": "high",
                        "detail": f"result.json absent; producer stage did not run ({k} "
                                  f"could not be checked)"} for k in sorted(contract)])
    else:
        missing = [k for k in contract if k not in result]
        if missing:
            g.add("V13", "result.json contains every key declared in case.json.output_contract",
                  "FAIL",
                  f"result.json is missing {len(missing)} declared key(s): {missing}",
                  find_key_line(result_path, sorted(result)[0]) if result else f"{result_path}:1",
                  defects=[{"severity": "high", "detail": f"output_contract key {k!r} absent "
                                                         f"from result.json"} for k in missing])
        else:
            nulls = [k for k in contract if result.get(k) in (None, [], {})]
            uncertainty_text = [str(u) for u in (result.get("uncertainty") or [])]
            unexplained = [k for k in nulls
                           if not any(str(k) in u for u in uncertainty_text)]
            # A result.json left over from an earlier run would satisfy this check spuriously,
            # so say so plainly rather than presenting a possibly-provisional file as final.
            stale_note = ""
            try:
                r_mtime = result_path.stat().st_mtime
                newer = [name for name, p in (("deck.pptx", deck_path),
                                              ("plan.json", plan_path),
                                              ("render_report.json", render_path))
                         if p.is_file() and p.stat().st_mtime > r_mtime]
                if newer:
                    stale_note = (f" | NOTE: this result.json is older than {newer}, so it may "
                                  f"predate the current artefacts (expected when the pipeline "
                                  f"runs result_contract twice: provisional before the gate, "
                                  f"final after it)")
            except OSError:
                pass
            if unexplained:
                g.add("V13", "result.json contains every key declared in "
                             "case.json.output_contract", "FAIL",
                      f"key(s) {unexplained} are present but null/empty and no uncertainty "
                      f"entry mentions them; CONTRACT.md:33-34 requires every undeliverable key "
                      f"to carry an explanation rather than just being blank",
                      find_key_line(result_path, sorted(contract)[0]),
                      defects=[{"severity": "medium",
                                "detail": f"declared key {k!r} is empty with no explanation"}
                               for k in unexplained])
            else:
                g.add("V13", "result.json contains every key declared in "
                             "case.json.output_contract", "PASS",
                      f"all {len(contract)} declared key(s) are present in result.json"
                      + (f"; {len(nulls)} are null/empty and each is explained in uncertainty: "
                         f"{nulls}" if nulls else "") + stale_note,
                      find_key_line(result_path, sorted(contract)[0]))

    # ---------------- V14 evidence map covers every slide ---------------- #
    if not emap:
        g.add("V14", "evidence map covers every slide", "UNVERIFIED",
              f"evidence_map.json is missing or empty at {evidence_path}.",
              f"{evidence_path}:1")
    else:
        pages = emap.get("pages") or []
        if not pages:
            g.add("V14", "evidence map covers every slide", "FAIL",
                  "evidence_map.json has no pages array.", find_key_line(evidence_path, "pages"),
                  defects=[{"severity": "high", "detail": "evidence map has no pages"}])
        else:
            have = sorted({int(p.get("slide", 0)) for p in pages if isinstance(p, dict)})
            expected = list(range(1, (deck_slide_count or len(have)) + 1))
            missing = [n for n in expected if n not in have]
            thin = []
            for p in pages:
                if not isinstance(p, dict):
                    continue
                items = (len(p.get("claims") or []) + len(p.get("figures") or [])
                         + len(p.get("tables") or []) + len(p.get("numbers") or []))
                if items == 0:
                    thin.append(p.get("slide"))
            if missing:
                g.add("V14", "evidence map covers every slide", "FAIL",
                      f"no evidence entry for slide(s) {missing} (map covers {have}, deck has "
                      f"{deck_slide_count} slides)", find_key_line(evidence_path, "pages"),
                      defects=[{"severity": "high", "slide": n,
                                "detail": "slide has no evidence entry"} for n in missing])
            elif thin:
                g.add("V14", "evidence map covers every slide", "FAIL",
                      f"slide(s) {thin} have an evidence entry but zero claims/figures/tables/"
                      f"numbers, so the slide is not actually evidenced",
                      find_key_line(evidence_path, "claims"),
                      defects=[{"severity": "medium", "slide": n,
                                "detail": "evidence entry is empty"} for n in thin])
            else:
                verified = sum(1 for p in pages for c in (p.get("claims") or [])
                               if isinstance(c, dict) and c.get("verified"))
                g.add("V14", "evidence map covers every slide", "PASS",
                      f"all {len(expected)} slide(s) have a non-empty evidence entry with a "
                      f"locator ({verified} claim(s) marked verified)",
                      find_key_line(evidence_path, "pages"))

    # ---------------- V15 version records ---------------- #
    # The output dir is authoritative when --out was passed: run.py copies the version records
    # there, so looking only in --work produced a false UNVERIFIED while the files sat in --out.
    search_dirs = ([Path(out), work] if out else [work])
    patterns = ("version*.json", "versions*.json", "change_records*.json", "changes*.json",
                "diff*.md", "diff*.json", "rollback*.md", "rollback*.txt", "VERSION", "CHANGELOG*")
    found_files = []
    for directory in search_dirs:
        if not Path(directory).is_dir():
            continue
        for pat in patterns:
            try:
                found_files += [p for p in Path(directory).glob(pat) if p.is_file()]
            except OSError:
                continue
    found_files = sorted({str(p) for p in found_files})
    version_records = []
    for path in found_files:
        data = load_json(path)
        if isinstance(data, dict):
            version_records.append((path, data))
    has_v1 = False
    has_changes = False
    has_diff = False
    has_rollback = False
    for path, data in version_records:
        blob = json.dumps(data, ensure_ascii=False)
        if '"v1"' in blob or "'v1'" in blob or data.get("version") == "v1" or \
                any(str(v.get("version", "")).startswith("v1")
                    for v in (data.get("versions") or []) if isinstance(v, dict)):
            has_v1 = True
        if data.get("changes") or data.get("change_records"):
            has_changes = True
        if "diff" in blob.lower():
            has_diff = True
        if "rollback" in blob.lower() or "revert" in blob.lower():
            has_rollback = True
    for path in found_files:
        low = Path(path).name.lower()
        if low.startswith("diff"):
            has_diff = True
        if low.startswith("rollback"):
            has_rollback = True
        if low.startswith("change"):
            has_changes = True
    # result.json's own change_records counts as a change record (CONTRACT.md:213)
    if result.get("change_records"):
        has_changes = True
    if not found_files and not result.get("change_records"):
        g.add("V15", "version records exist (v1 + changes + diff + rollback)", "UNVERIFIED",
              "No version/change/diff/rollback artefact was found in "
              f"{[str(d) for d in search_dirs]} (patterns: {patterns}) and result.json carries "
              f"no change_records, so version history cannot be attested.",
              f"{work}:1")
    else:
        gaps = []
        if not has_v1:
            gaps.append("no v1 version record")
        if not has_changes:
            gaps.append("no change records")
        if not has_diff:
            gaps.append("no diff material")
        if not has_rollback:
            gaps.append("no rollback instruction")
        ev = file_line(found_files[0], "schema") if found_files else f"{result_path}:1"
        if gaps:
            g.add("V15", "version records exist (v1 + changes + diff + rollback)", "FAIL",
                  f"version artefacts present ({[Path(p).name for p in found_files]}) but "
                  f"incomplete: {gaps}", ev,
                  defects=[{"severity": "medium", "detail": gap} for gap in gaps])
        else:
            g.add("V15", "version records exist (v1 + changes + diff + rollback)", "PASS",
                  f"v1 + change records + diff + rollback all present across "
                  f"{[Path(p).name for p in found_files]}"
                  + (" (+ result.json.change_records)" if result.get("change_records") else ""),
                  ev)

    # ---------------- V16 no duplicate slide content ---------------- #
    units = slide_text_units(plan_slides, slides)
    if not units:
        g.add("V16", "no duplicate slide content", "UNVERIFIED",
              "No plan bullets and no readable deck text, so duplication cannot be inspected.",
              f"{deck_path}:1")
    else:
        dups = duplicate_content_issues(units)
        if dups:
            g.add("V16", "no duplicate slide content", "FAIL",
                  f"{len(dups)} pair(s) of slides reuse content: "
                  + " | ".join(d for _, _, d in dups[:8]),
                  find_key_line(plan_path, "bullets") if plan_path.is_file() else f"{deck_path}:1",
                  defects=[{"severity": "high", "slide": si,
                            "detail": f"duplicates slide {sj}: {d}"} for si, sj, d in dups])
        else:
            g.add("V16", "no duplicate slide content", "PASS",
                  f"{len(units)} title/bullet/notes unit(s) checked across "
                  f"{max([u[0] for u in units] or [0])} slide(s): no bullet fingerprint repeats "
                  f"(a leading 'Label:' prefix is stripped first, and bullets under "
                  f"{MIN_REPEAT_CHARS} characters are ignored so short role labels are not false "
                  f"positives); no title/bullet is identical to, contained in, or "
                  f">={CONTENT_OVERLAP:.0%} content-word-overlapping with another's; no notes "
                  f"are exactly repeated",
                  find_key_line(plan_path, "bullets") if plan_path.is_file() else f"{deck_path}:1")

    # ---------------- V17 no constraint/input leakage into content ---------------- #
    meta = collect_constraint_metadata(case, plan)
    if not units:
        g.add("V17", "no constraint/input leakage into content", "UNVERIFIED",
              "No slide text to scan.", f"{deck_path}:1")
    elif not meta:
        g.add("V17", "no constraint/input leakage into content", "UNVERIFIED",
              "The case and plan declare no constraint metadata (audience, duration, "
              "template_path, language, slide_count), so there is nothing to leak and "
              "nothing to compare against.",
              f"{case_path}:1" if case_path else f"{plan_path}:1")
    else:
        leaks = constraint_leak_issues(units, meta)
        if leaks:
            g.add("V17", "no constraint/input leakage into content", "FAIL",
                  "the brief's own metadata was copied onto slides: "
                  + " | ".join(f"slide {s}: {d}" for s, d in leaks[:8]),
                  find_key_line(case_path, "audience") if case_path else f"{plan_path}:1",
                  defects=[{"severity": "medium", "slide": s, "detail": d} for s, d in leaks])
        else:
            g.add("V17", "no constraint/input leakage into content", "PASS",
                  f"{len(units)} unit(s) scanned against the declared metadata "
                  f"({sorted(meta)}): no slide title, bullet or note carries a labelled brief "
                  f"field, the declared audience text, the template file name, or the declared "
                  f"duration",
                  find_key_line(case_path, "audience") if case_path else f"{plan_path}:1")

    # ---------------- V18 (extra) build_report describes the deck under inspection -------- #
    if deck_mismatch:
        g.add("V18", "build_report describes the deck under inspection", "FAIL",
              f"build_report.json points at {deck_mismatch} but the work dir holds "
              f"{deck_path}; the two disagree, so the builder's report does not describe the "
              f"deck being verified (the work-dir deck was inspected, not the foreign one)",
              find_key_line(build_path, "pptx"),
              defects=[{"severity": "high",
                        "detail": f"build_report.pptx {deck_mismatch!r} != inspected "
                                  f"{deck_path!r}"}])
    elif (build or {}).get("pptx"):
        g.add("V18", "build_report describes the deck under inspection", "PASS",
              "build_report.pptx names the same deck file that was inspected",
              find_key_line(build_path, "pptx"))
    else:
        g.add("V18", "build_report describes the deck under inspection", "SKIP",
              "build_report.json declares no pptx path, so there is nothing to reconcile.",
              f"{build_path}:1")

    # ---------------- V19 minimum content coverage (anti false-green) ---------------- #
    # The class this exists to kill: if ingest drops most of the source (a scanned PDF with no
    # text layer, a bad decode), then every downstream check over an empty list is vacuously
    # satisfied -- `all()` of nothing is true -- and a deck of template boilerplate ships green.
    # Thresholds below were fixed by measuring the two real fixtures first (see
    # references/verification_gates.md section 4): caseA 8 slides / 14 bullets / 13 880
    # extracted chars / 4 figures / 2 tables / 370 numbers, 21 distinct numbers on slides, no
    # unreferenced-... -> and caseB 12 slides / 17 bullets / 13 091 chars / 3 figures / 2 tables
    # / 143 numbers, 22 distinct numbers on slides.
    MIN_EXTRACTED_CHARS = 200
    DIVIDER_ROLES = {"title", "section", "outline", "agenda", "ack", "acknowledgement",
                     "thanks", "thankyou", "backup", "appendix", "references", "refs", "qa",
                     "divider", "closing", "cover"}

    def _source_bytes(case_obj, material_obj, bases):
        paths, total = [], 0
        raw = []
        for s in (case_obj or {}).get("sources") or []:
            raw.append(s if isinstance(s, str) else (s or {}).get("path"))
        for s in (material_obj or {}).get("sources") or []:
            if isinstance(s, dict):
                raw.append(s.get("path"))
        seen = set()
        for item in raw:
            if not item:
                continue
            found = resolve_existing(item, bases)
            if found is None:
                continue
            try:
                key = str(found.resolve())
            except OSError:
                key = str(found)
            if key in seen:
                continue
            seen.add(key)
            paths.append(key)
            try:
                total += found.stat().st_size
            except OSError:
                pass
        return total, paths

    extracted_chars = sum(len(str(p.get("text") or ""))
                          for s in (material.get("sources") or []) if isinstance(s, dict)
                          for p in (s.get("pages") or []) if isinstance(p, dict))
    n_fig_src = len(material.get("figures") or [])
    n_tab_src = len(material.get("tables") or [])
    n_num_src = len(material.get("numbers") or [])
    n_claim_src = len(material.get("claims") or [])
    src_bytes, src_paths = _source_bytes(case, material, (case_dir, work))
    # An unreferenced figure/table counts as explained when its id appears in an
    # `excluded[].item` **or** its `.reason`.  deck-core writes items such as
    # "fig2 (figure Figure 2) not placed: …", so an exact-set membership test looked for the
    # bare id in a set of whole item strings, found nothing, and printed the false statement
    # "not listed in plan.excluded" into a shipped artefact.  Word boundaries keep `fig2` from
    # being satisfied by `fig20`.
    excluded_text = "\n".join(
        f"{e.get('item', '')} {e.get('reason', '')}"
        for e in (plan.get("excluded") or []) if isinstance(e, dict))

    def _is_explained(item_id):
        if not excluded_text:
            return False
        pattern = rf"(?<![\w\-]){re.escape(str(item_id))}(?![\w\-])"
        return re.search(pattern, excluded_text) is not None

    ref_figs, ref_tabs = set(), set()
    for ps in plan_slides:
        if not isinstance(ps, dict):
            continue
        if ps.get("figure_ref"):
            ref_figs.add(str(ps["figure_ref"]))
        if ps.get("table_ref"):
            ref_tabs.add(str(ps["table_ref"]))
        for ev in ps.get("evidence") or []:
            if isinstance(ev, dict) and ev.get("ref"):
                if ev.get("type") == "figure":
                    ref_figs.add(str(ev["ref"]))
                elif ev.get("type") == "table":
                    ref_tabs.add(str(ev["ref"]))

    deck_numbers = set()
    for _slide_no, kind, norm, raw in units:
        if kind == "notes":
            continue
        deck_numbers.update(numbers_on_text(raw))

    content_units = [u for u in units if u[1] != "notes"]
    bullets_total = sum(len(ps.get("bullets") or []) for ps in plan_slides
                        if isinstance(ps, dict))

    title_only = []
    for ps in plan_slides:
        if not isinstance(ps, dict):
            continue
        if (ps.get("bullets") or ps.get("figure_ref") or ps.get("table_ref")
                or ps.get("numbers_used")):
            continue
        role = str(ps.get("role") or "").strip().lower()
        if role in DIVIDER_ROLES:
            continue
        title_only.append(ps.get("index"))

    coverage = {
        "slides": len(plan_slides), "bullets_total": bullets_total,
        "extracted_chars": extracted_chars, "declared_source_files": len(src_paths),
        "declared_source_bytes": src_bytes,
        "material": {"figures": n_fig_src, "tables": n_tab_src, "numbers": n_num_src,
                     "claims": n_claim_src},
        "deck_numbers": len(deck_numbers),
        "figures_referenced": sorted(ref_figs), "tables_referenced": sorted(ref_tabs),
        "title_only_content_slides": title_only,
    }
    report_coverage = coverage

    if not material and not src_paths:
        g.add("V19", "minimum content coverage", "SKIP",
              "N/A: the case declares no source material and no material.json exists, so there "
              "is no coverage to measure. Recorded as N/A (SKIP with reason), not as PASS.",
              f"{material_path}:1")
    else:
        blocking = []
        # R1 -- the false-green engine: a real source file that yielded (almost) nothing.
        if src_paths and src_bytes > 0 and extracted_chars < MIN_EXTRACTED_CHARS \
                and not any((n_fig_src, n_tab_src, n_num_src, n_claim_src)):
            blocking.append(
                (f"R1 source read but nothing extracted: {len(src_paths)} declared source "
                 f"file(s) totalling {src_bytes} B yielded only {extracted_chars} chars of text "
                 f"and 0 figures/tables/numbers/claims. A scanned PDF with no text layer or a "
                 f"failed decode presents exactly this way, and every other gate would pass "
                 f"vacuously over the empty lists", None))
        # R2 -- numbers exist in the source but none ever reaches a slide.
        if len(plan_slides) >= 4 and n_num_src > 0 and not deck_numbers:
            blocking.append(
                (f"R2 {n_num_src} number(s) were extracted from the source but 0 distinct "
                 f"numeric tokens appear on any of the {len(plan_slides)} slides", None))
        # R3 -- the deck is template boilerplate: source had content, slides have no bullets.
        if len(plan_slides) >= 3 and (n_fig_src + n_tab_src + n_num_src) > 0 \
                and bullets_total == 0:
            blocking.append(
                (f"R3 the source supplied {n_fig_src} figure(s), {n_tab_src} table(s) and "
                 f"{n_num_src} number(s) but the deck carries 0 bullets across "
                 f"{len(plan_slides)} slides", None))
        # R4 -- a content page whose only content is its title.
        if len(plan_slides) >= 5 and title_only:
            blocking.append(
                (f"R4 slide(s) {title_only} have a content role but nothing but a title: no "
                 f"bullets, no figure, no table and no numbers. If a page is a divider, give it "
                 f"a divider role so this check can tell the two apart", title_only[0]))
        # R5 -- every figure dropped (complete loss, not partial curation).
        if n_fig_src >= 2 and not ref_figs:
            blocking.append(
                (f"R5 the source supplied {n_fig_src} figures and the deck references none of "
                 f"them", None))
        if n_tab_src >= 2 and not ref_tabs:
            blocking.append(
                (f"R5 the source supplied {n_tab_src} tables and the deck references none of "
                 f"them", None))
        # WARN -- partial drops that plan.excluded does not account for (advisory, not blocking).
        dropped = sorted(({str(f.get("id")) for f in material.get("figures") or []
                           if isinstance(f, dict)} | 
                          {str(t.get("id")) for t in material.get("tables") or []
                           if isinstance(t, dict)}) - (ref_figs | ref_tabs))
        unaccounted = [d for d in dropped if not _is_explained(d)]
        if unaccounted:
            g.warn("V19", f"figure/table item(s) {unaccounted} from the source are never "
                          f"referenced and are not listed in plan.excluded; this is advisory "
                          f"(a deck need not use every figure) but the audit trail is thinner "
                          f"than it could be")

        detail = (f"coverage: {len(plan_slides)} slide(s), {bullets_total} bullet(s), "
                  f"{len(deck_numbers)} distinct number(s) on slides, figures used "
                  f"{len(ref_figs)}/{n_fig_src}, tables used {len(ref_tabs)}/{n_tab_src}, "
                  f"extracted {extracted_chars} chars from {len(src_paths)} source file(s). "
                  f"{len(g.warnings)} advisory note(s).")
        if blocking:
            g.add("V19", "minimum content coverage", "FAIL",
                  "; ".join(text for text, _ in blocking) + " | " + detail,
                  find_key_line(material_path, "numbers") if material_path.is_file()
                  else f"{deck_path}:1",
                  defects=[{"severity": "high", "detail": text, "slide": slide}
                           for text, slide in blocking])
        else:
            g.add("V19", "minimum content coverage", "PASS",
                  f"the run had enough to work with: {detail}",
                  find_key_line(material_path, "figures") if material_path.is_file()
                  else f"{deck_path}:1")

    counts = g.counts()
    report = {
        "schema": SCHEMA,
        "case_id": case_id or (case.get("case_id") if isinstance(case, dict) else None),
        "checks": g.checks,
        "counts": counts,
        "verdict": g.verdict(),
        "defects": g.defects,
        "warnings": g.warnings,
        "coverage": report_coverage,
        "evidence_format": "path:line for files, <container>!<part>:line for OOXML parts",
        "artifacts": {
            "deck": str(deck_path) if deck_path.is_file() else None,
            "deck_path_source": deck_source,
            "build_report_pptx": reported_deck or None,
            "deck_sha256": sha256_file(deck_path) if deck_path.is_file() else None,
            "plan": str(plan_path) if plan_path.is_file() else None,
            "material": str(material_path) if material_path.is_file() else None,
            "build_report": str(build_path) if build_path.is_file() else None,
            "render_report": str(render_path) if render_path.is_file() else None,
            "evidence_map": str(evidence_path) if evidence_path.is_file() else None,
            "result": str(result_path) if result_path.is_file() else None,
            "template": str(template) if template else None,
        },
        "scope": {
            "V07_V08_V09": "OXML-level comparison against the template's own parts, when the "
                           "template is locatable",
            "V04": "re-derived from material.json plus the text on the deck's slides, not from "
                   "plan.json",
            "V10_V11": "deck geometry; V10 additionally cross-checks the rendered PDF when one "
                       "exists",
            "not_certified": "technical text checks do not certify scientific correctness or "
                             "presentation quality (task/task.md:97)",
        },
    }
    return report


def main(argv=None):
    ap = argparse.ArgumentParser(description="Run the V01-V15 self-check gate on a built deck.")
    ap.add_argument("--case", default=None)
    ap.add_argument("--work", required=True)
    ap.add_argument("--out", default=None, help="report path, or the output dir")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    out_dir = None
    report_path = Path(args.work) / "verify_report.json"
    if args.out:
        if str(args.out).lower().endswith(".json"):
            report_path = Path(args.out)
            out_dir = Path(args.out).parent
        else:
            out_dir = Path(args.out)
            report_path = out_dir / "verify_report.json"

    report = run_checks(args.case, args.work, out_dir)
    dump_json(report_path, report)
    # When --out is a directory, also place the report beside the work artefacts: CONTRACT.md
    # section 3 says this stage writes ``W/verify_report.json``, and run.py may look there.
    work_report = Path(args.work) / "verify_report.json"
    if out_dir and str(report_path) != str(work_report):
        dump_json(work_report, report)
    if not args.quiet:
        fails = [c for c in report["checks"] if c["status"] in ("FAIL", "UNVERIFIED")]
        print(json.dumps({"verdict": report["verdict"], "counts": report["counts"],
                          "report": str(report_path),
                          "notable": [f"{c['id']}:{c['status']}" for c in fails]},
                         ensure_ascii=False))
    return 0 if report["verdict"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
