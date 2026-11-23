#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""End-to-end self-test for the ARMI:2611.000074v1 academic-PPT Skill.

Two halves, both mandatory:

**Positive** -- run the real pipeline (``scripts/run.py``) on the synthetic fixtures
(``caseA`` experimental and ``caseB`` review/theory), then assert that the declared
artifacts exist, that ``verify_report.json.verdict == "PASS"`` with zero FAIL *and* zero
UNVERIFIED checks, and that ``result.json`` fills every key declared in
``case.json.output_contract``.  ``verify_deck.py`` is then re-run independently to prove
the gate is reproducible rather than trusting ``run.py``'s own invocation.

**Negative (defect injection)** -- a good work directory is copied, one defect is injected
at the OOXML/JSON level, and ``verify_deck.py`` must **detect** it.  Every injection first
proves that it actually changed something; an injection that could not be applied is
reported as ``MUTATION_FAILED`` and counts as a failure, so the self-test can never look
green because a mutation silently did nothing.  Each negative also rewrites the absolute
paths inside the copied JSON so that the *only* difference from the good run is the
injected defect (otherwise ``V16`` would fire on the relocated path and every negative
would be "detected" for the wrong reason).

Robot line: the last stdout line is ``SELFTEST_JSON {<json>}`` with the machine-readable
summary.  Exit codes: 0 = every positive passed and every negative was detected;
1 = a positive failed or a negative was not detected; 2 = a pipeline script is missing so
nothing can be tested honestly.

Usage::

    python scripts/selftest.py [--fixtures DIR] [--tmp DIR] [--case caseA]
                               [--json OUT] [--keep] [--timeout SEC]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")

HERE = Path(__file__).resolve()
SCRIPTS = HERE.parent
SUBMISSION = SCRIPTS.parent

PIPELINE = {
    "run.py": SCRIPTS / "run.py",
    "ingest.py": SCRIPTS / "ingest.py",
    "plan.py": SCRIPTS / "plan.py",
    "build_deck.py": SCRIPTS / "build_deck.py",
    "render.py": SCRIPTS / "render.py",
    "evidence.py": SCRIPTS / "evidence.py",
    "verify_deck.py": SCRIPTS / "verify_deck.py",
    "result_contract.py": SCRIPTS / "result_contract.py",
}

# Artifacts ``run.py`` must leave in the OUTPUT (deliverable) dir -- CONTRACT.md section 4.1.
OUT_FILES = ["result.json", "deck.pptx", "plan.json", "evidence_map.json",
             "verify_report.json", "run.log"]
# Stage artifacts that must exist in the WORK dir -- CONTRACT.md section 3.
WORK_FILES = ["material.json", "plan.json", "build_report.json", "evidence_map.json",
              "render_report.json", "deck.pptx"]
MIN_NOTES = 40


# --------------------------------------------------------------------------- #
# small helpers
# --------------------------------------------------------------------------- #
def log(msg):
    print(msg, flush=True)


def emitted_gate_ids(verify_path: Path):
    """Gate ids ``verify_deck.py`` currently emits, read from its source.

    Read at run time on purpose: the gate list is growing (V01..V19), and a negative for an
    id the gate does not emit yet must be recorded as SKIP-with-reason rather than silently
    reported as detected or as failed.
    """
    try:
        src = verify_path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return set()
    return set(re.findall(r'g\.add\(\s*"(V\d{2})"', src))


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def load_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return default


def dump_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(json.dumps(obj, ensure_ascii=False, indent=2) + "\n")


def run_cli(args, cwd, timeout):
    """Run a pipeline CLI; return (returncode, stdout, stderr)."""
    try:
        proc = subprocess.run([sys.executable] + [str(a) for a in args],
                              capture_output=True, text=True, encoding="utf-8",
                              errors="replace", cwd=str(cwd), timeout=timeout)
        return proc.returncode, proc.stdout or "", proc.stderr or ""
    except subprocess.TimeoutExpired:
        return 124, "", "timeout after %ss" % timeout


def rmtree(path):
    shutil.rmtree(str(path), ignore_errors=True)


def copy_work(src, dst):
    """Copy a work dir (and any result.json) into a fresh directory."""
    rmtree(dst)
    shutil.copytree(str(src), str(dst))
    return dst


def _replace_in_obj(obj, old, new):
    """Recursively replace ``old``->``new`` inside a JSON object's strings."""
    if isinstance(obj, str):
        before = obj
        obj = obj.replace(old, new)
        obj = obj.replace(old.replace("\\", "/"), new.replace("\\", "/"))
        return obj, (0 if obj == before else 1)
    if isinstance(obj, dict):
        total = 0
        for key, value in list(obj.items()):
            obj[key], n = _replace_in_obj(value, old, new)
            total += n
        return obj, total
    if isinstance(obj, list):
        total = 0
        for i, value in enumerate(obj):
            obj[i], n = _replace_in_obj(value, old, new)
            total += n
        return obj, total
    return obj, 0


def rewrite_absolute_paths(root: Path, old: str, new: str) -> int:
    """Rewrite every occurrence of the old path prefix under ``root``.

    Required because the copy relocates the work dir: without this, build_report.pptx and
    render_report.pdf.path would still name the original directory and V16 (and V10's PDF
    cross-check) would react to the *relocation* instead of to the injected defect.

    JSON files are rewritten through the parsed object, so the Windows backslash escaping
    inside JSON cannot make the substitution silently miss.
    """
    if not old:
        return 0
    rewritten = 0
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        suffix = path.suffix.lower()
        text_exts = (".log", ".txt", ".md", ".jsonl")
        if suffix == ".json":
            obj = load_json(path, None)
            if obj is None:
                continue
            new_obj, hits = _replace_in_obj(obj, old, new)
            if hits:
                dump_json(path, new_obj)
                rewritten += 1
        elif suffix in text_exts:
            try:
                text = path.read_text(encoding="utf-8", errors="replace")
            except OSError:
                continue
            new_text, hits = _replace_in_obj(text, old, new)
            if hits:
                path.write_text(new_text, encoding="utf-8", newline="\n")
                rewritten += 1
    return rewritten


# --------------------------------------------------------------------------- #
# OOXML surgery helpers
# --------------------------------------------------------------------------- #
def rewrite_zip_part(pptx: Path, part: str, transform) -> int:
    """Apply ``transform(str)->str`` to one part of a .pptx; return substitutions done.

    Returns 0 (and leaves the file untouched) when the part is absent or the transform
    changed nothing -- callers treat 0 as MUTATION_FAILED.
    """
    if not pptx.is_file():
        return 0
    with zipfile.ZipFile(str(pptx)) as zin:
        names = zin.namelist()
        if part not in names:
            return 0
        payload = {n: zin.read(n) for n in names}
        infos = {i.filename: i for i in zin.infolist()}

    before = payload[part].decode("utf-8", errors="replace")
    after = transform(before)
    if after == before:
        return 0
    payload[part] = after.encode("utf-8")

    tmp = pptx.with_suffix(pptx.suffix + ".tmp")
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            info = infos[name]
            zout.writestr(info, payload[name])
    os.replace(str(tmp), str(pptx))
    return 1


def drop_last_slide(pptx: Path) -> dict:
    """Delete the last slide from a deck (presentation.xml + rels + part + content type)."""
    with zipfile.ZipFile(str(pptx)) as zin:
        names = zin.namelist()
        payload = {n: zin.read(n) for n in names}
        infos = {i.filename: i for i in zin.infolist()}

    pres = payload["ppt/presentation.xml"].decode("utf-8")
    ids = re.findall(r'<p:sldId\b[^>]*/>', pres)
    if len(ids) < 2:
        return {"ok": False, "why": "deck has fewer than 2 slides"}
    victim = ids[-1]
    rid = re.search(r'r:id="([^"]+)"', victim)
    if not rid:
        return {"ok": False, "why": "could not parse r:id of the last slide"}
    rid = rid.group(1)
    payload["ppt/presentation.xml"] = pres.replace(victim, "").encode("utf-8")

    rels_name = "ppt/_rels/presentation.xml.rels"
    rels = payload[rels_name].decode("utf-8")
    rel = re.search(r'<Relationship\b[^>]*Id="%s"[^>]*/>' % re.escape(rid), rels)
    if not rel:
        return {"ok": False, "why": "no relationship %s in presentation rels" % rid}
    target = re.search(r'Target="([^"]+)"', rel.group(0))
    payload[rels_name] = rels.replace(rel.group(0), "").encode("utf-8")

    slide_part = None
    if target:
        slide_part = "ppt/" + target.group(1).replace("../", "").replace("\\", "/")
        slide_part = slide_part.replace("ppt/ppt/", "ppt/")
    removed = [rid]
    if slide_part and slide_part in payload:
        del payload[slide_part]
        removed.append(slide_part)
        side = slide_part.replace("slides/", "slides/_rels/") + ".rels"
        if side in payload:
            del payload[side]
            removed.append(side)
        ct = payload["[Content_Types].xml"].decode("utf-8")
        payload["[Content_Types].xml"] = ct.replace(
            '<Override PartName="/%s"' % slide_part, '<Override PartName="/__removed__%s"' % slide_part
        ).encode("utf-8")

    tmp = pptx.with_suffix(pptx.suffix + ".tmp")
    with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
        for name in payload:
            zout.writestr(infos.get(name, name), payload[name])
    os.replace(str(tmp), str(pptx))
    return {"ok": True, "removed": removed, "slide_part": slide_part}


def deck_slide_count(pptx: Path) -> int:
    try:
        with zipfile.ZipFile(str(pptx)) as z:
            if "ppt/presentation.xml" not in z.namelist():
                return 0
            pres = z.read("ppt/presentation.xml").decode("utf-8", errors="replace")
        return len(re.findall(r"<p:sldId\b", pres))
    except Exception:
        return 0


# --------------------------------------------------------------------------- #
# defect injections (each returns {'ok': bool, 'detail': str})
# --------------------------------------------------------------------------- #
def mut_fabricated_number(work: Path, case: Path):
    from pptx import Presentation
    deck = work / "deck.pptx"
    prs = Presentation(str(deck))
    for idx, slide in enumerate(prs.slides, 1):
        if idx < 2:
            continue
        box = slide.shapes.add_textbox(prs.slide_width // 8, prs.slide_height // 8,
                                       prs.slide_width // 2, prs.slide_height // 10)
        box.text_frame.text = "Sample purity 1777.7 % (fabricated by selftest)"
        prs.save(str(deck))
        return {"ok": True, "detail": "added '1777.7 %%' to slide %d" % idx}
    return {"ok": False, "detail": "no slide 2 to inject into"}


def mut_short_notes(work: Path, case: Path):
    from pptx import Presentation
    deck = work / "deck.pptx"
    prs = Presentation(str(deck))
    target = None
    for idx, slide in enumerate(prs.slides, 1):
        if idx >= 2:
            target = slide
            break
    if target is None:
        return {"ok": False, "detail": "no slide 2"}
    tf = target.notes_slide.notes_text_frame
    tf.text = "too short"
    prs.save(str(deck))

    plan_path = work / "plan.json"
    plan = load_json(plan_path, {}) or {}
    touched = 0
    for i, s in enumerate(plan.get("slides") or []):
        if i == 1:
            s["notes"] = "too short"
            touched += 1
    if touched:
        dump_json(plan_path, plan)
    return {"ok": True,
            "detail": "deck notes slide 2 -> %r (<%d chars); plan slide 2 also shortened"
                      % ("too short", MIN_NOTES)}


def mut_slide_count(work: Path, case: Path):
    deck = work / "deck.pptx"
    before = deck_slide_count(deck)
    result = drop_last_slide(deck)
    after = deck_slide_count(deck)
    if not result.get("ok") or after >= before:
        return {"ok": False, "detail": "could not remove a slide: %s" % result}
    plan_path = work / "plan.json"
    plan = load_json(plan_path, {}) or {}
    if plan.get("slides"):
        plan["slides"] = plan["slides"][:-1]
        dump_json(plan_path, plan)
    emap_path = work / "evidence_map.json"
    emap = load_json(emap_path, {}) or {}
    if emap.get("pages"):
        emap["pages"] = emap["pages"][:-1]
        dump_json(emap_path, emap)
    case_obj = load_json(case, {}) or {}
    want = ((case_obj.get("constraints") or {}).get("slide_count"))
    return {"ok": True, "detail": "deck slides %d -> %d (case requires %s)"
                                  % (before, after, want)}


def mut_theme_modified(work: Path, case: Path):
    deck = work / "deck.pptx"
    if not deck.is_file():
        return {"ok": False, "detail": "HARNESS: no deck.pptx staged at %s" % deck}

    def transform(xml):
        return re.sub(r'(<a:majorFont>\s*<a:latin typeface=")([^"]*)(")',
                      r"\1Tampered Serif\3", xml, count=1)

    if rewrite_zip_part(deck, "ppt/theme/theme1.xml", transform):
        return {"ok": True, "detail": "theme1.xml major latin typeface -> 'Tampered Serif'"}
    return {"ok": False,
            "detail": "no a:majorFont/a:latin@typeface node found in ppt/theme/theme1.xml"}


def mut_master_modified(work: Path, case: Path):
    deck = work / "deck.pptx"
    if not deck.is_file():
        return {"ok": False, "detail": "HARNESS: no deck.pptx staged at %s" % deck}

    def transform(xml):
        # tolerate both ``name=`` and namespaced ``p:name=`` on p:cNvPr
        for pattern in (r'(<p:cNvPr\b[^>]*?\bname=")([^"]*)(")',
                        r"(<p:cNvPr\b[^>]*?\bp:name=\")([^\"]*)(\")"):
            new, count = re.subn(pattern, r"\1tampered-\2\3", xml, count=1)
            if count:
                return new
        return xml

    if rewrite_zip_part(deck, "ppt/slideMasters/slideMaster1.xml", transform):
        return {"ok": True, "detail": "slideMaster1.xml first cNvPr name prefixed 'tampered-'"}
    return {"ok": False,
            "detail": "no p:cNvPr name attribute found in ppt/slideMasters/slideMaster1.xml"}


def mut_placeholder_figure(work: Path, case: Path):
    from PIL import Image
    from pptx import Presentation
    deck = work / "deck.pptx"
    img = work / "_selftest_placeholder.png"
    Image.new("RGB", (320, 240), (200, 60, 60)).save(str(img))
    prs = Presentation(str(deck))
    slide = None
    for idx, s in enumerate(prs.slides, 1):
        if idx >= 3:
            slide = s
            break
    if slide is None:
        return {"ok": False, "detail": "no slide 3"}
    slide.shapes.add_picture(str(img), prs.slide_width // 4, prs.slide_height // 4,
                             width=prs.slide_width // 4)
    prs.save(str(deck))

    br_path = work / "build_report.json"
    build = load_json(br_path, {}) or {}
    target = None
    for entry in build.get("slides") or []:
        if int(entry.get("index", 0) or 0) >= 3:
            target = entry
            break
    if target is None:
        return {"ok": False, "detail": "build_report has no slide >= 3 entry"}
    target.setdefault("images", []).append(
        {"source_ref": "fig99", "mode": "placeholder",
         "note": "injected by selftest: invented graphic"})
    dump_json(br_path, build)
    return {"ok": True,
            "detail": "embedded an extra picture on slide 3 and declared "
                      "mode='placeholder' / source_ref='fig99'"}


def _result_candidates(work: Path):
    return [work / "result.json", work.parent / "result.json"]


def mut_missing_contract_key(work: Path, case: Path):
    case_obj = load_json(case, {}) or {}
    want = list((case_obj.get("output_contract") or {}).keys())
    if not want:
        return {"ok": False, "detail": "case.json declares no output_contract"}
    result_path = None
    for cand in _result_candidates(work):
        if cand.is_file():
            result_path = cand
            break
    if result_path is None:
        return {"ok": False, "detail": "no result.json next to the work dir"}
    res = load_json(result_path, {}) or {}
    victim = None
    for pref in ("plan", "outline", "page_plan", "evidence", "page_evidence_map"):
        if pref in want and pref in res:
            victim = pref
            break
    if victim is None:
        for key in reversed(want):
            if key not in ("case_id", "status") and key in res:
                victim = key
                break
    if victim is None:
        return {"ok": False, "detail": "no removable contract key present in result.json"}
    del res[victim]
    dump_json(result_path, res)
    # keep a copy inside the work dir so verify_deck (no --out) reads the tampered file
    if result_path.parent != work:
        dump_json(work / "result.json", res)
    return {"ok": True, "detail": "removed output_contract key %r from %s"
                                  % (victim, Path(result_path).name)}


def mut_text_overflow(work: Path, case: Path):
    """V10: text that is genuinely clipped.

    Stuffing long text into a placeholder is not enough any more: V10 (correctly) treats
    ``normAutofit`` as "PowerPoint shrinks the text", so a shape carrying shrink-to-fit is
    not an overflow. The injection therefore also removes shrink-to-fit from the slides, so
    the long text is really clipped -- which is the defect class V10 exists to catch.
    """
    from pptx import Presentation
    from pptx.enum.text import MSO_AUTO_SIZE
    deck = work / "deck.pptx"
    prs = Presentation(str(deck))
    long_text = ("Overflow probe sentence repeated to exceed any placeholder extent. " * 60)
    target_slide, target_shape = None, None
    for idx, slide in enumerate(prs.slides, 1):
        if idx < 2:
            continue
        best = None
        for shape in slide.shapes:
            try:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    if best is None or len(shape.text_frame.text) > len(best.text_frame.text):
                        best = shape
            except Exception:
                continue
        if best is not None:
            target_slide, target_shape = idx, best
            break
    if target_shape is None:
        return {"ok": False, "detail": "no text-bearing slide >= 2 found"}
    tf = target_shape.text_frame
    first = None
    for para in tf.paragraphs:
        if para.runs:
            first = para
            break
    if first is None:
        return {"ok": False, "detail": "target text frame has no run to overwrite"}
    first.runs[0].text = long_text
    for extra in list(first.runs[1:]):
        extra.text = ""
    try:
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE      # -> noAutofit: text is clipped, not shrunk
        autofit_note = "auto_size set to NONE"
    except Exception as exc:  # noqa: BLE001
        autofit_note = "auto_size not settable via python-pptx (%s)" % type(exc).__name__
    prs.save(str(deck))

    with zipfile.ZipFile(str(deck)) as zin:
        slide_parts = [n for n in zin.namelist()
                       if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    stripped = 0
    for part in slide_parts:
        stripped += rewrite_zip_part(
            deck, part,
            lambda xml: xml.replace("<a:normAutofit/>", "<a:noAutofit/>")
                           .replace("<a:normAutofit", "<a:noAutofit"))
    if not stripped and "NONE" not in autofit_note:
        return {"ok": False, "detail": "could not remove shrink-to-fit (%s)" % autofit_note}
    return {"ok": True,
            "detail": "slide %d text frame set to %d chars, %s, shrink-to-fit stripped from "
                      "%d slide part(s)" % (target_slide, len(long_text), autofit_note, stripped)}


def _noop(work, case):
    return {"ok": False, "detail": "unused"}


def mut_duplicate_slide(work: Path, case: Path):
    """V16: two slides carrying the same title/bullets/notes."""
    from pptx import Presentation
    deck = work / "deck.pptx"
    plan_path = work / "plan.json"
    plan = load_json(plan_path, {}) or {}
    slides = plan.get("slides") or []
    if len(slides) < 3:
        return {"ok": False, "detail": "plan has fewer than 3 slides"}
    src, dst = slides[1], slides[2]
    copied = []
    for key in ("title", "claim", "bullets", "notes"):
        if src.get(key):
            dst[key] = src[key]
            copied.append(key)
    if not copied:
        return {"ok": False, "detail": "source slide carries nothing to duplicate"}
    dump_json(plan_path, plan)

    prs = Presentation(str(deck))
    s2, s3 = _slide_by_index(prs, 2), _slide_by_index(prs, 3)
    changed = 0
    for a, b in ((s2, s3),):
        if a is None or b is None:
            continue
        t_a, body_a = _title_and_body(a)
        t_b, body_b = _title_and_body(b)
        if t_a is not None and t_b is not None and t_a.text_frame.text.strip():
            t_b.text_frame.text = t_a.text_frame.text
            changed += 1
        if body_a is not None and body_b is not None and body_a.text_frame.text.strip():
            body_b.text_frame.text = body_a.text_frame.text
            changed += 1
    if not changed:
        return {"ok": False, "detail": "could not mirror slide 2's text onto slide 3 in the deck"}
    prs.save(str(deck))
    return {"ok": True, "detail": "slide 3 now repeats slide 2 (%s); deck mirrored (%d shape(s))"
                                  % (",".join(copied), changed)}


def mut_constraint_leak(work: Path, case: Path):
    """V17: the brief's own metadata pasted onto a slide as content."""
    from pptx import Presentation
    deck = work / "deck.pptx"
    case_obj = load_json(case, {}) or {}
    meta = case_obj.get("constraints") or {}
    lines = []
    if meta.get("audience"):
        lines.append("Audience: %s" % meta["audience"])
    if meta.get("duration_min") is not None:
        lines.append("Duration: %s minutes" % meta["duration_min"])
    if meta.get("slide_count") is not None:
        lines.append("Slide count: %s" % meta["slide_count"])
    if meta.get("template_path"):
        lines.append("Template: %s" % os.path.basename(str(meta["template_path"])))
    if not lines:
        return {"ok": False, "detail": "case.json declares no constraint metadata to leak"}
    for line in lines[:3]:
        if line.split(":", 1)[1].strip().isdigit():
            lines.remove(line)
    lines = lines[:3]
    if not lines:
        return {"ok": False, "detail": "no usable metadata line"}

    plan_path = work / "plan.json"
    plan = load_json(plan_path, {}) or {}
    slides = plan.get("slides") or []
    if len(slides) < 2:
        return {"ok": False, "detail": "plan has fewer than 2 slides"}
    slides[1]["bullets"] = list(slides[1].get("bullets") or []) + lines
    dump_json(plan_path, plan)

    prs = Presentation(str(deck))
    slide = _slide_by_index(prs, 2)
    if slide is None:
        return {"ok": False, "detail": "deck has no slide 2"}
    _, body = _title_and_body(slide)
    if body is None:
        return {"ok": False, "detail": "deck slide 2 has no body placeholder"}
    for line in lines:
        body.text_frame.add_paragraph().text = line
    prs.save(str(deck))
    return {"ok": True, "detail": "leaked brief metadata onto slide 2: %s" % " | ".join(lines)}


def mut_build_report_foreign(work: Path, case: Path):
    """V18: build_report.json describes a different deck than the one inspected."""
    br_path = work / "build_report.json"
    build = load_json(br_path, {}) or {}
    original = build.get("pptx")
    if not original:
        return {"ok": False, "detail": "build_report.json declares no pptx path to falsify"}
    build["pptx"] = str(work.parent / "elsewhere" / "deck.pptx")
    dump_json(br_path, build)
    return {"ok": True,
            "detail": "build_report.pptx %r -> %r" % (original, build["pptx"])}


def mut_gutted_material(work: Path, case: Path):
    """V19: the source was read but nothing came out of it (empty text layer / bad decode)."""
    mat_path = work / "material.json"
    material = load_json(mat_path, {}) or {}
    if not material:
        return {"ok": False, "detail": "material.json missing"}
    sources = material.get("sources") or []
    if not sources:
        return {"ok": False, "detail": "material.json has no sources"}
    before_chars = sum(len(str(p.get("text") or ""))
                       for s in sources for p in (s.get("pages") or []))
    for source in sources:
        for page in source.get("pages") or []:
            page["text"] = ""
        source["sections"] = []
    emptied = {k: len(material.get(k) or [])
               for k in ("figures", "tables", "numbers", "claims")}
    for key in emptied:
        material[key] = []
    dump_json(mat_path, material)
    return {"ok": True,
            "detail": "material.json gutted: %d text chars -> 0, emptied %s"
                      % (before_chars, emptied)}


def _slide_by_index(prs, index):
    for i, slide in enumerate(prs.slides, 1):
        if i == index:
            return slide
    return None


def _title_and_body(slide):
    """Return (title_placeholder, body_placeholder) for a slide, best effort."""
    title = None
    try:
        title = slide.shapes.title
    except Exception:
        title = None
    body = None
    for shape in slide.placeholders:
        try:
            if shape is title:
                continue
            if shape.has_text_frame and (shape.placeholder_format.idx or 0) >= 1:
                body = shape
                break
        except Exception:
            continue
    if body is None:
        for shape in slide.shapes:
            try:
                if shape is not title and shape.has_text_frame and shape.text_frame.text.strip():
                    body = shape
                    break
            except Exception:
                continue
    return title, body


NEGATIVES = [
    ("N01", "fabricated number not in the source", ["V04"], mut_fabricated_number),
    ("N02", "missing / too-short speaker notes", ["V03"], mut_short_notes),
    ("N03", "wrong slide count vs the case constraint", ["V01"], mut_slide_count),
    ("N04", "modified theme (fontScheme)", ["V08"], mut_theme_modified),
    ("N05", "modified master/layout", ["V07"], mut_master_modified),
    ("N06", "placeholder / invented figure", ["V06"], mut_placeholder_figure),
    ("N07", "missing output_contract key in result.json", ["V13"], mut_missing_contract_key),
    ("N08", "text overflow on a slide", ["V10"], mut_text_overflow),
    ("N09", "duplicate slide content", ["V16"], mut_duplicate_slide),
    ("N10", "brief metadata leaked into a bullet", ["V17"], mut_constraint_leak),
    ("N11", "build_report describes a foreign deck", ["V18"], mut_build_report_foreign),
    ("N12", "gutted source / empty text layer", ["V19"], mut_gutted_material),
]


# --------------------------------------------------------------------------- #
# positive checks
# --------------------------------------------------------------------------- #
def check_positive(case_path: Path, work: Path, out: Path, verify: Path, timeout: int,
                   enforce: bool = True):
    """Run run.py end to end; return a result dict.

    ``enforce=False`` records the case as EXERCISED: its red checks are reported (as
    ``red_notes``) but they do not fail this harness, which is what lets a teammate fix
    deck-core defects on the A/B cases without the self-test becoming a moving target. At
    least one case must be known-good and fully green.
    """
    res = {"case": case_path.parent.name, "case_json": str(case_path),
           "status": "FAIL", "problems": [], "skips": [], "steps": []}

    rmtree(work)
    rmtree(out)
    work.mkdir(parents=True, exist_ok=True)
    out.mkdir(parents=True, exist_ok=True)

    rc, so, se = run_cli([SCRIPTS / "run.py", "--case", case_path, "--work", work,
                          "--out", out], cwd=SUBMISSION, timeout=timeout)
    res["steps"].append({"name": "run.py", "exit": rc,
                         "stdout_tail": so.strip().splitlines()[-1][:200] if so.strip() else "",
                         "stderr_tail": se.strip()[-200:]})
    if rc != 0:
        res["problems"].append("run.py exited %d (expected 0); stderr=%s"
                               % (rc, se.strip()[-300:]))
        return res

    for name in WORK_FILES:
        if not (work / name).is_file():
            res["problems"].append("missing work-stage artifact %s" % name)
    for name in OUT_FILES:
        if not (out / name).is_file() and not (work / name).is_file():
            res["problems"].append("missing output artifact %s" % name)
    if not (work / "deck.pptx").is_file() and not (out / "deck.pptx").is_file():
        res["problems"].append("no deck.pptx anywhere")
    # version records (CONTRACT section 4.1 / gate V15)
    version_files = [n for n in ("versions.json", "change_records.json", "diff-v1.md",
                                 "rollback.md")
                     if (out / n).is_file() or (work / n).is_file()]
    res["version_records"] = version_files
    if len(version_files) < 4:
        res["problems"].append("version records incomplete in the output dir: %s"
                               % (version_files or "none"))

    # --- render honesty (renderer-dependent -> SKIP, never a false PASS) ------------- #
    render = load_json(work / "render_report.json", {}) or {}
    render_status = (render.get("status") or "unknown")
    renderer = render.get("renderer") or {}
    res["render_status"] = render_status
    if render_status == "ok":
        pdf = ((render.get("pdf") or {}).get("path")) or str(work / "render" / "deck.pdf")
        pdf_ok = bool(pdf) and Path(str(pdf)).is_file()
        pngs = [p for p in (render.get("pngs") or []) if Path(str(p.get("path"))).is_file()]
        out_pngs = sorted((out).glob("slide-*.png")) if out.is_dir() else []
        if not pdf_ok:
            res["problems"].append("render_report.status=ok but no PDF exists")
        if not pngs and not out_pngs:
            res["problems"].append("render_report.status=ok but no PNG preview exists")
        res["pdf"] = str(pdf) if pdf_ok else None
        res["pngs"] = len(pngs) or len(out_pngs)
    elif render_status in ("skipped", "failed"):
        honest = (renderer.get("available") is False
                  and (render.get("error") is not None or render.get("uncertainty")))
        if not honest:
            res["problems"].append(
                "render_report.status=%s but the report does not declare an unavailable "
                "renderer plus an error/uncertainty entry (not an honest degradation)"
                % render_status)
        why = "no renderer available (%s); PDF/PNG assertions skipped: %s" % (
            renderer.get("name") or "none",
            (render.get("uncertainty") or ["no uncertainty recorded"])[-1])
        res["skips"].append({"assertion": "deck.pdf + slide PNG previews", "reason": why})
    else:
        res["problems"].append("render_report status is %r (expected ok|skipped|failed)"
                               % render_status)

    # --- verify_report verdict ------------------------------------------------------- #
    report = load_json(out / "verify_report.json", None)
    if not isinstance(report, dict) or not report:
        report = load_json(work / "verify_report.json", {}) or {}
    if not report:
        res["problems"].append("verify_report.json unreadable/absent")
    else:
        counts = report.get("counts") or {}
        res["verdict"] = report.get("verdict")
        res["counts"] = counts
        res["failed_checks"] = [c.get("id") for c in (report.get("checks") or [])
                                if c.get("status") in ("FAIL", "UNVERIFIED")]
        if report.get("verdict") != "PASS":
            res["problems"].append("verify_report.verdict=%r (expected PASS); not-pass ids=%s"
                                   % (report.get("verdict"), res["failed_checks"]))
        if counts.get("FAIL"):
            res["problems"].append("%s FAIL check(s) in the good run" % counts["FAIL"])
        if counts.get("UNVERIFIED"):
            res["problems"].append("%s UNVERIFIED check(s) in the good run" % counts["UNVERIFIED"])
        ids = {c.get("id") for c in (report.get("checks") or [])}
        missing = sorted({v for v in (("V%02d" % i) for i in range(1, 16))} - ids)
        if missing:
            res["problems"].append("verify_report omits mandatory gate id(s): %s" % missing)
        res["gate_ids"] = sorted(ids)

    # --- independent re-run of the gate --------------------------------------------- #
    rc2, so2, se2 = run_cli([verify, "--case", case_path, "--work", work, "--out", out],
                            cwd=SUBMISSION, timeout=timeout)
    res["steps"].append({"name": "verify_deck.py (independent re-run)", "exit": rc2,
                         "stdout_tail": so2.strip().splitlines()[-1][:200] if so2.strip() else "",
                         "stderr_tail": se2.strip()[-200:]})
    if rc2 != 0:
        res["problems"].append("independent verify_deck.py re-run exited %d (expected 0)"
                               % rc2)

    # --- result.json contract coverage ---------------------------------------------- #
    case_obj = load_json(case_path, {}) or {}
    want = set((case_obj.get("output_contract") or {}).keys())
    res_json = load_json(out / "result.json", {}) or {}
    miss = sorted(want - set(res_json.keys()))
    if miss:
        res["problems"].append("result.json misses output_contract key(s): %s" % miss)
    for extra in ("case_id", "status", "sources", "uncertainty"):
        if extra not in res_json:
            res["problems"].append("result.json lacks the always-required key %r" % extra)
    if res_json.get("case_id") != case_obj.get("case_id"):
        res["problems"].append("result.json case_id=%r != case.json case_id=%r"
                               % (res_json.get("case_id"), case_obj.get("case_id")))
    if res_json.get("status") not in ("ok", "partial", "failed"):
        res["problems"].append("result.json status=%r not in ok|partial|failed"
                               % res_json.get("status"))
    res["contract_keys"] = sorted(want)

    if res["problems"] and not enforce:
        res["red_notes"] = list(res["problems"])
        res["problems"] = []
        res["status"] = "EXERCISED"
    else:
        res["status"] = "PASS" if not res["problems"] else "FAIL"
    return res


# --------------------------------------------------------------------------- #
# negative checks
# --------------------------------------------------------------------------- #
def run_negative(nid, name, expected_ids, mutate, base_out, chosen_work, case_path,
                 verify, tmp_root, timeout, good_failing_ids=(), emitted_ids=()):
    """Stage a copy of a real run's output dir, inject one defect, require detection.

    ``base_out`` is the flat output directory of a real ``run.py`` run -- it already holds
    deck.pptx, plan.json, material.json, build_report.json, render_report.json,
    evidence_map.json, result.json, verify_report.json and the version records. Each negative
    gets its own copy, mutates *that* copy, and is verified with ``--work <copy> --out <copy>``
    so V13 (which prefers ``--out``) and V15 (which scans ``--out``) see the staged files.
    """
    res = {"id": nid, "name": name, "expected_check_ids": expected_ids,
           "status": "NOT_DETECTED", "observed": [], "detail": "",
           "already_failing_before_injection": sorted(set(expected_ids) & set(good_failing_ids))}
    absent = [cid for cid in expected_ids if emitted_ids and cid not in emitted_ids]
    if absent:
        res["status"] = "GATE_ABSENT"
        res["detail"] = ("verify_deck.py does not emit %s yet, so this defect cannot be "
                         "detected; recorded as SKIP-with-reason, NOT as PASS" % absent)
        return res
    if not base_out.is_dir():
        res["status"] = "HARNESS_ERROR"
        res["detail"] = "the staged baseline output dir %s is missing" % base_out
        return res
    work = tmp_root / ("neg-%s" % nid)
    try:
        rmtree(work)
        shutil.copytree(str(base_out), str(work))
    except OSError as exc:
        res["status"] = "HARNESS_ERROR"
        res["detail"] = "could not stage the baseline output dir: %s" % exc
        return res

    # relocate absolute paths so the injected defect is the only difference
    replaced = rewrite_absolute_paths(work, str(base_out), str(work))
    replaced += rewrite_absolute_paths(work, str(chosen_work), str(work))
    res["paths_rewritten_in_files"] = replaced

    if not (work / "deck.pptx").is_file():
        res["status"] = "HARNESS_ERROR"
        res["detail"] = ("the staged copy has no deck.pptx, so no mutation can target the "
                         "deck; base_out=%s content=%s"
                         % (base_out, sorted(p.name for p in work.iterdir())[:12]))
        return res

    try:
        applied = mutate(work, case_path)
    except Exception as exc:  # noqa: BLE001
        res["status"] = "MUTATION_FAILED"
        res["detail"] = "injection raised %s: %s" % (type(exc).__name__, exc)
        return res
    res["detail"] = applied.get("detail", "")
    if not applied.get("ok"):
        res["status"] = "MUTATION_FAILED"
        return res

    report_path = work / "verify_report.json"
    if report_path.is_file():
        report_path.unlink()

    # Control: identical staging and path rewrites, but NO defect injected. Comparing against
    # this control is what proves the injection -- not the relocation of the directory and not
    # a verifier change that landed between the two runs -- caused the failure.
    control = tmp_root / ("control-%s" % nid)
    try:
        rmtree(control)
        shutil.copytree(str(base_out), str(control))
        rewrite_absolute_paths(control, str(base_out), str(control))
        rewrite_absolute_paths(control, str(chosen_work), str(control))
    except OSError as exc:
        res["status"] = "HARNESS_ERROR"
        res["detail"] = "could not stage the control copy: %s" % exc
        return res
    c_rc, c_so, c_se = run_cli([verify, "--case", case_path, "--work", control,
                                "--out", control], cwd=SUBMISSION, timeout=timeout)
    control_json = load_json(control / "verify_report.json", {}) or {}
    control_failing = [c.get("id") for c in (control_json.get("checks") or [])
                       if c.get("status") in ("FAIL", "UNVERIFIED")]
    res["control_exit"] = c_rc
    res["control_failing_ids"] = control_failing
    rmtree(control)

    rc, so, se = run_cli([verify, "--case", case_path, "--work", work, "--out", work],
                         cwd=SUBMISSION, timeout=timeout)
    res["verify_exit"] = rc
    report = load_json(report_path, {}) or {}
    observed = [c.get("id") for c in (report.get("checks") or [])
                if c.get("status") in ("FAIL", "UNVERIFIED")]
    res["observed"] = observed
    res["verdict"] = report.get("verdict")

    if rc == 0:
        res["status"] = "NOT_DETECTED"
        res["detail"] += " | verify_deck.py exited 0 (defect not detected)"
        return res
    hit = [cid for cid in expected_ids if cid in observed]
    if not hit:
        res["status"] = "WRONG_CHECK"
        res["detail"] += (" | verify_deck.py exited %d but the expected gate id(s) %s did not "
                          "fail; observed=%s" % (rc, expected_ids, observed))
        return res
    # A gate id that was ALREADY failing before the injection proves nothing: the negative
    # must show a transition PASS -> FAIL, otherwise the injection is not the cause.
    pre = sorted(set(hit) & (set(good_failing_ids) | set(control_failing)))
    if pre:
        res["status"] = "PREEXISTING"
        res["detail"] += (" | %s already failed without the injection (relocated control: %s), "
                          "so the injection is not proven to be the cause"
                          % (pre, control_failing or "-"))
        return res
    res["status"] = "DETECTED"
    res["detected_by"] = hit
    return res


# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="End-to-end self-test with defect injection")
    ap.add_argument("--fixtures", default=None,
                    help="fixtures dir (default: <work>/fixtures, sibling of submission/)")
    ap.add_argument("--tmp", default=None,
                    help="scratch root (default: <fixtures>/_selftest_tmp)")
    ap.add_argument("--case", action="append", default=None,
                    help="only test this case directory name (repeatable); negatives run on "
                         "the first one")
    ap.add_argument("--json", default=None, help="write the machine-readable summary here")
    ap.add_argument("--timeout", type=int, default=900, help="seconds per CLI call")
    ap.add_argument("--keep", action="store_true", help="keep the scratch tree")
    ap.add_argument("--known-good", action="append", default=None,
                    help="case dir name asserted fully green end-to-end (repeatable; default: "
                         "<fixtures>/known_good.txt, else every case)")
    ap.add_argument("--skip-negatives", action="store_true")
    ap.add_argument("--only", action="append", default=None,
                    help="run only these negative ids (repeatable), e.g. --only N08")
    args = ap.parse_args(argv)

    fixtures = Path(args.fixtures).resolve() if args.fixtures else SUBMISSION.parent / "fixtures"
    if args.tmp:
        tmp_root = Path(args.tmp).resolve()
    else:
        # One scratch dir PER RUN. Two concurrent self-tests used to share
        # ``fixtures/_selftest_tmp``, and whichever finished first rmtree'd the other's live
        # tree mid-flight -- which presented as "11 negatives MUTATION_FAILED".
        scratch_parent = fixtures / "_selftest_tmp"
        scratch_parent.mkdir(parents=True, exist_ok=True)
        now = time.time()
        for old in scratch_parent.glob("run-*"):
            try:                                     # prune only long-abandoned runs
                if now - old.stat().st_mtime > 1800:
                    rmtree(old)
            except OSError:
                pass
        tmp_root = scratch_parent / ("run-%d-%s" % (os.getpid(),
                                                    time.strftime("%Y%m%d-%H%M%S")))

    known_good = set(args.known_good or [])
    if not known_good:
        marker = fixtures / "known_good.txt"
        if marker.is_file():
            known_good = {line.strip() for line in
                          marker.read_text(encoding="utf-8").splitlines()
                          if line.strip() and not line.strip().startswith("#")}

    summary = {"schema": "selftest/v1", "fixtures": str(fixtures),
               "submission": str(SUBMISSION), "positives": [], "negatives": [],
               "skips": [], "problems": [], "status": "FAIL"}

    missing = [name for name, path in PIPELINE.items() if not path.is_file()]
    if missing:
        summary["status"] = "BLOCKED"
        summary["problems"].append(
            "pipeline script(s) missing, so nothing can be tested honestly: %s" % missing)
        summary["missing_pipeline_scripts"] = missing
        log("BLOCKED: missing pipeline script(s): %s" % ", ".join(missing))
        log("SELFTEST_JSON " + json.dumps(summary, ensure_ascii=False))
        if args.json:
            dump_json(Path(args.json), summary)
        return 2

    cases = []
    if args.case:
        for name in args.case:
            cand = fixtures / name / "case.json"
            if cand.is_file():
                cases.append(cand)
            else:
                summary["problems"].append("case.json not found: %s" % cand)
    else:
        for cand in sorted(fixtures.glob("*/case.json")):
            cases.append(cand)
    if not cases:
        summary["status"] = "BLOCKED"
        summary["problems"].append("no fixture case.json found under %s" % fixtures)
        log("BLOCKED: no fixtures found under %s" % fixtures)
        log("SELFTEST_JSON " + json.dumps(summary, ensure_ascii=False))
        if args.json:
            dump_json(Path(args.json), summary)
        return 2

    if tmp_root.exists():
        rmtree(tmp_root)
    tmp_root.mkdir(parents=True, exist_ok=True)
    summary["scratch_root"] = str(tmp_root)

    verify = PIPELINE["verify_deck.py"]
    emitted_ids = emitted_gate_ids(verify)
    summary["emitted_gate_ids"] = sorted(emitted_ids)
    summary["known_good_cases"] = sorted(known_good) or "all"
    log("verify_deck.py emits: %s" % ", ".join(sorted(emitted_ids)))
    good = {}
    for case_path in cases:
        tag = case_path.parent.name
        work = tmp_root / tag / "work"
        out = tmp_root / tag / "out"
        is_known_good = (not known_good) or (tag in known_good)
        log("== %s: %s ==" % ("known-good" if is_known_good else "exercised", tag))
        res = check_positive(case_path, work, out, verify, args.timeout,
                             enforce=is_known_good)
        res["known_good"] = is_known_good
        res["red_check_ids"] = res.get("failed_checks") or []
        summary["positives"].append(res)
        summary["skips"].extend({"case": tag, **s} for s in res.get("skips", []))
        usable = (work / "deck.pptx").is_file() and (work / "plan.json").is_file()
        if res["status"] == "PASS":
            log("   PASS  verdict=%s counts=%s render=%s" % (
                res.get("verdict"), res.get("counts"), res.get("render_status")))
        else:
            log("   FAIL  %s" % "; ".join(res["problems"])[:600])
        if not is_known_good and res.get("red_check_ids"):
            log("   (exercised only: red check(s) %s are deck-core's to fix; this case is not "
                "asserted green by this harness)" % ",".join(res["red_check_ids"]))
        if usable or res["status"] == "PASS":
            score = (0 if res["status"] == "PASS" else 1)
            good[tag] = (case_path, work, out, res["status"] == "PASS",
                         res.get("failed_checks") or [], score, is_known_good)

    if not args.skip_negatives and good:
        # Prefer a known-good case whose baseline is clean; fall back to any usable workspace.
        ordered = sorted(good.items(), key=lambda kv: (not kv[1][6], kv[1][5],
                                                       not kv[1][3], kv[1][4]))
        tag, (case_path, work, out, baseline_clean, good_failing, _score, is_kg) = ordered[0]
        summary["negative_baseline_clean"] = baseline_clean
        summary["negative_baseline_failing_ids"] = good_failing
        summary["negative_baseline_case"] = tag

        # Stage ONE stable baseline for every negative: copy the chosen run's flat output
        # directory once, so a negative can never be affected by a later stage rewriting the
        # live output dir, and so the staging itself is verified before the loop starts.
        base_out = tmp_root / "base" / "out"
        rmtree(tmp_root / "base")
        staged = None
        try:
            if out.is_dir() and (out / "deck.pptx").is_file():
                base_out.parent.mkdir(parents=True, exist_ok=True)
                shutil.copytree(str(out), str(base_out))
                staged = "out"
            elif work.is_dir() and (work / "deck.pptx").is_file():
                base_out.parent.mkdir(parents=True, exist_ok=True)
                shutil.copytree(str(work), str(base_out))
                # bring the output-dir-only deliverables (version records, run.log) along
                if out.is_dir():
                    for src in sorted(out.rglob("*")):
                        if src.is_file():
                            dst = base_out / src.relative_to(out)
                            dst.parent.mkdir(parents=True, exist_ok=True)
                            shutil.copy2(str(src), str(dst))
                staged = "work"
        except OSError as exc:
            staged = None
            summary["problems"].append("staging the negative baseline failed: %s" % exc)
        summary["negative_staged_from"] = staged
        summary["negative_staged_files"] = (
            len(list(base_out.rglob("*"))) if base_out.is_dir() else 0)
        if staged is None:
            summary["problems"].append(
                "could not stage a baseline with a deck.pptx from %s or %s, so no negative "
                "could be exercised" % (out, work))
            log("HARNESS ERROR: no staged baseline deck; negatives not exercised")
        else:
            log("staged baseline for negatives from the %s dir: %s (%d entries)"
                % (staged, base_out, summary["negative_staged_files"]))

        if not baseline_clean:
            log("NOTE: the un-injected run for %s is not clean (%s already failing), so a "
                "negative only counts as DETECTED when its gate id was PASS before the "
                "injection (PASS -> FAIL transition required)."
                % (tag, ",".join(good_failing) or "-"))
        log("== negatives on %s (expected to be detected) ==" % tag)
        wanted = set(args.only or [])
        selected = [n for n in NEGATIVES if not wanted or n[0] in wanted]
        if wanted:
            unknown = sorted(wanted - {n[0] for n in NEGATIVES})
            if unknown:
                summary["problems"].append("--only names unknown negative id(s): %s" % unknown)
        for nid, name, expected, mutate in selected:
            try:
                res = run_negative(nid, name, expected, mutate, base_out, work, case_path,
                                   verify, tmp_root, args.timeout,
                                   good_failing_ids=good_failing,
                                   emitted_ids=emitted_ids)
            except Exception as exc:  # noqa: BLE001
                res = {"id": nid, "name": name, "expected_check_ids": expected,
                       "status": "HARNESS_ERROR", "observed": [],
                       "detail": "%s: %s" % (type(exc).__name__, exc)}
            summary["negatives"].append(res)
            log("   %-13s %-46s expected=%s observed=%s | %s"
                % (res["status"], name[:46], ",".join(expected),
                   ",".join(res["observed"][:6]) or "-", res["detail"][:160]))

    pos_ok = all(p["status"] != "FAIL" for p in summary["positives"]) and any(
        p["status"] == "PASS" for p in summary["positives"])
    blocking_statuses = {"NOT_DETECTED", "WRONG_CHECK", "MUTATION_FAILED", "PREEXISTING",
                         "HARNESS_ERROR"}
    neg_ok = (args.skip_negatives or
              all(n["status"] not in blocking_statuses for n in summary["negatives"]))
    summary["status"] = "PASS" if (pos_ok and neg_ok) else "FAIL"
    summary["counts"] = {
        "positives_total": len(summary["positives"]),
        "positives_passed": sum(1 for p in summary["positives"] if p["status"] == "PASS"),
        "positives_exercised": sum(1 for p in summary["positives"]
                                   if p["status"] == "EXERCISED"),
        "negatives_total": len(summary["negatives"]),
        "negatives_detected": sum(1 for n in summary["negatives"]
                                  if n["status"] == "DETECTED"),
        "negatives_not_detected": sum(1 for n in summary["negatives"]
                                      if n["status"] == "NOT_DETECTED"),
        "negatives_mutation_failed": sum(1 for n in summary["negatives"]
                                         if n["status"] == "MUTATION_FAILED"),
        "negatives_wrong_check": sum(1 for n in summary["negatives"]
                                     if n["status"] == "WRONG_CHECK"),
        "negatives_preexisting": sum(1 for n in summary["negatives"]
                                     if n["status"] == "PREEXISTING"),
        "negatives_gate_absent_skipped": sum(1 for n in summary["negatives"]
                                             if n["status"] == "GATE_ABSENT"),
        "negatives_harness_error": sum(1 for n in summary["negatives"]
                                       if n["status"] == "HARNESS_ERROR"),
        "skipped_render_assertions": len(summary["skips"]),
    }
    for p in summary["positives"]:
        for problem in p["problems"]:
            summary["problems"].append("%s: %s" % (p["case"], problem))
    for n in summary["negatives"]:
        if n["status"] in blocking_statuses:
            summary["problems"].append("%s %s: %s" % (n["id"], n["status"], n["detail"]))
        elif n["status"] == "GATE_ABSENT":
            summary["skips"].append({"case": "-", "assertion": "%s %s" % (n["id"], n["name"]),
                                    "reason": n["detail"]})

    if not args.keep:
        rmtree(tmp_root)

    log("-" * 92)
    log("counts: %s" % json.dumps(summary["counts"], ensure_ascii=False))
    harness_ok = (summary["counts"]["negatives_mutation_failed"] == 0
                  and summary["counts"]["negatives_harness_error"] == 0)
    summary["harness_ok"] = harness_ok
    summary["success_requires"] = (
        "every negative DETECTED by its expected gate id, with negatives_mutation_failed == 0 "
        "and negatives_harness_error == 0; positives green for every known-good case")
    if not harness_ok:
        log("HARNESS ERROR: %d negative(s) MUTATION_FAILED, %d HARNESS_ERROR. These are "
            "defects in THIS harness, not successes -- read 'negatives_not_detected: %d' only "
            "together with the mutation/harness counts, never on its own."
            % (summary["counts"]["negatives_mutation_failed"],
               summary["counts"]["negatives_harness_error"],
               summary["counts"]["negatives_not_detected"]))
    for skip in summary["skips"]:
        log("SKIP  %s: %s -- %s" % (skip["case"], skip["assertion"], skip["reason"][:160]))
    for problem in summary["problems"]:
        log("PROBLEM: %s" % problem[:400])
    log("SELFTEST: %s" % summary["status"])
    log("SELFTEST_JSON " + json.dumps(summary, ensure_ascii=False))
    if args.json:
        dump_json(Path(args.json), summary)
    return 0 if summary["status"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
