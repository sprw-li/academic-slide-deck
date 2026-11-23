#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Measured-rework harness for the ARMI:2611.000074v1 academic-PPT Skill.

WHAT THIS IS
------------
The rubric (``task/task.md:43``) asks whether the Skill *reduces rework*: pages that need
re-ordering, export re-check rounds, numbers that fail traceability, template violations.
This script measures those four quantities on **our own synthetic fixture** and prints them
next to a **deterministic baseline proxy**.

WHAT THIS IS NOT
----------------
* It is **not** an official score. The platform scores the Skill on its own hidden cases
  against its own whiteboard; nothing here is that number. ``official_score`` is always
  ``null`` in the report.
* It is **not** a real no-Skill baseline run. Producing one would require running a base
  agent (no Skill loaded) on the same case under the same budget, and this host has no such
  second harness and sees no platform whiteboard. So the comparison column is a
  *deterministic proxy* built from the source material alone (one slide per source section,
  source order, text copied verbatim). It is labelled
  ``baseline_kind = "deterministic-proxy-not-a-model-run"`` everywhere it appears, and the
  numbers must not be quoted as "the base Agent's result".

METRIC DEFINITIONS (all computed from artifacts, never estimated by prose)
-------------------------------------------------------------------------
``pages_needing_reorder``
    Minimum number of slides a human would have to move: for each slide, take the source
    page of its first evidence locator; the ``L`` longest increasing subsequence of that
    page sequence is already in source order, so ``len(slides) - L`` slides must be moved.
``export_recheck_rounds``
    Number of verification passes actually recorded for this run (``run.log`` stage events
    plus the independent re-run), i.e. how many times the deck had to go back through the
    export/verify loop before the gate passed. A no-Skill baseline cannot be measured for
    this metric from our artifacts and is reported as ``null`` -- not as a guess.
``numbers_failing_traceability``
    Numeric tokens on the deck's slides with no verbatim match in ``material.numbers``.
    Recomputed here independently of ``verify_deck.py`` and cross-checked against gate V04.
``template_violations``
    Non-PASS results among gate ids V07 (masters/layouts), V08 (theme fonts), V09 (footer).

Usage::

    python scripts/gain_check.py [--fixtures DIR] [--case caseA] [--work W] [--out O]
                                 [--baseline-only] [--json OUT] [--tmp DIR]

Exit codes: 0 = metrics computed; 1 = a metric could not be computed / V04 disagreed with
our recomputation; 2 = the pipeline could not be run (e.g. ``run.py`` missing) and no
existing ``--work`` was supplied.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")

HERE = Path(__file__).resolve()
SCRIPTS = HERE.parent
SUBMISSION = SCRIPTS.parent

DISCLAIMER = (
    "Measured on our own synthetic fixture (SYNTHETIC / NOT AN OFFICIAL CASE). "
    "This is not an official score: the platform scores on its own hidden cases against "
    "its own whiteboard. The comparison column is a deterministic proxy, NOT a measured "
    "base-Agent (no-Skill) run."
)

NUMBER_RX = re.compile(r"(?<![A-Za-z0-9_.])-?\d+(?:[.,]\d+)?(?:[eE][-+]?\d+)?")


# --------------------------------------------------------------------------- #
def log(msg):
    print(msg, flush=True)


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


def rmtree(path):
    shutil.rmtree(str(path), ignore_errors=True)


def run_cli(args, cwd, timeout):
    try:
        proc = subprocess.run([sys.executable] + [str(a) for a in args],
                              capture_output=True, text=True, encoding="utf-8",
                              errors="replace", cwd=str(cwd), timeout=timeout)
        return proc.returncode, proc.stdout or "", proc.stderr or ""
    except subprocess.TimeoutExpired:
        return 124, "", "timeout after %ss" % timeout


def lis_length(seq):
    """Longest *non-decreasing* subsequence length (patience sorting, O(n log n)).

    Non-decreasing on purpose: two slides whose evidence sits on the same source page
    impose no ordering constraint between them, so equal page numbers must not be counted
    as "out of order".
    """
    import bisect
    tails = []
    for value in seq:
        i = bisect.bisect_right(tails, value)
        if i == len(tails):
            tails.append(value)
        else:
            tails[i] = value
    return len(tails)


def slide_texts(deck: Path):
    """List of (index, text) for every slide's shapes, or None if unreadable."""
    try:
        from pptx import Presentation
    except Exception:
        return None
    try:
        prs = Presentation(str(deck))
    except Exception:
        return None
    out = []
    for idx, slide in enumerate(prs.slides, 1):
        chunks = []
        for shape in slide.shapes:
            try:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    chunks.append(shape.text_frame.text)
            except Exception:
                continue
        out.append((idx, "\n".join(chunks)))
    return out


# --------------------------------------------------------------------------- #
# metrics
# --------------------------------------------------------------------------- #
def metric_pages_needing_reorder(plan: dict, material: dict):
    """Minimum slides to move, from the source-page order of each slide's evidence."""
    pages = {}
    for entry in material.get("figures") or []:
        pages[entry.get("id")] = entry.get("page")
    for entry in material.get("tables") or []:
        pages[entry.get("id")] = entry.get("page")
    for entry in material.get("sections") or []:
        pages[entry.get("id")] = entry.get("page")

    seq, slides_without = [], 0
    for slide in plan.get("slides") or []:
        chosen = None
        for ev in (slide.get("evidence") or []):
            loc = ev.get("locator") or {}
            if isinstance(loc.get("page"), int):
                chosen = loc["page"]
                break
            ref = ev.get("ref")
            if ref in pages and isinstance(pages[ref], int):
                chosen = pages[ref]
                break
        if chosen is None and slide.get("figure_ref") in pages:
            chosen = pages[slide["figure_ref"]]
        if chosen is None and slide.get("table_ref") in pages:
            chosen = pages[slide["table_ref"]]
        if chosen is None:
            slides_without += 1
            continue
        seq.append(chosen)

    total = len(plan.get("slides") or [])
    if not seq:
        return {"value": None, "slides": total, "slides_with_locator_page": 0,
                "note": "no slide carried a source-page locator; metric not computable"}
    ordered = lis_length(seq)
    return {"value": len(seq) - ordered, "slides": total,
            "slides_with_locator_page": len(seq),
            "slides_without_locator_page": slides_without,
            "source_page_sequence": seq,
            "longest_increasing_subsequence": ordered,
            "definition": "len(slides with locator) - LIS(page sequence); the LIS is the "
                          "already-correct part, so the remainder must be moved"}


def metric_export_recheck_rounds(work: Path, out: Path, log_text: str):
    stage_runs = len(re.findall(r'"stage"\s*:\s*"verify_deck"', log_text))
    stage_runs += len(re.findall(r"stage_start\s+.*verify", log_text))
    verify_files = [p for p in [work / "verify_report.json", out / "verify_report.json"]
                    if p.is_file()]
    return {
        "value": max(stage_runs, 1 if verify_files else 0),
        "verify_reports_present": len(verify_files),
        "verification_stage_events_in_log": stage_runs,
        "baseline_value": None,
        "definition": "recorded verification passes for this run (log stage events); a "
                      "no-Skill baseline cannot be measured from our artifacts, so the "
                      "baseline is null rather than guessed",
    }


def attested_numbers(material: dict):
    allowed = set()
    for entry in material.get("numbers") or []:
        raw = str(entry.get("raw", "")).strip()
        if raw:
            allowed.add(raw)
            allowed.add(raw.replace(",", "."))
        value = entry.get("value")
        if isinstance(value, (int, float)):
            allowed.add(("%g" % value))
    return allowed


def metric_untraceable_numbers(deck: Path, material: dict):
    texts = slide_texts(deck)
    if texts is None:
        return {"value": None, "note": "deck unreadable or python-pptx unavailable"}
    allowed = attested_numbers(material)
    offenders, checked = [], 0
    for idx, text in texts:
        for match in NUMBER_RX.finditer(text):
            token = match.group(0)
            checked += 1
            if token in allowed or token.replace(",", ".") in allowed:
                continue
            if token.rstrip("0").rstrip(".") in allowed:
                continue
            offenders.append({"slide": idx, "token": token})
    return {"value": len(offenders), "numeric_tokens_on_slides": checked,
            "attested_numbers_in_material": len(allowed),
            "offenders": offenders[:10],
            "definition": "numeric tokens with no verbatim match in material.numbers "
                          "(recomputed independently of verify_deck.py; numbers in figure "
                          "captions may legitimately be labels)"}


def metric_template_violations(report: dict):
    bad = [c.get("id") for c in (report.get("checks") or [])
           if c.get("id") in ("V07", "V08", "V09") and c.get("status") != "PASS"]
    statuses = {c.get("id"): c.get("status") for c in (report.get("checks") or [])
                if c.get("id") in ("V07", "V08", "V09")}
    return {"value": len(bad), "violating_gate_ids": bad, "statuses": statuses,
            "definition": "non-PASS results among V07 (masters/layouts), V08 (theme fonts), "
                          "V09 (footer/slide-number placeholders)"}


def extra_metrics(plan: dict, material: dict):
    slides = plan.get("slides") or []
    return {
        "slides_total": len(slides),
        "slides_over_bullet_cap": sum(1 for s in slides if len(s.get("bullets") or []) > 5),
        "slides_below_notes_min": sum(1 for s in slides
                                      if len((s.get("notes") or "").strip()) < 40),
        "slides_without_evidence": sum(1 for s in slides if not (s.get("evidence") or [])),
        "figures_in_material": len(material.get("figures") or []),
        "tables_in_material": len(material.get("tables") or []),
        "claims_in_material": len(material.get("claims") or []),
        "material_warnings": len(material.get("warnings") or []),
    }


# --------------------------------------------------------------------------- #
# deterministic baseline proxy
# --------------------------------------------------------------------------- #
def _sections_by_source(material: dict):
    """Sections grouped by source, preferring the PDF source when one exists.

    Both the with-Skill column and the baseline column then rest on the *same* page
    numbering, which is the only way the re-order metric can be compared honestly.
    """
    sources = [s for s in (material.get("sources") or []) if isinstance(s, dict)]
    pdf = [s for s in sources if (s.get("kind") or "") == "pdf" and (s.get("sections") or [])]
    chosen = pdf[0] if pdf else None
    if chosen is None:
        with_sections = [s for s in sources if (s.get("sections") or [])]
        chosen = with_sections[0] if with_sections else (sources[0] if sources else {})
    out = []
    for sec in (chosen.get("sections") or []):
        if (sec.get("title") or "").strip():
            out.append({"id": sec.get("id"), "title": sec.get("title"),
                        "page": sec.get("page"), "source": chosen.get("id")})
    return chosen, out


def baseline_proxy(material: dict, case: dict):
    """Two deterministic no-Skill constructions built from ``material.json`` alone.

    These are *proxies*, not the output of a base Agent:
    ``section_order``     one slide per source section, source order, text verbatim;
    ``paragraph_dump``    one slide per source claim/paragraph -- the documented naive
                          failure mode ("paste the abstract or paragraph into the slide"),
                          which is what makes the slide count explode.
    """
    chosen, sections = _sections_by_source(material)
    claims = material.get("claims") or []

    sec_slides = [{"title": s["title"], "page": s["page"]} for s in sections]
    sec_pages = [s["page"] for s in sections if isinstance(s.get("page"), int)]
    para_pages = []
    for claim in claims:
        loc = claim.get("locator") or {}
        if isinstance(loc.get("page"), int):
            para_pages.append(loc["page"])

    constraint = None
    if isinstance(case.get("constraints"), dict):
        constraint = case["constraints"].get("slide_count")
    if not isinstance(constraint, int):
        constraint = None

    def pack(name, construction, slides, pages, notes):
        return {
            "construction": construction,
            "slides": slides,
            "pages_needing_reorder": (len(pages) - lis_length(pages)) if pages else None,
            "slides_over_case_constraint": (max(0, slides - constraint)
                                            if constraint is not None else None),
            "slides_with_notes": notes,
            "slides_with_evidence": slides,
            "numbers_failing_traceability": 0,
            "template_violations": None,
            "export_recheck_rounds": None,
        }

    return {
        "baseline_kind": "deterministic-proxy-not-a-model-run",
        "page_basis_source": chosen.get("id"),
        "page_basis_note": "both columns use the page numbers of the same source (%s), so the "
                           "re-order metric is comparable" % chosen.get("id"),
        "case_slide_constraint": constraint,
        "section_order": pack("section_order",
                              "one slide per source section, source order, text verbatim",
                              len(sec_slides), sec_pages, notes=0),
        "paragraph_dump": pack("paragraph_dump",
                               "one slide per source claim/paragraph, copied verbatim "
                               "(the documented naive failure mode)",
                               len(claims), para_pages, notes=0),
        "caveat": "a real no-Skill baseline run is NOT available on this host (no second "
                  "agent harness, no platform whiteboard); these columns are deterministic "
                  "proxies and must not be quoted as the base Agent's measured result",
    }


# --------------------------------------------------------------------------- #
# naive baseline arm: paste-first-sentences-into-placeholders
# --------------------------------------------------------------------------- #
NAIVE_ARM_KIND = "deterministic-naive-plan-through-the-real-builder"

NAIVE_LIMITATION = (
    "The naive arm is NOT a base-Agent (no-Skill) model run: this host has no second agent "
    "harness and no platform whiteboard. It is a documented deterministic construction "
    "(one slide per source section, first sentences pasted verbatim into the placeholders, "
    "no speaker notes, no evidence locators, slide count not curated) pushed through the SAME "
    "scripts/build_deck.py and the SAME scripts/verify_deck.py as the real arm, so both "
    "columns are measured by one metric implementation. Quote it as 'our naive construction', "
    "never as the base Agent's measured result and never as an official score."
)


def _first_sentences(text, limit=3, max_chars=220):
    import re as _re
    text = " ".join(str(text or "").split())
    if not text:
        return []
    parts = _re.split(r"(?<=[.!?])\s+", text)
    out = []
    for part in parts[:limit]:
        out.append(part[:max_chars])
    return [p for p in out if p]


def build_naive_plan(case: dict, material: dict) -> dict:
    """A schema-conforming plan that pastes source prose into the content placeholders."""
    _chosen, sections = _sections_by_source(material)
    if not sections:
        sections = []
    constraints = case.get("constraints") if isinstance(case.get("constraints"), dict) else {}
    # first sentences per source page, used when the source exposes no headings
    pages = []
    for source in material.get("sources") or []:
        for page in source.get("pages") or []:
            if (page.get("text") or "").strip():
                pages.append({"id": "%s-p%s" % (source.get("id"), page.get("n")),
                              "title": "Page %s" % page.get("n"), "page": page.get("n"),
                              "text": page.get("text")})
    units = []
    for sec in sections:
        text = ""
        for source in material.get("sources") or []:
            for page in source.get("pages") or []:
                if isinstance(sec.get("page"), int) and page.get("n") == sec.get("page"):
                    text += " " + (page.get("text") or "")
        units.append((sec.get("title") or "Section", text or sec.get("title") or ""))
    if not units:
        units = [(p["title"], p["text"]) for p in pages]
    if not units:
        units = [("Content", "the source yielded no text")]

    slides = []
    for i, (title, text) in enumerate(units, 1):
        bullets = _first_sentences(text, limit=3) or [title]
        slides.append({
            "index": i, "role": "method", "title": title,
            "claim": bullets[0] if bullets else title,
            "bullets": bullets,
            "figure_ref": None, "table_ref": None,
            "evidence": [], "numbers_used": [],
            "notes": "",                       # naive: no speaker notes at all
            "layout_hint": "title-content",
        })
    return {
        "schema": "plan/v1",
        "case_id": case.get("case_id"),
        "constraints": dict(constraints),
        "narrative": {"arc": "naive: source order, one slide per section",
                      "main_thread": [s["title"] for s in slides]},
        "slides": slides,
        "excluded": [],
        "uncertainty": [
            "naive arm: slide count not curated against the case constraint",
            "naive arm: no speaker notes, no evidence locators, no figure/table binding",
        ],
    }


def _gate_failures(counts):
    """Number of non-PASS gate results (FAIL + UNVERIFIED) in a verify_report counts dict."""
    if not isinstance(counts, dict):
        return None
    return int(counts.get("FAIL") or 0) + int(counts.get("UNVERIFIED") or 0)


def manual_fix_count(metrics: dict) -> int:
    """Our own composite rework count (documented, not an official metric).

    Sum of: slides over the bullet cap + slides below the notes minimum + slides over the
    case constraint + pages needing a re-order + numbers failing traceability + template
    violations. Every term is measured, none is estimated.
    """
    extras = metrics.get("extras") or {}
    total = 0
    total += int(extras.get("slides_over_bullet_cap") or 0)
    total += int(extras.get("slides_below_notes_min") or 0)
    total += int(metrics.get("_slides_over_constraint") or 0)
    for key in ("pages_needing_reorder", "numbers_failing_traceability",
                "template_violations"):
        value = (metrics.get(key) or {}).get("value")
        total += int(value or 0)
    return total


def run_naive_arm(case: dict, case_path: Path, material: dict, fixture_root: Path,
                  tmp_root: Path, timeout: int):
    """Push the naive plan through the real builder and gate; measure the same metrics."""
    arm = {"arm": "naive (paste-first-sentences)", "arm_kind": NAIVE_ARM_KIND}
    work = tmp_root / "naive" / "work"
    out = tmp_root / "naive" / "out"
    rmtree(tmp_root / "naive")
    work.mkdir(parents=True, exist_ok=True)
    out.mkdir(parents=True, exist_ok=True)

    plan = build_naive_plan(case, material)
    dump_json(work / "plan.json", plan)
    dump_json(work / "material.json", material)
    for name, value in (case.get("synth") or {}).items():
        if name == "template":
            pass

    # the naive arm uses the case's own template, like a naive agent would
    import shutil as _shutil
    tmpl = None
    raw = ((case.get("constraints") or {}).get("template_path")
           or case.get("template") or "")
    candidates = [case_path.parent / raw if raw else None,
                  case_path.parent / "materials" / "template-organizer.pptx",
                  fixture_root / "template-organizer.pptx"]
    for cand in candidates:
        if cand and Path(cand).is_file():
            tmpl = Path(cand)
            break
    if tmpl is None:
        arm["status"] = "SKIPPED"
        arm["reason"] = "no template locatable for the naive arm"
        return arm, plan
    arm["template"] = tmpl.name

    steps = []
    rc, so, se = run_cli([SCRIPTS / "build_deck.py", "--case", case_path, "--work", work],
                         cwd=SUBMISSION, timeout=timeout)
    steps.append({"stage": "build_deck.py", "exit": rc,
                  "stderr_tail": se.strip()[-200:]})
    deck_ok = (work / "deck.pptx").is_file()
    if rc != 0 or not deck_ok:
        arm["steps"] = steps
        arm["status"] = "BUILD_FAILED"
        arm["reason"] = "the naive plan did not build (%s)" % (se.strip()[-160:] or rc)
        return arm, plan
    rc2, so2, se2 = run_cli([SCRIPTS / "verify_deck.py", "--case", case_path,
                             "--work", work, "--out", out],
                            cwd=SUBMISSION, timeout=timeout)
    steps.append({"stage": "verify_deck.py", "exit": rc2, "stderr_tail": se2.strip()[-200:]})
    arm["steps"] = steps

    verify = load_json(out / "verify_report.json", {}) or {}
    metrics = {
        "pages_needing_reorder": metric_pages_needing_reorder(plan, material),
        "export_recheck_rounds": metric_export_recheck_rounds(work, out, ""),
        "numbers_failing_traceability": metric_untraceable_numbers(work / "deck.pptx", material),
        "template_violations": metric_template_violations(verify),
        "extras": extra_metrics(plan, material),
        "verify_verdict": verify.get("verdict"),
        "verify_counts": verify.get("counts"),
    }
    constraint = (case.get("constraints") or {}).get("slide_count")
    metrics["_slides_over_constraint"] = (max(0, len(plan["slides"]) - constraint)
                                         if isinstance(constraint, int) else None)
    arm["metrics"] = metrics
    arm["manual_fix_count"] = manual_fix_count(metrics)
    arm["verify_failing_ids"] = [c.get("id") for c in (verify.get("checks") or [])
                                 if c.get("status") in ("FAIL", "UNVERIFIED")]
    arm["status"] = "MEASURED"
    return arm, plan


# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Measured-rework harness (not an official score)")
    ap.add_argument("--fixtures", default=None)
    ap.add_argument("--case", default="caseA", help="fixture case directory name")
    ap.add_argument("--work", default=None, help="use this existing work dir instead of running")
    ap.add_argument("--out", default=None, help="output dir of the run")
    ap.add_argument("--tmp", default=None, help="scratch root for a fresh run")
    ap.add_argument("--baseline-only", action="store_true",
                    help="compute only the deterministic baseline proxy from material.json")
    ap.add_argument("--timeout", type=int, default=900)
    ap.add_argument("--naive-arm", dest="naive_arm", action="store_true", default=True,
                    help="also measure the naive paste-first-sentences arm through the real "
                         "builder and gate (default: on)")
    ap.add_argument("--no-naive-arm", dest="naive_arm", action="store_false",
                    help="skip the naive arm")
    ap.add_argument("--json", default=None)
    args = ap.parse_args(argv)

    fixtures = Path(args.fixtures).resolve() if args.fixtures else SUBMISSION.parent / "fixtures"
    case_path = fixtures / args.case / "case.json"
    report = {"schema": "gain_check/v1", "case": args.case,
              "case_json": str(case_path), "disclaimer": DISCLAIMER,
              "official_score": None, "measured_on": "own synthetic fixture",
              "with_skill": {}, "baseline": {}, "cross_checks": {}, "problems": [],
              "status": "FAIL"}

    if not case_path.is_file():
        report["problems"].append("case.json not found: %s" % case_path)
        log("error: case.json not found: %s" % case_path)
        log("GAIN_CHECK_JSON " + json.dumps(report, ensure_ascii=False))
        if args.json:
            dump_json(Path(args.json), report)
        return 2

    case = load_json(case_path, {}) or {}
    work = Path(args.work).resolve() if args.work else None
    out = Path(args.out).resolve() if args.out else None
    tmp_root = Path(args.tmp).resolve() if args.tmp else fixtures / "_gain_tmp"

    if args.baseline_only and work is None:
        report["problems"].append("--baseline-only needs --work (a dir with material.json)")
        log("error: --baseline-only needs --work")
        log("GAIN_CHECK_JSON " + json.dumps(report, ensure_ascii=False))
        return 2

    if work is None:
        run_py = SCRIPTS / "run.py"
        if not run_py.is_file():
            report["status"] = "BLOCKED"
            report["problems"].append("run.py is missing and no --work was supplied, so no "
                                      "metrics can be measured")
            log("BLOCKED: scripts/run.py missing and no --work given")
            log("GAIN_CHECK_JSON " + json.dumps(report, ensure_ascii=False))
            if args.json:
                dump_json(Path(args.json), report)
            return 2
        tmp = tmp_root
        work = tmp / "work"
        out = tmp / "out"
        rmtree(tmp)
        work.mkdir(parents=True, exist_ok=True)
        out.mkdir(parents=True, exist_ok=True)
        rc, so, se = run_cli([run_py, "--case", case_path, "--work", work, "--out", out],
                             cwd=SUBMISSION, timeout=args.timeout)
        report["run"] = {"exit": rc,
                         "stdout_tail": so.strip().splitlines()[-1][:200] if so.strip() else "",
                         "stderr_tail": se.strip()[-200:]}
        if rc != 0:
            report["problems"].append("run.py exited %d; metrics cannot be measured" % rc)
            log("FAIL: run.py exited %d" % rc)
            log("GAIN_CHECK_JSON " + json.dumps(report, ensure_ascii=False))
            if args.json:
                dump_json(Path(args.json), report)
            return 1
    if out is None:
        out = work

    material = load_json(work / "material.json", {}) or {}
    plan = load_json(work / "plan.json", {}) or {}
    verify = load_json(work / "verify_report.json", {}) or {}
    log_text = ""
    for cand in (out / "run.log", work / "run.log"):
        if cand.is_file():
            log_text = cand.read_text(encoding="utf-8", errors="replace")
            break

    if not material:
        report["problems"].append("material.json missing in %s" % work)

    report["baseline"] = baseline_proxy(material, case)
    if args.baseline_only:
        report["status"] = "PASS" if not report["problems"] else "FAIL"
        print_gain_report(report)
        log("GAIN_CHECK_JSON " + json.dumps(report, ensure_ascii=False))
        if args.json:
            dump_json(Path(args.json), report)
        return 0 if report["status"] == "PASS" else 1

    with_skill = {
        "pages_needing_reorder": metric_pages_needing_reorder(plan, material),
        "export_recheck_rounds": metric_export_recheck_rounds(work, out, log_text),
        "numbers_failing_traceability": metric_untraceable_numbers(work / "deck.pptx", material),
        "template_violations": metric_template_violations(verify),
        "extras": extra_metrics(plan, material),
        "verify_verdict": verify.get("verdict"),
        "verify_counts": verify.get("counts"),
    }
    constraint = (case.get("constraints") or {}).get("slide_count")
    if not isinstance(constraint, int):
        constraint = None
    with_skill["_slides_over_constraint"] = (
        max(0, len(plan.get("slides") or []) - constraint)
        if constraint is not None else None)
    with_skill["manual_fix_count"] = manual_fix_count(with_skill)
    report["with_skill"] = with_skill

    # ---- naive baseline arm, through the real builder + real gate -------------------- #
    naive_arm, naive_plan = {"status": "SKIPPED", "reason": "--no-naive-arm"}, None
    if args.naive_arm:
        naive_arm, naive_plan = run_naive_arm(case, case_path, material, fixtures,
                                              tmp_root, args.timeout)
    report["baseline_arms"] = {
        "naive_paste_arm": naive_arm,
        "deterministic_proxies": report.get("baseline"),
        "baseline_note": NAIVE_LIMITATION,
    }
    if naive_arm.get("status") == "MEASURED":
        report["baseline"] = {
            "baseline_kind": NAIVE_ARM_KIND,
            "construction": naive_arm["arm"],
            "section_order": naive_arm["metrics"],
            "paragraph_dump": (report.get("baseline") or {}).get("paragraph_dump") or {},
            "page_basis_source": (report.get("baseline") or {}).get("page_basis_source"),
            "case_slide_constraint": constraint,
            "caveat": NAIVE_LIMITATION,
        }
    else:
        report["problems"].append("naive baseline arm not measured: %s"
                                  % naive_arm.get("reason", naive_arm.get("status")))

    # cross-check our independent V04 recomputation against the gate's own verdict
    v04 = {c.get("id"): c.get("status") for c in (verify.get("checks") or [])}.get("V04")
    untraceable = with_skill["numbers_failing_traceability"].get("value")
    report["cross_checks"]["V04_status"] = v04
    report["cross_checks"]["our_untraceable_count"] = untraceable
    if v04 == "PASS" and untraceable not in (None, 0):
        report["problems"].append(
            "gate V04 says PASS but our independent recomputation found %s untraceable "
            "numeric token(s): the two disagree" % untraceable)
    if v04 in ("FAIL", "UNVERIFIED") and untraceable == 0 and untraceable is not None:
        report["problems"].append(
            "gate V04 is %s but our recomputation found no untraceable token "
            "(check the gate's material source)" % v04)

    for key in ("pages_needing_reorder", "numbers_failing_traceability",
                "template_violations"):
        if with_skill[key].get("value") is None:
            report["problems"].append("metric %s could not be computed" % key)

    # ---- with/without comparison, and an explicit statement of what it cannot show ---- #
    base = report["baseline"]
    base_sec, base_par = base["section_order"], base["paragraph_dump"]
    constraint = base.get("case_slide_constraint")
    plan_slides = len((plan.get("slides") or []))
    naive = (naive_arm.get("metrics") or {}) if naive_arm.get("status") == "MEASURED" else {}

    def _v(arm_metrics, key):
        return (arm_metrics.get(key) or {}).get("value")

    comparisons = [
        {"metric": "pages_needing_reorder",
         "with_skill": _v(with_skill, "pages_needing_reorder"),
         "naive_arm": _v(naive, "pages_needing_reorder"),
         "proxy_section": base_sec.get("pages_needing_reorder")},
        {"metric": "numbers_failing_traceability",
         "with_skill": _v(with_skill, "numbers_failing_traceability"),
         "naive_arm": _v(naive, "numbers_failing_traceability"),
         "proxy_section": base_sec.get("numbers_failing_traceability")},
        {"metric": "template_violations",
         "with_skill": _v(with_skill, "template_violations"),
         "naive_arm": _v(naive, "template_violations"),
         "proxy_section": None,
         "note": "the deterministic proxy is never built through the template, so only the "
                 "naive arm and the real arm have a template column"},
        {"metric": "slides_over_case_constraint",
         "with_skill": (max(0, plan_slides - constraint) if isinstance(constraint, int)
                        else None),
         "naive_arm": naive.get("_slides_over_constraint"),
         "proxy_section": base_sec.get("slides_over_case_constraint")},
        {"metric": "slides_below_notes_min",
         "with_skill": (with_skill.get("extras") or {}).get("slides_below_notes_min"),
         "naive_arm": (naive.get("extras") or {}).get("slides_below_notes_min"),
         "proxy_section": (base_sec.get("extras") or {}).get("slides_below_notes_min")},
        {"metric": "manual_fix_count (composite)",
         "with_skill": with_skill.get("manual_fix_count"),
         "naive_arm": naive_arm.get("manual_fix_count"),
         "proxy_section": None,
         "note": "our own composite: over-cap slides + under-notes slides + over-constraint "
                 "slides + reordered pages + untraceable numbers + template violations"},
        {"metric": "gate_failures (FAIL+UNVERIFIED checks)",
         "with_skill": _gate_failures(with_skill.get("verify_counts")),
         "naive_arm": _gate_failures(naive.get("verify_counts")),
         "proxy_section": None,
         "note": "the clearest discriminator: how many of the Skill's own gates the arm's deck "
                 "fails. Naive failing ids: %s"
                 % (",".join(naive_arm.get("verify_failing_ids") or []) or "-")},
    ]
    for row in comparisons:
        cand = [v for k, v in row.items()
                if k in ("naive_arm", "proxy_section") and isinstance(v, (int, float))]
        row["discriminating"] = bool(cand) and row["with_skill"] is not None and any(
            row["with_skill"] != v for v in cand)
    report["comparison"] = comparisons
    report["naive_arm_status"] = naive_arm.get("status")
    report["limitations"] = [
        NAIVE_LIMITATION,
        "Noise floor: one run on one synthetic fixture. A difference of one or two units is "
        "not evidence of a real improvement.",
    ]
    if naive_arm.get("status") != "MEASURED":
        report["limitations"].append(
            "the naive arm could not be measured (%s), so the A/B table carries only the real "
            "arm" % naive_arm.get("reason", naive_arm.get("status")))
    for row in comparisons:
        if not row["discriminating"]:
            report["limitations"].append(
                "metric %r does not discriminate on this fixture (with Skill %s vs %s); it is "
                "reported for completeness, not as evidence"
                % (row["metric"], row["with_skill"],
                   {k: v for k, v in row.items() if k in ("naive_arm", "proxy_section")}))
    report["official_score_claim"] = None

    report["status"] = "PASS" if not report["problems"] else "FAIL"
    print_gain_report(report)
    log("GAIN_CHECK_JSON " + json.dumps(report, ensure_ascii=False))
    if args.json:
        dump_json(Path(args.json), report)
    return 0 if report["status"] == "PASS" else 1


def print_gain_report(report: dict) -> None:
    log("=" * 92)
    log("measured rework — case %s, on OUR synthetic fixture (NOT an official score)"
        % report["case"])
    log(report["disclaimer"])
    log("-" * 92)
    ws, base = report.get("with_skill") or {}, report.get("baseline") or {}
    if report.get("comparison"):
        log("%-34s %12s %12s %12s" % ("metric", "with Skill", "naive arm", "proxy:sec"))
        for row in report["comparison"]:
            log("%-34s %12s %12s %12s%s" % (
                row["metric"],
                "n/a" if row["with_skill"] is None else row["with_skill"],
                "n/a" if row.get("naive_arm") is None else row["naive_arm"],
                "n/a" if row.get("proxy_section") is None else row["proxy_section"],
                "" if row["discriminating"] else "   (does not discriminate here)"))
        log("-" * 92)
        log("lower is better for every metric above except slide counts, which are exact "
            "constraints (0 = compliant).")
        log("naive arm = %s [%s]" % ((base.get("construction") or "n/a"),
                                     report.get("naive_arm_status")))
        log("proxy:sec = %s" % ((report.get("baseline_arms", {}).get("deterministic_proxies")
                                 or {}).get("section_order") or {}).get("construction"))
    else:
        log("with-Skill metrics not measured in this mode (--baseline-only)")
    if base:
        log("-" * 92)
        log("baseline_kind = %s" % base.get("baseline_kind"))
    if ws.get("extras"):
        log("with-Skill extras: %s" % json.dumps(ws["extras"], ensure_ascii=False))
    if ws.get("verify_verdict"):
        log("verify verdict  : %s %s" % (ws.get("verify_verdict"), ws.get("verify_counts")))
    log("cross-checks    : %s" % json.dumps(report.get("cross_checks", {}), ensure_ascii=False))
    for limitation in report.get("limitations", []):
        log("LIMITATION: %s" % limitation)
    for problem in report.get("problems", []):
        log("PROBLEM: %s" % problem)
    log("OVERALL(gain_check): %s   (official_score=%s)"
        % (report["status"], report.get("official_score")))
    log("=" * 92)


if __name__ == "__main__":
    sys.exit(main())
