#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DEV PROBE (not shipped, not part of the submission).

Exercises every defect-injection helper in ``scripts/selftest.py`` against a throwaway
deck built from the synthetic organizer template.  It proves that each injection can
actually be applied (selftest.py treats a no-op injection as MUTATION_FAILED, so this probe
catches an injection that silently does nothing long before a full ``run.py`` exists).

It does *not* prove detection -- that is selftest.py's job with the real pipeline.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")

HERE = Path(__file__).resolve().parent
SCRIPTS = HERE.parents[0] / "submission" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import selftest as st  # noqa: E402

from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402
from PIL import Image  # noqa: E402


def build_deck(path: Path, case_json: Path) -> None:
    prs = Presentation(str(HERE / "template-organizer.pptx"))
    layouts = list(prs.slide_layouts)
    for i in range(3):
        layout = layouts[min(i + 1, len(layouts) - 1)]
        slide = prs.slides.add_slide(layout)
        placeholders = slide.placeholders
        if placeholders:
            try:
                placeholders[0].text = "Probe slide %d title" % (i + 1)
            except Exception:
                pass
        body = None
        for ph in placeholders:
            if ph.placeholder_format.idx == 1:
                body = ph
                break
        if body is not None:
            tf = body.text_frame
            tf.text = "Probe bullet one"
            tf.add_paragraph().text = "Probe bullet two"
        slide.notes_slide.notes_text_frame.text = (
            "Probe speaker notes for slide %d, comfortably longer than forty characters." % (i + 1))
    img = path.parent / "_probe.png"
    Image.new("RGB", (300, 200), (30, 110, 180)).save(str(img))
    prs.slides[1].shapes.add_picture(str(img), Inches(1), Inches(4), width=Inches(2))
    prs.save(str(path))

    # minimal artifacts the injections touch
    (path.parent / "plan.json").write_text(json.dumps({
        "schema": "plan/v1", "case_id": "PROBE",
        "constraints": {"slide_count": 8},
        "slides": [{"index": i, "title": "t%d" % i, "claim": "c%d" % i,
                    "bullets": ["b"], "notes": "n" * 60,
                    "evidence": [{"type": "figure", "ref": "fig1",
                                  "locator": {"page": 2}}]}
                   for i in range(1, 5)],
    }, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")
    (path.parent / "build_report.json").write_text(json.dumps({
        "schema": "build_report/v1", "pptx": str(path),
        "slides": [{"index": i, "images": []} for i in range(1, 4)],
    }, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")
    (path.parent / "result.json").write_text(json.dumps({
        "case_id": "PROBE", "status": "ok", "plan": "plan.json",
        "deck_pptx": "deck.pptx", "uncertainty": [], "change_records": [],
    }, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")
    (path.parent / "material.json").write_text(json.dumps({
        "schema": "material/v1", "case_id": "PROBE",
        "sources": [{"id": "src1", "kind": "pdf", "path": "paper.pdf",
                     "pages": [{"n": 1, "text": "probe source text " * 20}],
                     "sections": [{"id": "S1", "title": "Abstract", "page": 1}]}],
        "figures": [{"id": "fig1", "label": "Figure 1", "page": 1, "caption": "probe figure"}],
        "tables": [{"id": "tab1", "label": "Table 1", "page": 1, "cells": [["a", "b"]]}],
        "numbers": [{"value": 1.5, "raw": "1.5", "unit": "%", "context": "probe"}],
        "claims": [{"text": "probe claim"}],
        "warnings": [],
    }, ensure_ascii=False, indent=2), encoding="utf-8", newline="\n")


def main() -> int:
    tmp = HERE / "_probe_mut"
    shutil.rmtree(str(tmp), ignore_errors=True)
    tmp.mkdir(parents=True, exist_ok=True)

    case_json = HERE / "caseA" / "case.json"
    failures = []
    print("%-4s %-46s %-8s %s" % ("id", "injection", "applied", "detail"))
    for nid, name, expected, mutate in st.NEGATIVES:
        work = tmp / nid
        shutil.rmtree(str(work), ignore_errors=True)
        work.mkdir(parents=True, exist_ok=True)
        build_deck(work / "deck.pptx", case_json)
        before = st.deck_slide_count(work / "deck.pptx")
        try:
            res = mutate(work, case_json)
        except Exception as exc:  # noqa: BLE001
            res = {"ok": False, "detail": "%s: %s" % (type(exc).__name__, exc)}
        after = st.deck_slide_count(work / "deck.pptx")
        status = "ok" if res.get("ok") else "NO-OP"
        if not res.get("ok"):
            failures.append("%s %s: %s" % (nid, name, res.get("detail")))
        print("%-4s %-46s %-8s slides %d->%d | %s"
              % (nid, name[:46], status, before, after, res.get("detail", "")[:90]))

    # the path-rewrite helper must actually rewrite relocated paths
    (tmp / "nt").mkdir(exist_ok=True)
    p = tmp / "nt" / "x.json"
    p.write_text(json.dumps({"pptx": str(tmp / "nope" / "deck.pptx")}), encoding="utf-8")
    hits = st.rewrite_absolute_paths(tmp / "nt", str(tmp / "nope"), str(tmp / "nt"))
    if hits != 1:
        failures.append("rewrite_absolute_paths: expected 1 rewrite, got %d" % hits)
    print("\nrewrite_absolute_paths rewrote %d file(s)" % hits)

    if failures:
        print("PROBE RESULT: FAIL (%d)" % len(failures))
        for f in failures:
            print("  - %s" % f)
        return 1
    print("PROBE RESULT: OK (all %d injections applied)" % len(st.NEGATIVES))
    return 0


if __name__ == "__main__":
    sys.exit(main())
