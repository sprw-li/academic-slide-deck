#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DEV SCRIPT (lives in work/fixtures, never shipped).

Takes a real pipeline run (``--work`` / ``--out``) and copies a *sanitized* development
example into ``work/submission/examples/dev-run-caseA/``:

* the input ``case.json`` (platform shape, with its ``output_contract``),
* real ``result.json`` / ``plan.json`` / ``evidence_map.json`` / ``verify_report.json``,
* key excerpts of ``build_report.json`` and ``render_report.json``,
* a truncated-but-real ``run.log`` timing excerpt,
* a small PNG contact sheet of the rendered slides.

Sanitizing is mandatory because the shipped tree must not contain author-machine absolute
paths (``work/verify/check.py`` S04): every occurrence of the run's own directories, and any
remaining ``<drive>:\\...`` absolute path, is replaced with a portable token. Content is
otherwise untouched, and every JSON is re-parsed after sanitizing so a broken escape cannot
ship silently.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")

HERE = Path(__file__).resolve().parent
SUBMISSION = HERE.parent / "submission"
DEST = SUBMISSION / "examples" / "dev-run-caseA"
FEATURED = ["result.json", "plan.json", "evidence_map.json", "verify_report.json"]

ABS_RX = re.compile(r"[A-Za-z]:\\\\?[^\"\s,)\]]*")


def sanitize(text: str, pairs) -> str:
    for old, new in pairs:
        if not old:
            continue
        text = text.replace(old, new)
        text = text.replace(old.replace("\\", "/"), new)
        text = text.replace(old.replace("\\", "\\\\"), new)
    return ABS_RX.sub("<abs-path-not-shipped>", text)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--case", required=True)
    ap.add_argument("--work", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--case-name", default=None, help="fixture dir name for the note")
    ap.add_argument("--thumb-width", type=int, default=320)
    args = ap.parse_args(argv)

    case = Path(args.case).resolve()
    work = Path(args.work).resolve()
    out = Path(args.out).resolve()

    pairs = [
        (str(out), "output/"),
        (str(work), "work/"),
        (str(case.parent), "input/"),
        (str(case), "input/case.json"),
        (str(HERE), "work/fixtures/"),
        (str(SUBMISSION), "submission/"),
    ]

    if DEST.exists():
        shutil.rmtree(DEST)
    DEST.mkdir(parents=True, exist_ok=True)

    written = {}

    # ---- input ---------------------------------------------------------------------- #
    shutil.copy2(str(case), str(DEST / "case.json"))
    written["case.json"] = (DEST / "case.json").stat().st_size

    # ---- real artifacts ------------------------------------------------------------- #
    for name in FEATURED:
        src = out / name if (out / name).is_file() else (work / name)
        if not src.is_file():
            print("missing artifact: %s" % name, file=sys.stderr)
            return 1
        obj = json.loads(src.read_text(encoding="utf-8"))
        text = sanitize(json.dumps(obj, ensure_ascii=False, indent=2) + "\n", pairs)
        json.loads(text)  # must still parse
        (DEST / name).write_text(text, encoding="utf-8", newline="\n")
        written[name] = (DEST / name).stat().st_size

    # ---- excerpts of the two reports ------------------------------------------------ #
    build = json.loads((work / "build_report.json").read_text(encoding="utf-8"))
    build_excerpt = {
        "_excerpt": "key fields of the real build_report.json (slide image/provenance detail "
                    "trimmed); paths sanitized for shipping",
        "schema": build.get("schema"),
        "template": build.get("template"),
        "invariants": build.get("invariants"),
        "fonts_used": build.get("fonts_used"),
        "warnings": build.get("warnings"),
        "slides": [{k: s.get(k) for k in ("index", "layout", "placeholders_filled",
                                          "overflow_risk", "notes_chars", "images")}
                   for s in (build.get("slides") or [])],
    }
    text = sanitize(json.dumps(build_excerpt, ensure_ascii=False, indent=2) + "\n", pairs)
    json.loads(text)
    (DEST / "build_report.excerpt.json").write_text(text, encoding="utf-8", newline="\n")
    written["build_report.excerpt.json"] = (DEST / "build_report.excerpt.json").stat().st_size

    render = json.loads((work / "render_report.json").read_text(encoding="utf-8"))
    render_excerpt = {
        "_excerpt": "key fields of the real render_report.json (per-PNG hashes/paths trimmed)",
        "schema": render.get("schema"),
        "renderer": render.get("renderer"),
        "status": render.get("status"),
        "error": render.get("error"),
        "uncertainty": render.get("uncertainty"),
        "pdf": {k: v for k, v in (render.get("pdf") or {}).items() if k != "path"},
        "pngs": [{"slide": p.get("slide"), "width": p.get("width"), "height": p.get("height"),
                  "bytes": p.get("bytes")} for p in (render.get("pngs") or [])],
    }
    if render_excerpt["pdf"]:
        render_excerpt["pdf"]["path"] = "output/render/deck.pdf"
    text = sanitize(json.dumps(render_excerpt, ensure_ascii=False, indent=2) + "\n", pairs)
    json.loads(text)
    (DEST / "render_report.excerpt.json").write_text(text, encoding="utf-8", newline="\n")
    written["render_report.excerpt.json"] = (
        DEST / "render_report.excerpt.json").stat().st_size

    # ---- run.log timing excerpt (real stage events, truncated) ---------------------- #
    log_src = out / "run.log" if (out / "run.log").is_file() else work / "run.log"
    lines = []
    if log_src.is_file():
        for line in log_src.read_text(encoding="utf-8", errors="replace").splitlines():
            try:
                event = json.loads(line)
            except Exception:
                continue
            if event.get("event") in ("stage_start", "stage_end", "run_end"):
                event.pop("ts", None)
                lines.append(json.dumps(event, ensure_ascii=False))
    body = ("# real stage/timing events from one development run (ts removed, %d event line(s))\n"
            % len(lines)) + "\n".join(lines) + "\n"
    (DEST / "run.log").write_text(sanitize(body, pairs), encoding="utf-8", newline="\n")
    written["run.log"] = (DEST / "run.log").stat().st_size

    # ---- contact sheet -------------------------------------------------------------- #
    pngs = []
    for entry in (render.get("pngs") or []):
        path = Path(str(entry.get("path")))
        if path.is_file():
            pngs.append((entry.get("slide") or len(pngs) + 1, path))
    if not pngs:
        for cand in sorted(out.rglob("*.png")) + sorted(work.rglob("*.png")):
            pngs.append((len(pngs) + 1, cand))
    if not pngs:
        print("no rendered PNG found; contact sheet skipped", file=sys.stderr)
        return 1

    from PIL import Image, ImageDraw
    tw = args.thumb_width
    cols = 4
    rows = (len(pngs) + cols - 1) // cols
    pad, label_h = 8, 16
    sheets = []
    for _slide, path in pngs:
        img = Image.open(str(path)).convert("RGB")
        th = max(1, int(img.height * tw / img.width))
        sheets.append(img.resize((tw, th), Image.LANCZOS))
    cell_h = max(s.height for s in sheets) + label_h
    W = cols * tw + (cols + 1) * pad
    H = rows * cell_h + (rows + 1) * pad
    sheet = Image.new("RGB", (W, H), (245, 246, 248))
    draw = ImageDraw.Draw(sheet)
    for i, (img, (slide_no, _p)) in enumerate(zip(sheets, pngs)):
        r, c = divmod(i, cols)
        x = pad + c * (tw + pad)
        y = pad + r * (cell_h + pad)
        sheet.paste(img, (x, y))
        draw.text((x + 2, y + img.height + 2), "slide %s" % slide_no, fill=(60, 60, 60))
    draw.text((pad, 2), "development example run - synthetic fixture - NOT an official result",
              fill=(150, 30, 30))
    sheet_path = DEST / "contact-sheet.png"
    sheet.save(str(sheet_path), optimize=True)
    raw = sheet_path.stat().st_size
    if raw > 150_000:
        # a contact sheet is a thumbnail grid, not a deliverable preview: the submission
        # budget is 2 MB total, so palette-quantize it down.
        sheet.convert("P", palette=Image.ADAPTIVE, colors=96).save(str(sheet_path),
                                                                   optimize=True)
    written["contact-sheet.png"] = sheet_path.stat().st_size

    print(json.dumps({"dest": str(DEST), "files": written,
                      "total_bytes": sum(written.values()),
                      "slides_in_sheet": len(pngs)}, ensure_ascii=False, indent=2))
    print("RESULT: OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
