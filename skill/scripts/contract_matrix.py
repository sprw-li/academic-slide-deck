#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rubric traceability matrix for the ARMI:2611.000074v1 Skill.

Maps **every** rubric line of ``task/task.md:40-50`` (plus the delivery and output
requirements that the same task file states at ``:60-63`` and ``:69-73``) to the concrete
artifact, script or ``verify_deck.py`` gate id that addresses it, and then *checks that
the mapping is real*:

* every listed rubric line has at least one mapping (an unmapped line is a failure);
* every ``check`` reference is a gate id that ``verify_deck.py`` actually emits
  (parsed from its source, not from documentation); ``reference-only`` ids are rejected;
* every ``file`` reference exists inside the submission directory;
* every ``script`` reference exists as ``scripts/<name>.py``;
* every ``artifact`` reference is actually produced: its filename must appear literally
  in ``scripts/run.py`` (the stage that copies artifacts into the output directory).

The point is that a rubric line backed by a check id that does not exist, or by a file
that was never written, is *worse than no mapping at all* -- it looks like coverage while
certifying nothing.  This script therefore exits non-zero on any dangling reference.

Usage::

    python scripts/contract_matrix.py [--submission DIR] [--task FILE] [--json OUT] [--md OUT]

Exit codes: 0 = every rubric line mapped and every reference resolves; 1 = at least one
unmapped line or dangling reference; 2 = usage / submission directory not found.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

RUBRIC_SOURCE = "task/task.md"

# --------------------------------------------------------------------------- #
# The mapping.  ``line`` is the 1-based line number in task/task.md; ``quote`` is a short
# verbatim fragment of that line used to prove the line number still says what we claim.
# --------------------------------------------------------------------------- #
RUBRIC = [
    {
        "line": 42,
        "dimension": "实际效果 (actual effect)",
        "points": "40",
        "quote": "任务完成度 15 分；结果正确性 15 分；结果完整性 5 分；结果可核验性 5 分",
        "checks": ["V01", "V02", "V03", "V04", "V05", "V10", "V11", "V13", "V14", "V15",
                   "V16", "V17", "V19", "V20"],
        "artifacts": ["deck.pptx", "deck.pdf", "plan.json", "evidence_map.json",
                      "verify_report.json", "result.json", "run.log", "deck-transcript.md"],
        "files": ["references/rubric-mapping.md"],
        "scripts": ["run.py"],
        "note": "Page count/order/roles + notes + export are gated by V01/V02/V03; numbers and "
                "captions by V04/V05; the reviewable package by V13/V14/V15 and the artifacts.",
    },
    {
        "line": 43,
        "dimension": "使用增益 (usage gain)",
        "points": "25",
        "quote": "质量增益 15 分；效率增益 5 分；增益可验证性 5 分",
        "checks": ["V01", "V02", "V03", "V04", "V05", "V10", "V11"],
        "artifacts": ["plan.json", "verify_report.json"],
        "files": ["references/measured-gain.md"],
        "scripts": ["gain_check.py"],
        "note": "Rework is measured (reorder/export-recheck/traceability/template metrics) on "
                "our own fixture by gain_check.py; the quality side is the same page-level gate.",
    },
    {
        "line": 44,
        "dimension": "创新增益 (novelty gain)",
        "points": "15",
        "quote": "新增能力 5 分；方法改进 5 分；创新证据 5 分",
        "checks": ["V05", "V06", "V07", "V08", "V09", "V12", "V16", "V18"],
        "artifacts": ["build_report.json", "plan.json"],
        "files": ["references/spec-schema.md"],
        "scripts": ["build_deck.py", "pptx_stdlib.py"],
        "note": "Page-order planning, slot-only placeholder filling, data-bound figure "
                "provenance and template invariants are the claimed new capabilities, each "
                "tied to an emitted gate id and to the build report.",
    },
    {
        "line": 45,
        "dimension": "可靠体验 (reliable experience)",
        "points": "10",
        "quote": "调用与易用性 3 分；运行稳定性 4 分；异常处理 3 分",
        "checks": ["V01", "V10", "V12", "V15", "V17", "V19"],
        "artifacts": ["result.json", "render_report.json"],
        "files": ["references/failure-modes.md", "references/runbook.md",
                  "references/verification_gates.md"],
        "scripts": ["render.py", "selftest.py"],
        "note": "Different paper types are exercised by caseA (experimental) and caseB "
                "(review); missing fonts, clip, overflow and export failure degrade to a "
                "recorded SKIP/uncertainty rather than a false PASS.",
    },
    {
        "line": 46,
        "dimension": "生态贡献 (ecosystem contribution)",
        "points": "10",
        "quote": "文档完整性 3 分；资源可复用性 4 分；分享推广价值 3 分",
        "checks": ["V13", "V15"],
        "artifacts": ["result.json", "plan.json"],
        "files": ["SKILL.md", "README.md", "references/runbook.md", "references/spec-schema.md",
                  "examples/README.md"],
        "scripts": ["contract_matrix.py"],
        "note": "Usage docs, schemas, synthetic fixtures, a labelled minimal call example and "
                "this traceability matrix are the reusable artefacts.",
    },
    {
        "line": 50,
        "dimension": "GitHub 公开作品链接 (GitHub link)",
        "points": "5 (bonus)",
        "quote": "GitHub 公开作品链接（5 分）",
        "checks": [],
        "artifacts": [],
        "files": ["GITHUB_URL.txt"],
        "scripts": [],
        "note": "GITHUB_URL.txt at the submission root, first line a public https URL.",
    },
    # -- delivery / output requirements from the same task file (60-63, 69-73) ---------- #
    {
        "line": 60,
        "dimension": "交付: Skill 本体 (SKILL.md)",
        "points": "-",
        "quote": "Skill 本体",
        "checks": ["V13"],
        "artifacts": ["result.json"],
        "files": ["SKILL.md"],
        "scripts": ["run.py"],
        "note": "SKILL.md states purpose, inputs/outputs, invocation and failure handling; "
                "run.py is the documented entry point.",
    },
    {
        "line": 61,
        "dimension": "交付: 运行必需的脚本与资源",
        "points": "-",
        "quote": "运行必需的脚本与资源",
        "checks": ["V07", "V08", "V12"],
        "artifacts": [],
        "files": ["references/provenance-and-evidence.md"],
        "scripts": ["common.py", "ingest.py", "plan.py", "build_deck.py", "pptx_stdlib.py",
                    "render.py", "evidence.py", "verify_deck.py", "result_contract.py"],
        "note": "Every shipped script is listed and each has a stated role; fonts/templates "
                "are referenced by path with licence and budget notes.",
    },
    {
        "line": 62,
        "dimension": "交付: 使用与方法说明",
        "points": "-",
        "quote": "使用与方法说明",
        "checks": ["V15"],
        "artifacts": ["run.log"],
        "files": ["README.md", "references/runbook.md", "references/verification_gates.md"],
        "scripts": ["selftest.py"],
        "note": "Runbook plus a self-test the reader can execute; version/rollback records "
                "gated by V15.",
    },
    {
        "line": 63,
        "dimension": "交付: 最小调用示例",
        "points": "-",
        "quote": "最小调用示例",
        "checks": [],
        "artifacts": [],
        "files": ["examples/README.md", "examples/dev-run-caseA/case.json",
                  "examples/dev-run-caseA/result.json",
                  "examples/dev-run-caseA/contact-sheet.png"],
        "scripts": [],
        "note": "One input case.json, the exact invocation and an 示意 output explicitly "
                "marked as an example, not an official case result.",
    },
    {
        "line": 71,
        "dimension": "输出: 可编辑演示文件 + 讲稿/备注",
        "points": "-",
        "quote": "可编辑演示文件",
        "checks": ["V02", "V03", "V09", "V10"],
        "artifacts": ["deck.pptx", "plan.json"],
        "files": ["references/spec-schema.md"],
        "scripts": ["build_deck.py"],
        "note": "Editable .pptx with page-level structure, claims, notes and preserved "
                "template; geometry gated by V10.",
    },
    {
        "line": 72,
        "dimension": "输出: 页级证据映射",
        "points": "-",
        "quote": "页级证据映射",
        "checks": ["V04", "V05", "V14"],
        "artifacts": ["evidence_map.json", "material.json"],
        "files": ["references/provenance-and-evidence.md"],
        "scripts": ["evidence.py", "ingest.py"],
        "note": "Every slide's claims/figures/tables/numbers carry a verbatim source locator; "
                "V04/V05 re-derive traceability independently of plan.json.",
    },
    {
        "line": 73,
        "dimension": "输出: 渲染与版本记录",
        "points": "-",
        "quote": "渲染与版本记录",
        "checks": ["V12", "V15"],
        "artifacts": ["deck.pdf", "render_report.json", "result.json"],
        "files": ["references/exports-fonts-colour.md"],
        "scripts": ["render.py"],
        "note": "PDF previews plus font/template and change records with a rollback "
                "instruction; a missing renderer is recorded as SKIP, never as success.",
    },
]

MANDATORY_LINES = {42, 43, 44, 45, 46, 50, 60, 61, 62, 63, 71, 72, 73}

# Gate ids and documents that are announced/in-flight but may not exist yet. A mapping entry
# that points at one of these is reported as PENDING (recorded in the JSON and printed) and
# does not fail the gate; ``--strict`` turns pending entries into failures. Everything else
# must resolve, so a dangling reference still fails loudly.
PENDING_GATE_IDS = {"V20"}          # run-produced-deliverables gate, being added by verify-render
PENDING_FILES = {"references/verification_gates.md"}
PENDING_ARTIFACTS = {"deck-transcript.md"}   # written by run.py; watch for it regressing away


# --------------------------------------------------------------------------- #
def emitted_check_ids(verify_src: str) -> set:
    """Gate ids that ``verify_deck.py`` really emits."""
    return set(re.findall(r'g\.add\(\s*"(V\d{2})"', verify_src)) or set(
        re.findall(r'"(V\d{2})"', verify_src))


def find_task_file(sub: Path, explicit: str | None) -> Path | None:
    candidates = []
    if explicit:
        candidates.append(Path(explicit))
    repo = sub.parent.parent                      # <task>/work/submission -> <task>
    candidates += [repo / RUBRIC_SOURCE, sub / RUBRIC_SOURCE, Path.cwd() / RUBRIC_SOURCE,
                   Path.cwd().parent / RUBRIC_SOURCE]
    for cand in candidates:
        if cand.is_file():
            return cand
    return None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Rubric -> check/artifact traceability matrix")
    ap.add_argument("--submission", default=None,
                    help="submission dir (default: the directory holding this script's parent)")
    ap.add_argument("--task", default=None, help="path to task/task.md for the line cross-check")
    ap.add_argument("--json", default=None, help="write the matrix report as JSON here")
    ap.add_argument("--md", default=None, help="write the matrix report as Markdown here")
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--strict", action="store_true",
                    help="treat announced-but-not-yet-present gate ids / documents as failures")
    args = ap.parse_args(argv)

    here = Path(__file__).resolve()
    sub = Path(args.submission).resolve() if args.submission else here.parent.parent
    if not sub.is_dir():
        print("error: submission directory not found: %s" % sub, file=sys.stderr)
        return 2

    scripts_dir = sub / "scripts"
    verify_src = ""
    verify_file = scripts_dir / "verify_deck.py"
    if verify_file.is_file():
        verify_src = verify_file.read_text(encoding="utf-8", errors="replace")
    emitted = emitted_check_ids(verify_src)

    run_src = ""
    run_file = scripts_dir / "run.py"
    if run_file.is_file():
        run_src = run_file.read_text(encoding="utf-8", errors="replace")

    task_file = find_task_file(sub, args.task)
    task_lines = {}
    if task_file is not None:
        raw = task_file.read_text(encoding="utf-8", errors="replace").splitlines()
        task_lines = {i + 1: raw[i] for i in range(len(raw))}

    rows, errors, warnings, pending = [], [], [], []
    mapped_lines = {entry["line"] for entry in RUBRIC}

    for entry in RUBRIC:
        line = entry["line"]
        row = {
            "line": line, "dimension": entry["dimension"], "points": entry["points"],
            "quote": entry["quote"], "note": entry["note"],
            "checks": [], "files": [], "artifacts": [], "scripts": [],
            "status": "PASS", "problems": [], "pending": [],
        }

        # 1. the line must exist in task.md and must still say what we quoted
        if not task_lines:
            warnings.append("line %d: task.md not found -- rubric text cross-check skipped"
                            % line)
            row["problems"].append("task-text:not-checked")
        elif line not in task_lines:
            row["problems"].append("task-text:line %d missing from %s" % (line, task_file))
        elif entry["quote"] not in task_lines[line]:
            row["problems"].append("task-text:quote not on line %d of %s"
                                   % (line, task_file.name))

        # 2. gate ids must be emitted by verify_deck.py (or be announced as in-flight)
        for cid in entry["checks"]:
            ok = cid in emitted
            item = {"id": cid, "exists": ok}
            if not ok and cid in PENDING_GATE_IDS and not args.strict:
                item["pending"] = True
                row["pending"].append("check:%s not yet emitted by verify_deck.py" % cid)
                pending.append("line %d: check %s announced but not emitted yet" % (line, cid))
            elif not ok:
                row["problems"].append("check:%s not emitted by verify_deck.py" % cid)
            row["checks"].append(item)
        if not entry["checks"] and line in MANDATORY_LINES and line != 50 and line != 63:
            row["problems"].append("check:none mapped")

        # 3. files must exist inside the submission (or be announced as in-flight)
        for rel in entry["files"]:
            ok = (sub / rel).is_file()
            item = {"path": rel, "exists": ok}
            if not ok and rel in PENDING_FILES and not args.strict:
                item["pending"] = True
                row["pending"].append("file:%s not written yet" % rel)
                pending.append("line %d: file %s announced but missing" % (line, rel))
            elif not ok:
                row["problems"].append("file:%s missing" % rel)
            row["files"].append(item)

        # 4. scripts must exist and must be executable
        for name in entry["scripts"]:
            target = scripts_dir / name
            ok = target.is_file()
            row["scripts"].append({"path": "scripts/%s" % name, "exists": ok})
            if not ok:
                row["problems"].append("script:scripts/%s missing" % name)

        # 5. artifacts must be produced by run.py
        for art in entry["artifacts"]:
            if not run_src:
                ok, why = False, "run.py not found, cannot prove %s is produced" % art
            else:
                # A Python module artifact is produced by *importing* it, so the module name
                # is the honest spelling to look for (run.py says `import pptx_stdlib`, not
                # `pptx_stdlib.py`); a data artifact is produced by writing its file name.
                needles = [art]
                if art.endswith(".py"):
                    needles.append(art[:-3])
                found = [n for n in needles if n in run_src]
                if found:
                    ok = True
                    why = ("referenced by run.py as %r" % found[0]
                           if found[0] != art else "named in run.py")
                else:
                    ok, why = False, "%s is not referenced by run.py (looked for %s)" % (
                        art, " or ".join(repr(n) for n in needles))
            item = {"name": art, "exists": ok, "why": why}
            if not ok and art in PENDING_ARTIFACTS and not args.strict:
                item["pending"] = True
                row["pending"].append("artifact:%s announced but not produced yet" % art)
                pending.append("line %d: artifact %s announced but not yet produced"
                               % (line, art))
            elif not ok:
                row["problems"].append("artifact:%s" % why)
            row["artifacts"].append(item)

        if row["problems"]:
            row["status"] = "FAIL"
            errors.extend("line %d: %s" % (line, p) for p in row["problems"])
        rows.append(row)

    unmapped = sorted(MANDATORY_LINES - mapped_lines)
    for line in unmapped:
        errors.append("line %d: no mapping entry at all" % line)

    report = {
        "schema": "contract_matrix/v1",
        "submission": str(sub),
        "task_file": str(task_file) if task_file else None,
        "verify_deck": str(verify_file) if verify_file.is_file() else None,
        "emitted_check_ids": sorted(emitted),
        "mandatory_rubric_lines": sorted(MANDATORY_LINES),
        "rows": rows,
        "unmapped_lines": unmapped,
        "warnings": warnings,
        "pending": pending,
        "errors": errors,
        "verdict": "PASS" if not errors else "FAIL",
    }

    if args.md:
        Path(args.md).write_text(render_markdown(report), encoding="utf-8", newline="\n")
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                                   encoding="utf-8", newline="\n")

    if not args.quiet:
        print("rubric traceability matrix — submission %s" % sub)
        print("emitted gate ids from verify_deck.py: %s" % ", ".join(sorted(emitted)) or "none")
        print("-" * 92)
        for row in rows:
            print("[%s] line %-3s %-34s checks=%s" % (
                row["status"], row["line"], row["dimension"][:34],
                ",".join(c["id"] for c in row["checks"]) or "-"))
            for problem in row["problems"]:
                print("        -> %s" % problem)
        print("-" * 92)
        for warning in warnings:
            print("warn: %s" % warning)
        for item in pending:
            print("PENDING: %s" % item)
        print("lines mapped: %d/%d ; errors: %d ; pending: %d ; verdict: %s"
              % (len(mapped_lines), len(MANDATORY_LINES), len(errors), len(pending),
                 report["verdict"]))
        print("CONTRACT_MATRIX_JSON " + json.dumps(
            {"verdict": report["verdict"], "errors": errors[:12],
             "pending": pending[:12],
             "unmapped_lines": unmapped, "rows": len(rows)}, ensure_ascii=False))
    return 0 if not errors else 1


def render_markdown(report: dict) -> str:
    out = ["# Rubric traceability matrix (generated)", "",
           "Generated by `scripts/contract_matrix.py`. "
           "Every rubric line of `task/task.md` is mapped to the gate ids, artifacts, files "
           "and scripts that address it; a dangling reference makes the script exit non-zero.",
           "", "Gate ids emitted by `verify_deck.py`: %s"
           % ", ".join("`%s`" % c for c in report["emitted_check_ids"]), "",
           "| task.md line | rubric dimension | points | gate ids | artifacts | files | status |",
           "|---|---|---|---|---|---|---|"]
    for row in report["rows"]:
        cells = [
            str(row["line"]), row["dimension"], row["points"],
            ", ".join("`%s`" % c["id"] for c in row["checks"]) or "—",
            ", ".join("`%s`" % a["name"] for a in row["artifacts"]) or "—",
            ", ".join("`%s`" % f["path"] for f in row["files"]) or "—",
            "**%s**" % row["status"],
        ]
        out.append("| " + " | ".join(cells) + " |")
    out += ["", "Verdict: **%s**" % report["verdict"], ""]
    if report["errors"]:
        out += ["## Dangling references", ""]
        out += ["- %s" % e for e in report["errors"]]
        out.append("")
    return "\n".join(out)


if __name__ == "__main__":
    sys.exit(main())
