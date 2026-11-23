#!/usr/bin/env python3
"""run.py — the one-shot orchestrator for the ARMI:2611.0000744 PPT Skill.

Contract: ``work/CONTRACT.md`` §3.  Usage::

    python run.py --case <case.json> --work <workdir> --out <outdir>
                  [--only ingest,plan,build] [--skip-render]

It runs the 8-CLI pipeline in order, each stage as a **separate subprocess**
writing its own JSON artifact, then copies the deliverables into ``--out`` and
finally asks ``result_contract.py`` for ``result.json``.

Stages
------
``ingest`` → ``material.json`` · ``plan`` → ``plan.json`` · ``build`` →
``deck.pptx`` + ``build_report.json`` · ``render`` → ``render/deck.pdf`` +
``render/slide-XX.png`` + ``render_report.json`` · ``evidence`` →
``evidence_map.json`` · ``verify`` → ``verify_report.json`` · ``result`` →
``result.json``.

Partial-result safety (non-negotiable)
-------------------------------------
Every stage's stdout/stderr and duration go to ``<out>/run.log``.  When a stage
fails, the orchestrator **stops running stages that depend on it**, still copies
everything that does exist into ``--out``, writes the version records, and still
calls ``result_contract.py`` so that ``result.json`` is emitted with an honest
``status="partial"``.  A non-zero exit never happens without a partial result on
disk.

Exit codes
----------
``0`` all requested stages succeeded **and** ``verify_deck.py`` returned PASS;
``1`` verify did not pass (or its report is missing while it was requested) —
partial deliverables are already on disk; ``2`` usage / IO error;
``3`` at least one requested stage failed (partial deliverables are on disk).
"""
from __future__ import annotations

import sys as _sys

_sys.dont_write_bytecode = True

import argparse
import glob
import json
import os
import shutil
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in _sys.path:
    _sys.path.insert(0, HERE)

import common as C  # noqa: E402

#: stage name -> (script, extra args, human description)
STAGES = [
    ("ingest", "ingest.py", [], "extract material.json from the real sources"),
    ("plan", "plan.py", [], "page roles, claims, bullets, evidence and layout hints"),
    ("build", "build_deck.py", [], "editable deck.pptx from the case template"),
    ("render", "render.py", [], "PDF + PNG previews via LibreOffice Kit"),
    ("evidence", "evidence.py", [], "page-level evidence map"),
    ("verify", "verify_deck.py", [], "V01-V15 self-check gate"),
    ("result", "result_contract.py", ["--out", "{out}"], "result.json"),
]
STAGE_ORDER = [name for name, _s, _a, _d in STAGES]
STAGE_SCRIPT = {name: script for name, script, _a, _d in STAGES}
STAGE_EXTRA = {name: extra for name, _s, extra, _d in STAGES}
STAGE_DESC = {name: desc for name, _s, _a, desc in STAGES}

#: which stage produces which artifact (used for the copy-to-out step and for
#: deciding what a failure invalidates)
STAGE_OUTPUTS = {
    "ingest": ["material.json"],
    "plan": ["plan.json"],
    "build": ["deck.pptx", "build_report.json"],
    "render": ["render_report.json", "render"],
    "evidence": ["evidence_map.json"],
    "verify": ["verify_report.json"],
    "result": ["result.json"],
}

#: stages that cannot run without these artifacts
STAGE_REQUIRES = {
    "plan": ["material.json"],
    "build": ["plan.json"],
    "render": ["deck.pptx"],
    "evidence": ["plan.json"],
    "verify": ["plan.json"],
    "result": [],
}

COPY_FILES = ["deck.pptx", "deck.pdf", "material.json", "plan.json",
              "build_report.json", "render_report.json", "evidence_map.json",
              "verify_report.json", "versions.json", "change_records.json",
              "diff-v1.md", "rollback.md"]
COPY_DIRS = ["render", "figures", "derived"]


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def parse_only(value):
    """Parse ``--only`` into an ordered, validated stage list."""
    if not value:
        return list(STAGE_ORDER)
    wanted = [part.strip() for part in value.split(",") if part.strip()]
    unknown = [name for name in wanted if name not in STAGE_ORDER]
    if unknown:
        raise SystemExit("error: --only names unknown stage(s): %s (known: %s)"
                         % (", ".join(unknown), ", ".join(STAGE_ORDER)))
    return [name for name in STAGE_ORDER if name in wanted]


def run_stage(name, case_path, work_dir, out_dir, log):
    """Run one stage as a subprocess; returns a result dict (never raises)."""
    script = os.path.join(HERE, STAGE_SCRIPT[name])
    if not os.path.isfile(script):
        log.error("stage script is missing", stage=name, script=STAGE_SCRIPT[name])
        return {"stage": name, "ok": False, "exit_code": None,
                "error": "script %s does not exist" % STAGE_SCRIPT[name],
                "duration_ms": 0}
    command = [_sys.executable, script, "--case", case_path, "--work", work_dir]
    for extra in STAGE_EXTRA[name]:
        command.append(extra.replace("{out}", out_dir))
    if name == "verify":
        # verify_deck.py must judge the *output* directory: V13 (result.json carries
        # every contract key) and V15 (version records exist) are about deliverables,
        # and both live in --out.
        command += ["--out", out_dir]
    env = dict(os.environ)
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    started = C.monotonic()
    try:
        completed = subprocess.run(command, cwd=HERE, env=env, capture_output=True,
                                   text=True, encoding="utf-8", errors="replace",
                                   timeout=1800)
        code = completed.returncode
        stdout, stderr = completed.stdout or "", completed.stderr or ""
    except subprocess.TimeoutExpired:
        code, stdout, stderr = None, "", "timed out after 1800s"
    except OSError as exc:
        code, stdout, stderr = None, "", "could not start: %s" % exc
    duration = int((C.monotonic() - started) * 1000)
    tail = (stdout.strip().splitlines() or [""])[-1][:400]
    log.event("stage", stage=name, exit_code=code, duration_ms=duration,
              ok=(code == 0), command=" ".join(os.path.basename(part)
                                               for part in command),
              description=STAGE_DESC[name], stdout_tail=tail)
    if code != 0:
        for line in (stderr or "").strip().splitlines()[-5:]:
            log.warn(line[:400], stage=name)
    return {"stage": name, "ok": code == 0, "exit_code": code, "duration_ms": duration,
            "stdout_tail": tail, "stderr_tail": (stderr or "")[-600:],
            "error": None if code == 0 else ("exit %s" % code)}


def missing_inputs(name, work_dir):
    return [item for item in STAGE_REQUIRES.get(name, [])
            if not os.path.exists(os.path.join(work_dir, item))]


def copy_artifacts(work_dir, out_dir, log):
    """Copy every deliverable that exists; returns the list of copied paths."""
    copied = []
    C.ensure_dir(out_dir)
    for name in COPY_FILES:
        source = os.path.join(work_dir, name)
        if os.path.isfile(source):
            target = os.path.join(out_dir, name)
            shutil.copy2(source, target)
            copied.append(name)
    for name in COPY_DIRS:
        source = os.path.join(work_dir, name)
        if os.path.isdir(source):
            target = os.path.join(out_dir, name)
            C.ensure_dir(target)
            count = 0
            for entry in sorted(os.listdir(source)):
                full = os.path.join(source, entry)
                if os.path.isfile(full):
                    shutil.copy2(full, os.path.join(target, entry))
                    count += 1
            if count:
                copied.append("%s/ (%d files)" % (name, count))
    if log:
        log.event("copy", out=out_dir, copied=copied)
    return copied


def write_version_records(out_dir, case_id, stage_results, plan):
    """V15: version records = versions.json + change_records.json + diff + rollback."""
    changes = []
    for index, result in enumerate(stage_results, start=1):
        changes.append({
            "id": "c%d" % index,
            "version": "v1",
            "kind": "stage-%s" % result["stage"],
            "what": "ran %s (%s)" % (result["stage"], STAGE_DESC.get(result["stage"], "")),
            "why": "the pipeline contract requires this stage for a complete deliverable",
            "ok": bool(result.get("ok")),
            "exit_code": result.get("exit_code"),
            "duration_ms": result.get("duration_ms"),
            "artifact": (STAGE_OUTPUTS.get(result["stage"]) or [None])[0],
            "rollback": "v1 is the initial build; delete the output directory and re-run "
                        "run.py to regenerate it",
        })
    failed = [c for c in changes if not c["ok"]]
    rollback = ("v1 is the initial build of this deck. Nothing earlier exists, so the "
                "rollback instruction is: delete the produced files in this output "
                "directory and re-run `python scripts/run.py --case <case.json> --work "
                "<workdir> --out <this dir>`; every artifact is regenerated from the "
                "case inputs. Do NOT hand-edit deck.pptx — edit plan.json and rebuild.")
    versions = {
        "schema": "versions/v1",
        "case_id": case_id,
        "current": "v1",
        "created": C.now_iso(),
        "versions": [{"version": "v1", "created": C.now_iso(),
                      "status": "partial" if failed else "ok",
                      "stages_ok": sum(1 for c in changes if c["ok"]),
                      "stages_total": len(changes),
                      "changes_n": len(changes)}],
        "rollback": rollback,
    }
    C.safe_write_json(os.path.join(out_dir, "versions.json"), versions)
    C.safe_write_json(os.path.join(out_dir, "change_records.json"),
                      {"schema": "change_records/v1", "case_id": case_id,
                       "changes": changes})
    C.write_text(os.path.join(out_dir, "rollback.md"), rollback + "\n")

    lines = ["# Per-slide record for v1", "",
             "This is the **initial build**: there is no earlier version of this deck to",
             "diff against, so this file lists what v1 contains, slide by slide. Any",
             "later revision must be diffed against it.", ""]
    for slide in (plan or {}).get("slides", []):
        lines.append("## Slide %s — %s (%s)" % (slide.get("index"), slide.get("role"),
                                                slide.get("layout_name") or
                                                slide.get("layout_hint")))
        lines.append("")
        lines.append("* title: %s" % slide.get("title"))
        lines.append("* claim: %s" % slide.get("claim"))
        for bullet in slide.get("bullets") or []:
            lines.append("  * bullet: %s" % bullet)
        if slide.get("figure_ref"):
            lines.append("* figure: %s" % slide["figure_ref"])
        if slide.get("table_ref"):
            lines.append("* table: %s" % slide["table_ref"])
        lines.append("* notes: %d chars" % len(slide.get("notes") or ""))
        lines.append("")
    C.write_text(os.path.join(out_dir, "diff-v1.md"), "\n".join(lines))
    return versions, changes


def skipped_render_report(out_dir, reason):
    """An honest render report for ``--skip-render`` (never claims a preview)."""
    return {
        "schema": C.SCHEMA_RENDER_REPORT,
        "renderer": {"name": "none", "version": "", "command": "", "available": False},
        "pdf": {"path": "", "pages": 0, "bytes": 0, "sha256": ""},
        "pngs": [], "status": "skipped", "error": None,
        "uncertainty": [reason],
    }


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="run.py", description="Run the full ingest → … → result pipeline.")
    parser.add_argument("--case", required=True, help="path to case.json")
    parser.add_argument("--work", required=True, help="working directory (artifacts)")
    parser.add_argument("--out", required=True, help="output directory (deliverables)")
    parser.add_argument("--only", default=None,
                        help="comma-separated subset of stages to run, in pipeline "
                             "order (default: all)")
    parser.add_argument("--skip-render", action="store_true",
                        help="skip the render stage (no PDF/PNG is claimed)")
    args = parser.parse_args(argv)

    case_path = os.path.abspath(args.case)
    work_dir = os.path.abspath(args.work)
    out_dir = os.path.abspath(args.out)
    if not os.path.isfile(case_path):
        print("error: case file not found: %s" % case_path, file=_sys.stderr)
        return 2
    C.ensure_dir(work_dir)
    C.ensure_dir(out_dir)
    log = C.RunLog(os.path.join(out_dir, "run.log"))
    only = parse_only(args.only)
    if args.skip_render and "render" in only:
        only = [name for name in only if name != "render"]
    log.event("run_start", case=os.path.basename(case_path), work=work_dir, out=out_dir,
              stages=only, skip_render=bool(args.skip_render))

    case = C.read_json(case_path, {}) or {}
    case_id = C.extract_case_id(case, case_path)

    # The gate (V13) reads result.json, so `result` runs before `verify`, and is then
    # re-run so that its recorded status/_self_check reflect the real verify verdict.
    # This is the only ordering that lets verify judge a complete output directory.
    build_stages = [name for name in only if name not in ("result", "verify")]
    stage_results, failed = [], []

    def execute(name, break_on_failure=True):
        absent = missing_inputs(name, work_dir)
        if absent:
            log.error("stage skipped: required artifact missing", stage=name,
                      missing=absent)
            stage_results.append({"stage": name, "ok": False, "exit_code": None,
                                  "duration_ms": 0,
                                  "error": "missing input(s): %s" % ", ".join(absent)})
            failed.append(name)
            return False
        with log.stage("stage_%s" % name):
            result = run_stage(name, case_path, work_dir, out_dir, log)
        stage_results.append(result)
        if not result["ok"]:
            failed.append(name)
            return False
        return True

    for name in build_stages:
        if not execute(name):
            break
        if name == "render":
            copy_artifacts(work_dir, out_dir, log)
    if "evidence" in build_stages:
        copy_artifacts(work_dir, out_dir, log)

    if args.skip_render:
        report_path = os.path.join(work_dir, "render_report.json")
        if not os.path.isfile(report_path):
            C.safe_write_json(report_path, skipped_render_report(
                out_dir, "rendering was skipped by --skip-render; no PDF or PNG preview "
                         "is claimed for this run"))
            log.warn("render skipped: wrote an honest status=skipped render_report.json")

    plan = C.read_json(os.path.join(work_dir, "plan.json"), {}) or {}
    copied = copy_artifacts(work_dir, out_dir, log)

    if "result" in only and not failed:
        # result_contract.py needs the deliverables in --out (artifacts are re-checked)
        stage_results = [r for r in stage_results if r["stage"] != "result"]
        execute("result")
        copy_artifacts(work_dir, out_dir, log)

    write_version_records(out_dir, case_id, stage_results, plan)

    verify_verdict, verify_requested = None, "verify" in only
    if verify_requested and not failed:
        stage_results = [r for r in stage_results if r["stage"] != "verify"]
        execute("verify")
        copy_artifacts(work_dir, out_dir, log)
        verify_report = C.read_json(os.path.join(out_dir, "verify_report.json"), None)
        verify_verdict = (verify_report or {}).get("verdict")
        if "result" in only:
            # rewrite result.json so status/_self_check carry the verify verdict
            stage_results = [r for r in stage_results if r["stage"] != "result"]
            execute("result")
            copy_artifacts(work_dir, out_dir, log)
    elif verify_requested:
        verify_report = C.read_json(os.path.join(out_dir, "verify_report.json"), None)
        verify_verdict = (verify_report or {}).get("verdict")

    result_exists = os.path.isfile(os.path.join(out_dir, "result.json"))
    if "result" in only and not result_exists:
        log.warn("result.json is missing after the result stage")

    log.event("run_end", failed=failed, verify_verdict=verify_verdict,
              copied=copied, exit_planned=None)
    print(json.dumps({"out": out_dir, "stages": {r["stage"]: r.get("exit_code")
                                                for r in stage_results},
                      "failed": failed, "verify_verdict": verify_verdict,
                      "result_json": result_exists}, ensure_ascii=False))

    if failed:
        return 3
    if verify_requested and verify_verdict != "PASS":
        print("verify_deck.py verdict is %r; partial deliverables are in %s"
              % (verify_verdict, out_dir), file=_sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
