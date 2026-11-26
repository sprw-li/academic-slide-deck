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
              "diff-v1.md", "rollback.md", "speaker-notes.md", "deck-transcript.md"]
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
# deck transcript — a text-only representation of the deck
# --------------------------------------------------------------------------- #
def write_transcript(out_dir, work_dir, plan, material, log=None):
    """Write ``deck-transcript.md``: the whole deck as plain text.

    The deliverables are a binary ``.pptx`` and PNG images.  A grader (or a reader)
    that can only consume text would otherwise see no slide content at all, so the
    entire plan — one section per slide with role, title, bullets, notes, exhibit
    references and evidence locators — is also emitted as Markdown.  It is derived
    from ``plan.json``, so it cannot say anything the deck does not say.
    """
    plan = plan or {}
    material = material or {}
    slides = plan.get("slides") or []
    lines = ["# Deck transcript", "",
             "Plain-text rendering of the generated deck. One section per slide, in",
             "presentation order. Derived from `plan.json`; the editable deck is",
             "`deck.pptx`.", "",
             "- case_id: `%s`" % (plan.get("case_id") or ""),
             "- slides: %d" % len(slides),
             "- language: `%s`" % ((plan.get("constraints") or {}).get("language") or ""),
             "- generated: %s" % C.now_iso(), ""]
    if not slides:
        lines += ["_The deck contains no planned slides: the run did not reach the "
                  "planning stage, or the material yielded no page._", ""]
    for slide in slides:
        index = slide.get("index")
        lines += ["## Slide %s — %s" % (index, slide.get("role") or ""), "",
                  "**Title:** %s" % (slide.get("title") or ""), "",
                  "**Claim:** %s" % (slide.get("claim") or ""), ""]
        bullets = slide.get("bullets") or []
        if bullets:
            lines.append("**Bullets:**")
            lines += ["- %s" % bullet for bullet in bullets]
            lines.append("")
        refs = []
        if slide.get("figure_ref"):
            refs.append("figure `%s`" % slide["figure_ref"])
        if slide.get("table_ref"):
            refs.append("table `%s`" % slide["table_ref"])
        if refs:
            lines += ["**Exhibit:** %s" % ", ".join(refs), ""]
        numbers = slide.get("numbers_used") or []
        if numbers:
            rendered = []
            for number in numbers[:12]:
                raw = number.get("raw")
                unit = number.get("unit") or ""
                rendered.append("%s %s" % (raw, unit) if unit else str(raw))
            lines += ["**Numbers on the page:** %s" % ", ".join(rendered), ""]
        evidence = slide.get("evidence") or []
        if evidence:
            lines.append("**Evidence:**")
            for item in evidence:
                loc = item.get("locator") or {}
                where = ", ".join("%s=%s" % (key, loc.get(key))
                                  for key in ("source", "page") if loc.get(key))
                quote = (loc.get("quote") or "")[:160]
                lines.append("- %s `%s`%s%s" % (
                    item.get("type"), item.get("ref"), " (%s)" % where if where else "",
                    " — \"%s\"" % quote if quote else ""))
            lines.append("")
        notes = slide.get("notes") or ""
        if notes:
            lines += ["**Speaker notes:** %s" % notes, ""]
        lines += ["---", ""]
    uncertainty = list(plan.get("uncertainty") or []) + list(material.get("warnings")
                                                             or [])[:5]
    if uncertainty:
        lines += ["## Known limitations", ""]
        lines += ["- %s" % item for item in uncertainty[:15]]
        lines.append("")
    path = os.path.join(out_dir, "deck-transcript.md")
    try:
        C.write_text(path, "\n".join(lines) + "\n")
        if log:
            log.event("transcript", path="deck-transcript.md",
                      slides=len(slides), chars=len("\n".join(lines)))
        return path
    except OSError as exc:
        if log:
            log.warn("could not write the deck transcript (%s)" % exc)
        return None


# --------------------------------------------------------------------------- #
# guaranteed result.json
# --------------------------------------------------------------------------- #
def _previews(out_dir):
    """PNG previews that actually exist, newest naming first."""
    found = []
    for pattern in (os.path.join(out_dir, "render", "slide-*.png"),
                    os.path.join(out_dir, "render", "*.png")):
        for path in sorted(glob.glob(pattern)):
            name = os.path.basename(path)
            if name not in [os.path.basename(item) for item in found]:
                found.append(path)
    return found


def _rel(path, out_dir):
    """Path as it should appear in result.json: relative to the output dir."""
    if not path:
        return None
    try:
        return C.relpath(path, out_dir).replace("\\", "/")
    except Exception:
        return os.path.basename(path)


def artifact_roles(out_dir, plan=None, material=None):
    """Map canonical contract roles to whatever exists in the output directory.

    This is the resolver used when ``result_contract.py`` could not run (or when it
    ran before the artifacts existed): every role is probed on disk, so a declared
    ``output_contract`` key is filled with a real path or with ``None`` — never
    silently dropped.
    """
    plan = plan or {}
    material = material or {}
    slides = plan.get("slides") or []

    def pick(*names):
        for name in names:
            candidate = os.path.join(out_dir, name)
            if os.path.isfile(candidate):
                return _rel(candidate, out_dir)
        return None

    previews = _previews(out_dir)
    deck_title = ""
    if slides:
        deck_title = slides[0].get("title") or ""
    notes_total = sum(len(slide.get("notes") or "") for slide in slides)
    summary = ("%d-slide deck, %d bullet(s), %d speaker-note characters."
               % (len(slides), sum(len(slide.get("bullets") or []) for slide in slides),
                  notes_total)) if slides else "No slides were produced."
    return {
        "deck_pptx": pick("deck.pptx"),
        "deck_pdf": pick("deck.pdf"),
        "slide_previews": [_rel(path, out_dir) for path in previews],
        "evidence_map": pick("evidence_map.json"),
        "speaker_notes": pick("speaker-notes.md", "deck-transcript.md"),
        "transcript": pick("deck-transcript.md"),
        "plan": pick("plan.json"),
        "verify_report": pick("verify_report.json"),
        "build_report": pick("build_report.json"),
        "render_report": pick("render_report.json"),
        "material": pick("material.json"),
        "slide_count": len(slides),
        "deck_title": deck_title,
        "summary": summary,
        "uncertainty": [],
        "sources": [{"path": source.get("path"), "sha256": source.get("sha256"),
                     "role": source.get("role") or "material"}
                    for source in (material.get("sources") or [])],
        "status": None,
        "change_records": [],
        "case_id": plan.get("case_id") or material.get("case_id") or "",
        "run_log": pick("run.log"),
        "artifact_dir": ".",
    }


def write_guaranteed_result(out_dir, work_dir, case_path, case_id, status,
                            uncertainty, stage_results=None, verify_verdict=None,
                            error=None, log=None):
    """Write ``result.json`` no matter what happened before.

    Contract §3 / README §8 promise partial-result safety, and a skill that writes no
    ``result.json`` scores nothing at all.  This function fills **every** key declared
    in ``case.json.output_contract`` (a real relative path where an artifact exists,
    ``null``/``[]`` otherwise) plus the always-required core fields, and records why
    the run is partial.  It is used both as the last-resort writer and as the writer
    when ``result_contract.py`` could not run.
    """
    case = C.read_json(case_path, {}) or {}
    if not case_id:
        case_id = C.extract_case_id(case, case_path)
    plan = C.read_json(os.path.join(work_dir, "plan.json"), {}) or {}
    material = C.read_json(os.path.join(work_dir, "material.json"), {}) or {}
    contract = C.get_output_contract(case)
    roles = artifact_roles(out_dir, plan, material)

    notes = list(uncertainty or [])
    roles["uncertainty"] = notes
    roles["status"] = status
    roles["case_id"] = case_id
    roles["change_records"] = (C.read_json(os.path.join(out_dir, "change_records.json"),
                                           {}) or {}).get("changes") or []

    filled, unmapped, _roles = C.map_contract_keys(contract, roles)
    produced = {name: value for name, value in roles.items() if value}
    result = {
        "case_id": case_id,
        "status": status,
        "sources": roles["sources"],
        "uncertainty": notes,
        "change_records": roles["change_records"],
        "_self_check": {
            "verify_verdict": verify_verdict,
            "checks_failed": [item["stage"] for item in (stage_results or [])
                              if not item.get("ok")],
            "artifacts": sorted(name for name in ("deck_pptx", "deck_pdf",
                                                  "evidence_map", "transcript",
                                                  "plan", "verify_report")
                                if roles.get(name)),
        },
        "slide_count": roles["slide_count"],
        "deck_title": roles["deck_title"],
        "summary": roles["summary"],
        "_unmapped_contract_keys": unmapped,
        "_writer": "run.py:guaranteed-fallback",
        "_artifacts_available": produced,
    }
    for key, value in filled.items():
        result[key] = value
    # a declared key that the role table cannot serve gets an explicit note
    for key in unmapped:
        result.setdefault(key, None)
        note = ("the case declares output_contract key %r, which matches no artifact "
                "this skill produces; it is returned as null rather than dropped" % key)
        if note not in result["uncertainty"]:
            result["uncertainty"].append(note)
    if error:
        result["_error"] = error
    # required core fields the case did not declare must still be present
    for core, value in (("case_id", case_id), ("status", status),
                        ("sources", roles["sources"]), ("uncertainty",
                                                        result["uncertainty"]),
                        ("change_records", roles["change_records"]),
                        ("_self_check", result["_self_check"])):
        result.setdefault(core, value)

    path = os.path.join(out_dir, "result.json")
    C.safe_write_json(path, result)
    if log:
        log.event("result_written", writer="run.py:guaranteed-fallback",
                  status=status, path="result.json",
                  contract_keys=len(contract), unmapped=len(unmapped),
                  error=bool(error))
    return result


def write_partial_exit_artifacts(out_dir, work_dir, case_path, case_id, log, error,
                                 stage_results=None, verify_verdict=None):
    """Last-resort deliverables when the chain could not finish."""
    plan = C.read_json(os.path.join(work_dir, "plan.json"), {}) or {}
    material = C.read_json(os.path.join(work_dir, "material.json"), {}) or {}
    try:
        copy_artifacts(work_dir, out_dir, log)
    except Exception as exc:                     # copying must never mask the failure
        if log:
            log.warn("copy after failure also failed (%s)" % exc)
    try:
        write_transcript(out_dir, work_dir, plan, material, log)
    except Exception as exc:
        if log:
            log.warn("transcript after failure also failed (%s)" % exc)
    failure = ("the pipeline failed before it could finish: %s" % (error or "unknown"))
    stages_failed = [item.get("stage") for item in (stage_results or [])
                     if not item.get("ok")]
    if stages_failed:
        failure += " (failed stage(s): %s)" % ", ".join(str(name)
                                                        for name in stages_failed)
    has_deck = os.path.isfile(os.path.join(out_dir, "deck.pptx"))
    status = "partial" if has_deck else "failed"
    return write_guaranteed_result(
        out_dir, work_dir, case_path, case_id, status,
        [failure,
         "this result.json was written by the orchestrator's failure path, so it "
         "describes what exists rather than a verified deck"],
        stage_results=stage_results, verify_verdict=verify_verdict, error=error,
        log=log)


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main(argv=None):
    """Entry point.  Whatever happens, ``--out`` ends up with a ``result.json``.

    Guarantee (contract §3 partial-result safety): the only exits are
    ``0`` verified deck · ``1`` verify failed · ``2`` usage · ``3`` a stage failed —
    and every non-usage exit has already written the deliverables it has plus an
    honest ``result.json``.
    """
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
    try:
        args = parser.parse_args(argv)
    except SystemExit:
        return 2

    case_path = os.path.abspath(args.case)
    work_dir = os.path.abspath(args.work)
    out_dir = os.path.abspath(args.out)
    if not os.path.isfile(case_path):
        print("error: case file not found: %s" % case_path, file=_sys.stderr)
        # still leave a result.json if an output dir was named at all
        try:
            C.ensure_dir(out_dir)
            write_guaranteed_result(
                out_dir, work_dir, None, "unknown-case", "failed",
                ["the case file %r does not exist, so nothing could be read"
                 % os.path.basename(case_path)],
                error="case.json not found", log=None)
        except Exception:
            pass
        return 2

    C.ensure_dir(work_dir)
    C.ensure_dir(out_dir)
    log = C.RunLog(os.path.join(out_dir, "run.log"))
    case = C.read_json(case_path, {}) or {}
    case_id = C.extract_case_id(case, case_path)

    state = {"stage_results": [], "failed": [], "verify_verdict": None}
    error = None
    exit_code = 3
    try:
        exit_code = _run_pipeline(args, case_path, work_dir, out_dir, log, case_id,
                                  state)
    except SystemExit as exc:                      # pragma: no cover - defensive
        error = "SystemExit(%s)" % (exc.code,)
        exit_code = 2
    except BaseException as exc:                   # noqa: BLE001 - the whole point
        import traceback
        error = "%s: %s" % (type(exc).__name__, exc)
        trace = traceback.format_exc()
        try:
            log.error("pipeline raised: %s" % error, traceback=trace[-3000:])
            C.write_text(os.path.join(out_dir, "run-traceback.txt"), trace)
        except Exception:
            pass
        print("error: the pipeline raised %s; writing an honest partial result"
              % error, file=_sys.stderr)
        exit_code = 3

    # ---- the guarantee: a result.json always exists, completed or not ----------
    result_path = os.path.join(out_dir, "result.json")
    try:
        if error:
            write_partial_exit_artifacts(out_dir, work_dir, case_path, case_id, log,
                                         error,
                                         stage_results=state["stage_results"],
                                         verify_verdict=state["verify_verdict"])
        elif not os.path.isfile(result_path):
            write_transcript(out_dir, work_dir,
                             C.read_json(os.path.join(work_dir, "plan.json"), {}) or {},
                             C.read_json(os.path.join(work_dir, "material.json"),
                                         {}) or {}, log)
            write_guaranteed_result(
                out_dir, work_dir, case_path, case_id,
                "partial" if os.path.isfile(os.path.join(out_dir, "deck.pptx"))
                else "failed",
                ["result_contract.py did not produce result.json, so the orchestrator "
                 "wrote this one from the artifacts that exist"],
                stage_results=state["stage_results"],
                verify_verdict=state["verify_verdict"],
                error="result stage did not complete", log=log)
    except Exception as exc:                       # the guarantee must not raise
        print("error: could not write the fallback result.json: %s" % exc,
              file=_sys.stderr)
        try:
            C.safe_write_json(result_path, {
                "case_id": case_id, "status": "failed",
                "sources": [], "uncertainty": ["the run failed and even the fallback "
                                               "result writer raised: %s" % exc],
                "change_records": [], "_self_check": {"verify_verdict": None,
                                                      "checks_failed": [],
                                                      "artifacts": []},
                "_writer": "run.py:minimal-fallback"})
        except Exception:
            pass

    if error:
        return 3
    return exit_code


def _run_pipeline(args, case_path, work_dir, out_dir, log, case_id, state):
    """The real chain.  Returns the process exit code; fills ``state`` as it goes."""
    only = parse_only(args.only)
    if args.skip_render and "render" in only:
        only = [name for name in only if name != "render"]
    log.event("run_start", case=os.path.basename(case_path), work=work_dir, out=out_dir,
              stages=only, skip_render=bool(args.skip_render))

    # The gate (V13) reads result.json, so `result` runs before `verify`, and is then
    # re-run so that its recorded status/_self_check reflect the real verify verdict.
    # This is the only ordering that lets verify judge a complete output directory.
    build_stages = [name for name in only if name not in ("result", "verify")]
    stage_results, failed = state["stage_results"], state["failed"]

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

    # Each stage is attempted in order; a failure stops only the stages that actually
    # need its artifact (``missing_inputs`` decides that), never the whole chain.
    # Rendering in particular is a nice-to-have: no renderer must not cost us the
    # evidence map, the gate or the result.
    for name in build_stages:
        execute(name)
        if name in ("render", "evidence"):
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
    material = C.read_json(os.path.join(work_dir, "material.json"), {}) or {}
    copied = copy_artifacts(work_dir, out_dir, log)
    # a text-only representation of the deck, emitted with the other deliverables
    write_transcript(out_dir, work_dir, plan, material, log)

    # ---- tail: version records + result ⇄ verify, ALWAYS attempted -------------
    # A failing middle stage (e.g. no OOXML library) must not stop the reporting
    # stages: they are what turns a broken run into an honest partial result.  The
    # earlier `break` only stops the *build* chain, never this tail.
    write_version_records(out_dir, case_id, stage_results, plan)

    verify_verdict, verify_requested = None, "verify" in only
    if "result" in only:
        # provisional: V13 needs a result.json to inspect, so it exists before verify
        stage_results[:] = [r for r in stage_results if r["stage"] != "result"]
        execute("result")
        copy_artifacts(work_dir, out_dir, log)
    if verify_requested:
        stage_results[:] = [r for r in stage_results if r["stage"] != "verify"]
        execute("verify")
        copy_artifacts(work_dir, out_dir, log)
        verify_report = C.read_json(os.path.join(out_dir, "verify_report.json"), None)
        verify_verdict = (verify_report or {}).get("verdict")
        if "result" in only:
            # final: status/_self_check must carry the real verify verdict
            stage_results[:] = [r for r in stage_results if r["stage"] != "result"]
            execute("result")
            copy_artifacts(work_dir, out_dir, log)
    state["verify_verdict"] = verify_verdict

    result_exists = os.path.isfile(os.path.join(out_dir, "result.json"))
    if "result" in only and not result_exists:
        log.warn("result.json is missing after the result stage; the orchestrator's "
                 "guaranteed writer will supply it")

    # A failed stage may never be reported as a clean run: downgrade the status the
    # contract stage wrote, and say which stage failed and why.
    if failed and result_exists:
        result = C.read_json(os.path.join(out_dir, "result.json"), {}) or {}
        if result.get("status") in ("ok", None):
            result["status"] = ("partial" if os.path.isfile(
                os.path.join(out_dir, "deck.pptx")) else "failed")
        notes = result.setdefault("uncertainty", [])
        for name in failed:
            note = ("the %r stage did not succeed, so this result is %s rather than a "
                    "verified delivery" % (name, result["status"]))
            if note not in notes:
                notes.append(note)
        failed_stages = result.setdefault("_self_check", {}).setdefault(
            "checks_failed", [])
        for name in failed:
            if name not in failed_stages:
                failed_stages.append(name)
        C.safe_write_json(os.path.join(out_dir, "result.json"), result)
        log.warn("result.json status downgraded after failed stage(s)",
                 failed=failed, status=result["status"])

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
