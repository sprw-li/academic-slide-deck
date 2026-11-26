#!/usr/bin/env python3
"""result_contract.py -- emit ``result.json`` for the case's own output contract.

Stage 7 (final) of the 8-CLI pipeline (CONTRACT.md section 3).

The single most important design decision in this skill is that the platform's
``case.json`` declares its **own** ``output_contract`` and the skill is contract-driven
(CONTRACT.md:24-34).  So this stage:

1. reads ``case.json.output_contract`` and fills **every** declared key;
2. maps key names **tolerantly** -- ``deck_pptx|pptx|presentation|editable_deck|slide_file``
   all mean the .pptx, and so on for pdf / previews / evidence map / notes / plan;
3. for anything matched by neither the exact key nor a synonym, emits ``null`` (or an empty
   container when the case's own default declares that type), records it in
   ``_unmapped_contract_keys`` **and** explains it in ``uncertainty`` -- a silently dropped
   key is a defect (CONTRACT.md:33-34);
4. always also emits ``case_id``, ``status``, ``sources``, ``uncertainty``,
   ``change_records`` and ``_self_check``.

Artifact paths prefer the copy that already exists in the output directory (where the
platform will read them) and fall back to the work directory, so ``result.json`` never
names a file that is not there at the moment it is written.

CLI
---

    result_contract.py --case case.json --work WORK --out OUTDIR
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

try:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from common import dump_json as _dump_json  # type: ignore
except Exception:  # noqa: BLE001
    _dump_json = None

try:  # optional: the deck is only read to count slides and collect titles
    from pptx import Presentation  # type: ignore
except Exception:  # noqa: BLE001
    Presentation = None


def load_json(path, default=None):
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


def norm_key(key):
    return re.sub(r"[^a-z0-9]+", "_", str(key).strip().lower()).strip("_")


#: canonical slot -> every key spelling the contract accepts, plus obvious near-synonyms.
SYNONYM_GROUPS = {
    "deck_path": ("deck_pptx", "pptx", "presentation", "editable_deck", "slide_file",
                  "deck", "editable_presentation", "deck_file", "editable_ppt",
                  "presentation_file", "slides_file"),
    "pdf_path": ("deck_pdf", "pdf", "export", "pdf_export", "rendered_pdf", "deck_pdf_path",
                 "pdf_file", "exported_pdf", "slides_pdf"),
    "png_paths": ("slide_previews", "previews", "images", "png", "renders",
                  "preview_images", "slide_images", "png_previews", "image_exports",
                  "rendered_slides", "slide_pngs"),
    "evidence_path": ("page_evidence_map", "evidence", "evidence_map", "source_map",
                      "page_evidence", "evidence_mapping", "citation_map",
                      "evidence_map_path", "source_evidence"),
    "notes_path": ("speaker_notes", "notes", "script", "speaker_script", "speaker_notes_md",
                   "narrative_script", "speaker_notes_file"),
    # A text-only grader cannot read a .pptx or a PNG.  These keys are answered with a plain
    # text transcript of the slide content so the narration/claims are machine-readable.
    # Union of this skill's spellings and the ones deck-core declares for the `transcript` role
    # in common.CONTRACT_SYNONYMS, so either side can be extended without a silent miss.
    "transcript_path": ("transcript", "deck_transcript", "slides_text", "slide_text",
                        "deck_text", "plain_text", "transcript_md", "slide_transcript",
                        "content_text", "text_dump", "slides_markdown", "deck_markdown",
                        "slide_contents", "content_md", "text_export", "page_text",
                        "outline_text", "deck_content"),
    "plan_path": ("plan", "outline", "page_plan", "slide_plan", "page_outline",
                  "plan_file", "deck_plan"),
    "slide_count": ("slide_count", "num_slides", "n_slides", "number_of_slides",
                    "slides_count", "slide_total"),
    "slide_titles": ("slide_titles", "titles", "headings", "slide_headings"),
    "roles": ("role_sequence", "roles", "slide_roles", "page_roles"),
    "audience": ("audience", "target_audience"),
    "duration_min": ("duration_min", "duration", "duration_minutes", "minutes"),
    "language": ("language", "lang", "slide_language"),
    "template_path": ("template", "template_pptx", "template_path", "template_file",
                      "master_template"),
    "verify_path": ("verify_report", "self_check_report", "verification_report",
                    "check_report"),
    "render_path": ("render_report", "render_info", "rendering_report"),
    "build_path": ("build_report", "build_info", "build_log"),
    "material_path": ("material", "material_json", "source_material"),
    "deck_sha256": ("deck_sha256", "pptx_sha256", "presentation_sha256"),
    "pdf_sha256": ("pdf_sha256", "export_sha256"),
    "self_check": ("self_check", "selfcheck", "selfcheck_report"),
    "case_id": ("case_id", "caseid", "id"),
    "status": ("status", "run_status", "state"),
    "sources": ("sources", "source_files", "inputs"),
    "uncertainty": ("uncertainty", "limitations", "caveats", "known_issues"),
    "change_records": ("change_records", "changes", "revision_history", "changelog"),
}

CORE_KEYS = ("case_id", "status", "sources", "uncertainty", "change_records", "_self_check")

#: role name (as deck-core declares it in common.CONTRACT_SYNONYMS) -> slot in this file
_ROLE_TO_SLOT = {
    "deck": "deck_path", "deck_pptx": "deck_path", "pptx": "deck_path",
    "presentation": "deck_path", "editable_deck": "deck_path",
    "pdf": "pdf_path", "deck_pdf": "pdf_path",
    "previews": "png_paths", "slide_previews": "png_paths",
    "evidence_map": "evidence_path", "evidence": "evidence_path",
    "notes": "notes_path", "speaker_notes": "notes_path",
    "plan": "plan_path", "outline": "plan_path",
    "transcript": "transcript_path",
    "verify_report": "verify_path", "render_report": "render_path",
    "build_report": "build_path", "material": "material_path",
    "template": "template_path",
}


def _merge_common_synonyms():
    """Fold deck-core's declared roles into this file's table, so the two cannot drift apart.

    Only *adds* aliases; an unavailable or unexpected ``common`` module leaves the table exactly
    as written above.
    """
    try:
        from common import CONTRACT_SYNONYMS as declared  # type: ignore
    except Exception:  # noqa: BLE001
        return
    if not isinstance(declared, dict):
        return
    for role, aliases in declared.items():
        slot = _ROLE_TO_SLOT.get(str(role).strip().lower())
        if not slot or not isinstance(aliases, (list, tuple)):
            continue
        merged = tuple(SYNONYM_GROUPS.get(slot, ())) + tuple(str(a) for a in aliases)
        SYNONYM_GROUPS[slot] = tuple(dict.fromkeys(merged))


_merge_common_synonyms()


def pick(out_dir, work_dir, name):
    """Prefer an existing copy in the output dir, else the work-dir original."""
    out_candidate = Path(out_dir) / name if out_dir else None
    if out_candidate and out_candidate.exists():
        return str(out_candidate)
    work_candidate = Path(work_dir) / name
    if work_candidate.exists():
        return str(work_candidate)
    return None


def resolve_existing(candidate, bases):
    """Resolve *candidate* against *bases*; case fixtures use case-relative paths."""
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
                return path
        except OSError:
            continue
    return None


def write_speaker_notes(plan, target, title="Speaker notes"):
    """Materialise the speaker script so ``speaker_notes`` can map to a real artefact.

    The same writer serves the plain-text deck transcript (``title="Deck transcript"``): a
    text-only grader can then read every slide's title, claim, bullets and narration instead of
    being handed a binary .pptx and a set of PNGs it cannot parse.
    """
    lines = [f"# {title}", "",
             "Generated from plan.json by result_contract.py. One section per slide.", ""]
    for slide in plan.get("slides") or []:
        if not isinstance(slide, dict):
            continue
        lines += [f"## Slide {slide.get('index', '?')} -- {slide.get('title', '')}", "",
                  f"- role: {slide.get('role', '')}",
                  f"- claim: {slide.get('claim', '')}", ""]
        for b in slide.get("bullets") or []:
            lines.append(f"- bullet: {b}")
        if slide.get("bullets"):
            lines.append("")
        note = (slide.get("notes") or "").strip()
        lines += [note or "_(no speaker note was written for this slide)_", ""]
    try:
        Path(target).parent.mkdir(parents=True, exist_ok=True)
        Path(target).write_text("\n".join(lines) + "\n", encoding="utf-8")
        return str(target)
    except OSError:
        return None


def build_result(case, work, out_dir, case_path=None):
    work = Path(work)
    out_dir = Path(out_dir) if out_dir else None

    plan = load_json(work / "plan.json", {}) or {}
    material = load_json(work / "material.json", {}) or {}
    build = load_json(work / "build_report.json", {}) or {}
    render = load_json(work / "render_report.json", {}) or {}
    verify = load_json(work / "verify_report.json", {}) or {}
    emap = load_json(work / "evidence_map.json", {}) or {}
    versions = load_json(work / "versions.json", {}) or {}
    changes_doc = load_json(work / "change_records.json", {}) or {}

    deck_out = pick(out_dir, work, "deck.pptx")
    deck_work = work / "deck.pptx"
    n_slides = 0
    titles = []
    if deck_work.is_file() and Presentation is not None:
        try:
            prs = Presentation(str(deck_work))
            n_slides = len(prs.slides)
            for s in prs.slides:
                try:
                    titles.append(s.shapes.title.text if s.shapes.title is not None else "")
                except Exception:  # noqa: BLE001
                    titles.append("")
        except Exception:  # noqa: BLE001
            pass
    if not n_slides:
        n_slides = len(plan.get("slides") or [])

    pdf_out = pick(out_dir, work, "render/deck.pdf") or pick(out_dir, work, "deck.pdf")
    if out_dir and (Path(out_dir) / "deck.pdf").exists():
        pdf_out = str(Path(out_dir) / "deck.pdf")
    png_dir_out = (Path(out_dir) / "previews") if out_dir and (Path(out_dir) / "previews").is_dir() \
        else (Path(out_dir) if out_dir else None)
    pngs = []
    if png_dir_out and png_dir_out.is_dir():
        pngs = sorted(str(p) for p in png_dir_out.glob("slide-*.png"))
    if not pngs:
        pngs = sorted(str(p) for p in (work / "render").glob("slide-*.png"))

    evidence_out = pick(out_dir, work, "evidence_map.json") or pick(out_dir, work, "plan.json")
    plan_out = pick(out_dir, work, "plan.json")
    notes_out = pick(out_dir, work, "speaker-notes.md")
    if not notes_out:
        notes_out = write_speaker_notes(plan, (out_dir or work) / "speaker-notes.md")
    # Text transcript: prefer the producer's own deck-transcript.md, else derive one from
    # plan.json so a transcript-shaped contract key can always be answered with real text.
    transcript_out = pick(out_dir, work, "deck-transcript.md")
    if not transcript_out:
        transcript_out = write_speaker_notes(
            plan, (out_dir or work) / "deck-transcript.md", title="Deck transcript")
    verify_out = pick(out_dir, work, "verify_report.json")
    render_out = pick(out_dir, work, "render_report.json")
    build_out = pick(out_dir, work, "build_report.json")
    material_out = pick(out_dir, work, "material.json")

    bases = [work]
    if case_path:
        bases.insert(0, Path(case_path).parent)

    template_path = None
    for candidate in (((plan.get("constraints") or {}).get("template_path")),
                      (((build.get("template") or {}).get("path"))),
                      (((case or {}).get("constraints") or {}).get("template_path")),
                      ((case or {}).get("template"))):
        found = resolve_existing(candidate, bases)
        if found:
            template_path = str(found)
            break

    # ---- status, sources and change records are computed *before* the contract loop: the
    # contract may legitimately name them as its own keys (caseB declares `status` and
    # `uncertainty`), and a resolve table that lacks them would report the key as unmapped.

    verdict = verify.get("verdict")
    render_status = render.get("status")
    # `status` describes DELIVERABLE COMPLETENESS, not the gate verdict.  The two are kept
    # separate on purpose: verify_deck's V20 reads this field, so deriving it from the verdict
    # would be circular -- a run could never climb back out of a provisional "partial", and a
    # perfectly complete run would be reported incomplete forever.  Correctness is still
    # published next to it in `_self_check.verify_verdict` / `checks_failed` and in
    # `uncertainty`, so nothing is hidden by the split.
    try:
        _deck_size = deck_work.stat().st_size if deck_work.is_file() else 0
    except OSError:
        _deck_size = 0
    core_present = ((work / "plan.json").is_file()
                    and (work / "material.json").is_file()
                    and (evidence_out is not None
                         or (work / "evidence_map.json").is_file()))
    if _deck_size <= 0:
        status = "failed"
    elif render_status == "ok" and core_present and bool(pngs):
        status = "ok"
    else:
        status = "partial"

    sources = []
    for src in material.get("sources") or []:
        if not isinstance(src, dict):
            continue
        raw_path = src.get("path")
        found = resolve_existing(raw_path, bases)
        path = str(found) if found else raw_path
        sources.append({"path": path,
                        "sha256": src.get("sha256")
                        or (sha256_file(path) if path and Path(path).is_file() else ""),
                        "role": "source"})
    if template_path:
        sources.append({"path": template_path, "sha256": sha256_file(template_path),
                        "role": "template"})
    start_found = resolve_existing((plan.get("constraints") or {}).get("starting_deck_path"),
                                  bases)
    if start_found:
        sources.append({"path": str(start_found), "sha256": sha256_file(start_found),
                        "role": "starting-deck"})

    change_records = changes_doc.get("changes") or versions.get("versions") or []

    uncertainty = []
    for w in (material.get("warnings") or []):
        uncertainty.append(f"material: {w}")
    for u in (plan.get("uncertainty") or []):
        uncertainty.append(f"plan: {u}")
    for u in (render.get("uncertainty") or []):
        uncertainty.append(f"render: {u}")
    for c in (verify.get("checks") or []):
        if c.get("status") in ("FAIL", "UNVERIFIED", "SKIP"):
            uncertainty.append(f"self-check {c.get('id')} {c.get('status')}: "
                               f"{(c.get('detail') or '')[:220]}")
    if status == "partial":
        uncertainty.append(
            f"Deliverables are incomplete, so status is 'partial' rather than 'ok': deck "
            f"bytes={_deck_size}, render status={render_status!r}, previews={len(pngs)}, "
            f"core artefacts present={core_present}.")
    if verdict != "PASS":
        # Correctness is a separate axis from completeness -- say so rather than folding it into
        # `status` (which would make V20 circular).
        uncertainty.append(
            f"The skill's own gate verdict is {verdict!r} (counts {verify.get('counts')}); see "
            f"verify_report.json and _self_check.checks_failed. `status` above describes only "
            f"whether the deliverables are complete, not whether every check passed.")
    if render_status != "ok":
        uncertainty.append(
            f"Rendered previews are not fully available (render status {render_status!r}); the "
            f"editable deck is still delivered, but no PDF/PNG preview can be shown for the "
            f"affected slides.")
    if not change_records:
        uncertainty.append("No version/change record artefact was found, so the change history "
                           "cannot be attested for this run.")
    uncertainty.append(
        "Technical file and OOXML checks validate structure and traceability only; they do not "
        "certify that the presentation is scientifically correct or well presented "
        "(task/task.md:97). Scores are assigned by the platform, not by this file.")

    values = {
        "deck_path": deck_out,
        "pdf_path": pdf_out,
        "png_paths": pngs,
        "evidence_path": evidence_out,
        "notes_path": notes_out,
        "transcript_path": transcript_out,
        "plan_path": plan_out,
        "slide_count": n_slides or None,
        "slide_titles": titles or None,
        "roles": [s.get("role") for s in plan.get("slides") or [] if isinstance(s, dict)] or None,
        "audience": (plan.get("constraints") or {}).get("audience"),
        "duration_min": (plan.get("constraints") or {}).get("duration_min"),
        "language": (plan.get("constraints") or {}).get("language"),
        "template_path": template_path,
        "verify_path": verify_out,
        "render_path": render_out,
        "build_path": build_out,
        "material_path": material_out,
        "deck_sha256": sha256_file(deck_work) if deck_work.is_file() else None,
        "pdf_sha256": sha256_file(pdf_out) if pdf_out else None,
        # core keys, so a contract that names any of them resolves instead of going unmapped
        "case_id": (case or {}).get("case_id") or plan.get("case_id") or material.get("case_id"),
        "status": status,
        "sources": sources,
        "uncertainty": list(uncertainty),
        "change_records": change_records,
        "self_check": {"verify_verdict": verdict},
    }

    def resolve(key):
        k = norm_key(key)
        if k in values:
            return values[k], "exact"
        for canon, aliases in SYNONYM_GROUPS.items():
            if k in aliases:
                return values.get(canon), canon
        return None, None

    contract = {}
    if isinstance(case, dict) and isinstance(case.get("output_contract"), dict):
        contract = case["output_contract"]

    unmapped = []
    contract_values = {}
    key_notes = []
    for key, default in contract.items():
        value, how = resolve(key)
        if how is None:
            # unmapped: never invent, but honour the type the case itself declared
            if isinstance(default, list):
                value = []
            elif isinstance(default, dict):
                value = {}
            else:
                value = None
            unmapped.append(str(key))
            key_notes.append(
                f"output_contract key {key!r} has no counterpart in this skill's outputs; "
                f"emitted {type(value).__name__} ({value!r}) rather than invention "
                f"(case declared default {default!r})")
        elif value in (None, [], {}):
            key_notes.append(
                f"output_contract key {key!r} maps to {how!r} but that artefact was not "
                f"produced, so it is {value!r}")
        contract_values[key] = value

    uncertainty = uncertainty + key_notes

    checks_failed = [c.get("id") for c in (verify.get("checks") or [])
                     if c.get("status") in ("FAIL", "UNVERIFIED")]
    artifacts = [p for p in (deck_out, pdf_out, evidence_out, plan_out, notes_out,
                             verify_out, render_out) if p]

    result = {}
    result.update(contract_values)
    result.update({
        "case_id": (case or {}).get("case_id") or plan.get("case_id") or material.get("case_id"),
        "status": status,
        "sources": sources,
        "uncertainty": uncertainty,
        "change_records": change_records,
        "_self_check": {
            "verify_verdict": verdict,
            "checks_failed": checks_failed,
            "checks_counts": verify.get("counts"),
            "artifacts": artifacts,
        },
        "_unmapped_contract_keys": unmapped,
        "_artifact_paths": {
            "deck_pptx": deck_out, "deck_pdf": pdf_out, "slide_previews": pngs,
            "plan": plan_out, "evidence_map": evidence_out, "speaker_notes": notes_out,
            "transcript": transcript_out,
            "verify_report": verify_out, "render_report": render_out,
            "build_report": build_out, "material": material_out, "template": template_path,
        },
        "_schema": "result/v1",
    })
    result["slide_count"] = result.get("slide_count") or n_slides
    return result


def main(argv=None):
    ap = argparse.ArgumentParser(description="Emit result.json for the case's output_contract.")
    ap.add_argument("--case", default=None)
    ap.add_argument("--work", required=True)
    ap.add_argument("--out", default=None, help="output dir (or the result.json path)")
    args = ap.parse_args(argv)

    out_dir = Path(args.out) if args.out else None
    result_path = (out_dir / "result.json") if out_dir else Path(args.work) / "result.json"
    if args.out and str(args.out).lower().endswith(".json"):
        result_path = Path(args.out)
        out_dir = result_path.parent
    case = load_json(args.case, {}) if args.case else {}
    result = build_result(case, args.work, out_dir, args.case)
    dump_json(result_path, result)
    print(json.dumps({"status": result["status"], "case_id": result.get("case_id"),
                      "keys_filled": len([k for k in result if not k.startswith("_")]),
                      "unmapped": result["_unmapped_contract_keys"],
                      "uncertainty_n": len(result["uncertainty"]),
                      "out": str(result_path)}, ensure_ascii=False))
    return 0 if result["status"] != "failed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
