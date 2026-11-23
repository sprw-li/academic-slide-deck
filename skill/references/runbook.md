# Runbook — invoking the Skill, stage by stage, and debugging it when a stage fails

Companion to `SKILL.md` (what it is) and `references/spec-schema.md` (what each artifact
contains). This file is the operator's view: what to type, what should appear, and what to
do when it does not.

---

## 1. The normal path

```bash
python scripts/run.py --case /input/case.json --work /app/work --out /app/output
```

Expected on success (`exit 0`):

```
/app/work/material.json
/app/work/plan.json
/app/work/deck.pptx
/app/work/build_report.json
/app/work/render/deck.pdf
/app/work/render/slide-01.png ... slide-NN.png
/app/work/render_report.json
/app/work/evidence_map.json
/app/work/verify_report.json          <- verdict: PASS
/app/output/deck.pptx  deck.pdf  slide-*.png
/app/output/plan.json  evidence_map.json  verify_report.json  run.log
/app/output/result.json
/app/work/versions/**
```

`exit 0` means **the chain ran and the gates passed** — nothing more. Read
`verify_report.json.verdict` and `result.json.status` before believing anything.

Useful flags:

| Flag | Effect |
|---|---|
| `--case PATH` | the case file (default `/input/case.json`) |
| `--work DIR` | intermediate artifacts (default `/app/work`) |
| `--out DIR` | final deliverables (default `/app/output`) |
| `--skip-render` | skip PDF/PNG (use on a host with no renderer, or for a fast logic pass) |
| `--only STAGE[,STAGE]` | run a prefix/subset of the chain, for debugging |
| `--in PATH` / `--out PATH` | on individual stages: override the input or output JSON |

## 2. Stage by stage

Run these in order when you need to see where something went wrong. Each is independently
runnable and each writes one JSON artifact.

```bash
# 1. Read the material. The ONLY stage that touches the paper.
python scripts/ingest.py --case C --work W
#    -> W/material.json
#    Check: figures[]/tables[]/numbers[] each carry a locator with a verbatim quote.
#    Check: material.json.warnings[]; a dropped item is explained there, never silently gone.

# 2. Turn material into pages. This is the file a human edits across iterations.
python scripts/plan.py --case C --work W
#    -> W/plan.json
#    Check: slides[].role/claim/notes; <=5 bullets, each <=120 chars; layout_hint exists
#    in the REAL template; excluded[] explains what was cut; slide count within the case.

# 3. Build the deck on the template.
python scripts/build_deck.py --case C --work W
#    -> W/deck.pptx + W/build_report.json
#    Check: invariants.master_unchanged / theme_unchanged / footer_preserved all true,
#    layouts_added==0, layouts_deleted==0; slides[].images[].mode is source-crop or
#    redraw-from-data (never a placeholder).

# 4. Render previews.
python scripts/render.py --case C --work W
#    -> W/render/deck.pdf + W/render/slide-XX.png + W/render_report.json
#    Check: renderer.name != "none"; status "ok". A "skipped" here is honest, not a failure
#    of the run, but it means no preview exists to inspect -- say so.

# 5. Build the per-page evidence map.
python scripts/evidence.py --case C --work W
#    -> W/evidence_map.json
#    Check: coverage.slides_with_evidence == slides_total; every entry verified:true or
#    listed under unverified[].

# 6. Gate everything.
python scripts/verify_deck.py --case C --work W
#    -> W/verify_report.json ; exit 0 iff verdict PASS
#    Check: counts.FAIL == 0 AND counts.UNVERIFIED == 0. Read defects[] before anything else.

# 7. Emit the contract-driven result.
python scripts/result_contract.py --case C --work W --out O
#    -> O/result.json
#    Check: every key in case.json.output_contract is present;
#    _unmapped_contract_keys[] is empty or explained in uncertainty[].
```

## 3. What the exit codes mean

| Exit | Meaning |
|---|---|
| `0` | the stage succeeded / the gate passed |
| non-zero | that stage failed; read its JSON artifact and `run.log` first, then the traceback |

`verify_deck.py` is the only stage whose exit code is a **verdict**: `0` iff
`verdict == "PASS"`, which requires zero `FAIL` and zero `UNVERIFIED`. `SKIP` is permitted
only with a recorded reason. A stage that crashes exits non-zero *and* still leaves the
artifacts produced before the crash in place.

## 4. Debugging recipes

**A number on a slide looks wrong / a figure caption mismatches.**
Dump the slide's numbers from `plan.json` and compare against `material.json.numbers`
(`raw` must match exactly). Run `verify_deck.py` and read the `V04` defect: it reports the
slide, the offending raw string, and whether the value exists in the source at all. If the
number is genuinely in the paper but ingest missed it, that is an *ingest* problem (the
quote was not matched) — fix `material.json` by re-running ingest, do not edit the number
into the plan.

**The deck looks wrong but the gates are green.**
Expected: the gates cover structure, traceability, template invariants and export
integrity — not aesthetics. Inspect `slide-XX.png`. If the problem is layout/overflow,
`verify_report.json` V10/V11 should have caught it on geometry; if it did not, that is a
gate defect worth reporting rather than a reason to stop trusting the gate broadly.

**The template's fonts/colours/footer changed.**
`build_report.json.invariants` plus gates V07–V09 compare the produced package's XML
against the template's. If the deck built from a *starting deck* (edit/adapt case), compare
against that starting deck instead — V07–V09 take the reference from the case input.

**Rendering is skipped.**
`render_report.json.renderer.available == false`. Discovery order and the environment
variables to set are in `references/exports-fonts-colour.md`. On a host with no renderer at
all this is the correct, honest outcome: the run is still valid, but no preview exists, so
`result.json` must say so and the reviewer must be told.

**A stage died mid-way.**
`run.py` is partial-safe by design: it writes what succeeded, then `result.json` with
`status:"partial"` and the limits. Re-run just the failing stage with `--only` or directly,
then re-run the tail.

**The gate is `UNVERIFIED`, not `FAIL`.**
`UNVERIFIED` means the check could not actually inspect its target — e.g. no render to
cross-check against, or the template could not be re-opened. It is deliberately treated as
failing the gate. Fix the pre-condition (usually: make the run render, or point `--case` at
a readable template) rather than editing the report.

## 5. Iterating on content without losing the version

The plan is the source of truth; the deck is a build product.

1. Edit `plan.json` (or regenerate it with a different emphasis) — **never** edit `deck.pptx`.
2. Re-run `build_deck.py` then the tail.
3. A `versions/` change record captures what changed, why, and the artifact sha256.
4. To go back a round: restore the earlier `plan.json` and rebuild. The output is
   byte-reproducible from the plan plus the template, which is what makes rollback safe.

## 6. Self-test before you trust the pipeline

```bash
python scripts/selftest.py
python scripts/contract_matrix.py    # rubric line -> real path + real check id
python scripts/gain_check.py         # measured rework metrics on OUR OWN synthetic fixture
```

`selftest.py` runs the chain end-to-end **and** injects defects (fabricated number, missing
notes, wrong slide count, modified master/theme, placeholder figure, missing
`output_contract` key, overflowing text). It exits non-zero if any injected defect is *not*
detected — a self-test that only prints "OK" would be worse than none, because it would
manufacture false confidence.

## 7. Where the development fixtures live (not part of the submission)

```
work/fixtures/
  build_template.py            rebuilds the SYNTHETIC organizer template (exit 0 prints RESULT: OK)
  template-organizer.pptx      11 layouts; theme fonts = Liberation Serif; accent1 = 1F6FEB;
                               master stamp textbox "TemplateStamp"; layout footer placeholders
  caseA/ caseB/                SYNTHETIC cases (different paper types, different
                               output_contract key spellings) — for exercising the mapping
```

These are **self-made development fixtures, not official cases**, and they are labelled as
such in `case.json.note` and in `work/fixtures/README.md`. They are deliberately built so
that "the template was preserved" is observable: a distinctive theme accent colour, theme
name, and an `ORGANIZER-ACADEMIC-2026` master stamp.
