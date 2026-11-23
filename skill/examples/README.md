# `examples/` — one complete development example (NOT an official case result)

This directory shows the interface and the whole usage flow with **real artifacts from one
actual pipeline run**, so a reader does not have to read the scripts to see what the Skill
produces.

**Everything here is a development example.** The input is one of our **synthetic** fixtures
(`work/fixtures/caseA`, experimental paper, every number invented). It is **not** one of the
3 official hidden cases, and **none of it is an official platform score**. The platform
scores the Skill itself on its own cases.

## Contents

| file | what it is |
|---|---|
| `dev-run-caseA/case.json` | the **input**: the platform shape `{case_id, note, output_contract:{...}}` plus the case's declared `constraints`; note the spelling of the contract keys (`deck_pptx`/`deck_pdf`/`slide_previews`/`speaker_notes`/`page_evidence_map`/`plan`) |
| `dev-run-caseA/result.json` | the **real** returned document: every declared `output_contract` key plus `case_id`, `status`, `sources`, `uncertainty`, `change_records`, `_self_check` |
| `dev-run-caseA/plan.json` | real page-level plan (role, claim, bullets, evidence locators, speaker notes) |
| `dev-run-caseA/evidence_map.json` | real slide → figure/table/section/number map with verbatim source locators |
| `dev-run-caseA/verify_report.json` | real gate report: V01–V19, zero FAIL, zero UNVERIFIED |
| `dev-run-caseA/build_report.excerpt.json` | key fields of the real build report (template invariants, fonts, per-slide layout/overflow/notes) |
| `dev-run-caseA/render_report.excerpt.json` | key fields of the real render report (renderer, PDF, per-slide preview geometry) |
| `dev-run-caseA/run.log` | real stage/timing events of that run |
| `dev-run-caseA/contact-sheet.png` | thumbnail contact sheet of the 8 rendered slides — the quickest way to see the deck |

## Invocation

```bash
# the platform's invocation
python /skills/weekly-skill/scripts/run.py \
    --case /input/case.json --work /app/work --out /app/output

# the same thing locally, on the synthetic fixture this example came from
python scripts/run.py \
    --case ../fixtures/caseA/case.json \
    --work ../fixtures/_example_work --out ../fixtures/_example_out
```

Individual stages run standalone and each writes one JSON artifact (`ingest.py`, `plan.py`,
`build_deck.py`, `render.py`, `evidence.py`, `verify_deck.py`, `result_contract.py`); see
`references/runbook.md`. `verify_deck.py` exits `0` only when its verdict is `PASS`.

## Two honesty notes a reviewer should read

1. **The paths in these files were sanitized for shipping.** The run happened on the
   author's machine, whose absolute paths must not appear in the submission tree, so they
   were rewritten to portable tokens (`work/…`, `output/…`, `input/…`). The values, ordering
   and structure are the run's own output, unmodified; only path strings changed, and every
   JSON was re-parsed after rewriting.
2. **`contact-sheet.png` is a downscaled thumbnail grid**, not the deliverable previews. The
   real run produced one full-resolution `slide-NN.png` per slide and a `deck.pdf`; the PNGs
   are downsampled here purely to respect the submission size budget.

## Reproducing this example

```bash
$env:PYTHONDONTWRITEBYTECODE='1'
cd ARMI2611000074v1/work
python fixtures/build_paper.py --json fixtures/fixture_build_report.json   # rebuild fixtures
python submission/scripts/run.py --case fixtures/caseA/case.json \
    --work fixtures/_example_work --out fixtures/_example_out
python fixtures/make_example.py --case fixtures/caseA/case.json \
    --work fixtures/_example_work --out fixtures/_example_out --case-name caseA
python submission/scripts/selftest.py --fixtures fixtures    # must exit 0
```

## What this example does not claim

* Not an official case result, not a benchmark score, not a whiteboard comparison.
* The synthetic paper's numbers are invented and internally consistent only inside the
  fixture; they are not measurements.
* The gate proves the invariants it names (traceability, template preservation, export
  integrity) and nothing about scientific or presentation quality.
