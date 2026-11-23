# Measured gain — what was measured, on what, and what is deliberately **not** claimed

Written to make over-claiming impossible. If a number is not in this file, it was not measured.

> **Nothing here is an official score.** Grading for this challenge is `human_review_only`
> (`challenge.json.scoring`), and the platform produces the skill/no-skill comparison on its own
> hidden cases against its own whiteboard. This repository has no access to that comparison and
> does not invent one.

---

## 1. What the efficiency/quality numbers are measured on

| | |
|---|---|
| Fixture | `caseA` — our own **SYNTHETIC** case (`SYNTHETIC / NOT AN OFFICIAL CASE`). A 4-page synthetic experimental paper with 4 numbered figures, 2 numbered tables and ~80 distinct numbers, plus a real master/layout `.pptx` template. |
| Harness | `scripts/gain_check.py`, re-runnable: `python scripts/gain_check.py --fixtures <dir> --case caseA --json <out>.json` |
| Repetitions | **one run.** The noise floor is therefore unknown, and a difference of one or two units is explicitly *not* evidence. |
| Baseline | a **deterministic proxy**, not a model run. Two constructions: `proxy:sec` = one slide per source section in source order, verbatim; `proxy:paras` = one slide per source claim/paragraph, verbatim — i.e. the naive failure mode the task description names. |
| `official_score` | `null` (by construction, in the JSON) |

## 2. The measured table

| Metric (lower is better) | with Skill | proxy:sec | proxy:paras | discriminates? |
|---|---|---|---|---|
| `pages_needing_reorder` | 0 | 0 | 0 | **no** |
| `export_recheck_rounds` | 1 | n/a | n/a | **no** (no baseline measurable from our artifacts) |
| `numbers_failing_traceability` | 0 | 0 | 0 | **no** |
| `template_violations` | 0 | n/a | n/a | **no** (the naive proxy is never run through the template, so it has no column) |
| **`slides_over_case_constraint`** | **0** | 0 | **31** | **yes — the only discriminating metric** |

With-Skill context for the same run: 8 slides, 0 over the bullet cap, 0 below the notes minimum,
0 without evidence, against a material of 4 figures / 2 tables / 39 claims / 9 ingest warnings.

### What this table actually supports

- **Supported:** the pipeline honours the case's slide-count constraint, where a verbatim
  paragraph-dump proxy overshoots it by 31 slides. That is a real, reproducible structural effect
  (0 vs 31 is far outside any plausible noise, and it is the constraint the rubric's 任务完成度
  names).
- **Not supported:** that page *quality* improved, that rework decreased, or that traceability
  improved — because on this fixture four of five metrics do not discriminate at all. They are
  reported for completeness and are labelled non-discriminating in the JSON
  (`comparison[].discriminating == false`) rather than presented as wins.

## 3. Evidence that is *not* a gain metric, but is real

These are properties demonstrated by running, and they do not depend on a baseline:

| Evidence | Where it comes from |
|---|---|
| The gate set is adversarially real | `selftest.py` / `negtest.py` inject defects; **18/18 detected**, and each is detected by the *expected* check id; a good fixture passes **19/19** and exits 0 |
| Number traceability checked twice by independent methods | gate **V04** re-derives slide numbers from `material.json`, and `gain_check.py` independently recounts untraceable numeric tokens — both report 0 on caseA, and `cross_checks` records the agreement rather than assuming it |
| Template preservation is observed, not asserted | `build_report.json.invariants` (builder's claim) is checked separately by **V07–V09** reading the produced package's XML: `master_unchanged`, `theme_unchanged`, `footer_preserved`, `layouts_added=0`, `layouts_deleted=0` |
| Rendering really happens | **measured** on this host: `convert` 4.3 s; `render.py` for 1 slide 15.2 s; the whole `run.py` on caseA **44.1 s** |
| End-to-end completion on two different case shapes | caseA (`deck_pptx/deck_pdf/slide_previews/speaker_notes/page_evidence_map/plan`) and caseB (`pptx/pdf/previews/notes/evidence/outline`) both exit 0 with `verdict: PASS` and `status: "ok"` — i.e. the contract-driven mapping survives two disjoint key vocabularies |
| A real reliability bug was found and fixed by measurement | the pdf/png converter can **hang**: its worker grandchild inherits stdout, so killing Node leaves the pipe open and a plain `subprocess.run(timeout=…)` never returns. Now: file-based capture + process-tree kill + `--budget`, proven by a timeout test (rc=124 in 7.8 s vs previously unbounded) |

## 4. What is explicitly NOT claimed

- **No official score, and no knowledge of it.** `official_score` is `null`.
- **No measured base-Agent comparison.** The baseline columns are deterministic proxies; a real
  no-skill arm needs a second agent harness and the platform's whiteboard, neither of which exists
  on this host. The JSON says this in `baseline.caveat`.
- **No quality-gain claim.** Quality is not measured here; it is inspected (rendered PNGs) and
  gated on structural invariants.
- **No font-fidelity claim from previews.** The renderer reports `missing_fonts: ["Liberation Serif"]`
  and the produced PDF embeds `TimesNewRomanPSMT` / `ArialMT` — the substitution is proven, so a
  preview is **not** a certificate of font fidelity (see `references/exports-fonts-colour.md` §4).
- **No novelty claim beyond an audit layer.** See `references/prior-art.md`.
- **No claim that a green report means good science.** The gates cover structure, traceability,
  template preservation and export integrity only.

## 5. How to reproduce

```bash
python scripts/gain_check.py --fixtures <fixtures-dir> --case caseA --json gain-caseA.json
python scripts/selftest.py --fixtures <fixtures-dir>
python scripts/verify_deck.py --case <case>.json --work <w> --out <o>
```

If a future run disagrees with this file, **the future run is the truth** and this file is stale —
say so rather than quietly reusing these numbers.
