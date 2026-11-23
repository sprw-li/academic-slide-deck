---
name: academic-slide-deck
description: >-
  Turn a paper or research material into an editable, template-conformant academic
  presentation deck with per-page evidence mapping: read /input/case.json, extract the
  research question, method, results and limits into a page-level plan (one claim per
  page, figures/tables referenced, speaker notes), build a real .pptx from the supplied
  master/layout template without touching its master, theme, fonts or footer, insert
  figures as source crops or data-driven redraws (never placeholders), preserve verbatim
  provenance for every number, render real PDF/PNG previews, run gates V01-V15, and write
  result.json with honest status, uncertainty and change records. Use when the user asks
  for 学术PPT, 组会汇报, 开题/中期/答辩 slides, 项目评审 or 基金申请 slides from a paper,
  a 论文转PPT task, a slide deck that must follow a .pptx template, a presentation with
  per-page 证据映射 and 演讲备注, or an editable deck plus export that must be reviewable
  and reproducible. Do NOT use for writing, translating or polishing paper prose, for
  literature review, for designing the statistics or the figures' analysis, for authoring
  a figure from raw data alone (that is a plotting task), or for developing presentation
  software. This Skill owns only page selection and ordering, template slot filling,
  figure/table binding, speaker notes, editability and export quality.
---

# 学术 PPT 生成与编辑 — page planning with retained evidence mapping

Turn **one paper / research material** into a deck that a researcher could actually
present: pages that carry the argument, numbers that match the source, a template that
survives, and an export you can inspect. The hard part is **取舍与重排** — which result
earns its own page, how far the method can be compressed, whether Figure 3 or Figure 4
comes first — not copying paragraphs into text boxes.

## 0. The one design rule

> **The LLM decides the science and the narrative; the scripts do every measurement,
> every write into the file, and every check. Do not blur the two.**

| Work | Who | Where |
|---|---|---|
| Which pages exist, what each page claims, what is *cut* and why, which figure/table backs which page, what the speaker says | **You (the agent)** | authored into `plan.json` |
| Reading the PDF/DOCX, locating figures/tables/numbers with a verbatim quote, opening the template, filling placeholders, cropping figures, rendering, gates V01–V15, `result.json` | **Deterministic scripts** | `scripts/*.py`, never hand-edited |

Corollary: **never hand-edit `deck.pptx`.** Change `plan.json` (or apply an edit
instruction) and rebuild. That is what makes a round reversible.

## 1. Input and output contract

Input is whatever the case provides — always read it, never assume it from labels:

- `/input/case.json` — declares `case_id` and an **`output_contract`**: a dict whose
  keys are the result fields the platform wants. **The key names differ between cases.**
  The Skill is contract-driven: it fills everything the case declares (using tolerant
  synonym mapping) and additionally always emits `case_id`, `status`, `sources`,
  `uncertainty`, `change_records`, `_self_check`. A declared key that cannot be delivered
  is emitted as `null` with an `uncertainty` entry and listed in `_unmapped_contract_keys`
  — it is **never silently dropped**.
- `/input/task.md` — the actual instruction, audience, time budget and slide count.
- Material files: the paper (PDF, sometimes DOCX/MD), a **master/layout-based `.pptx`
  template**, and for edit/adapt cases an **actual starting deck**.
- The evaluator installs this Skill at `/skills/weekly-skill`, runs it under
  `/app/work`, and reads final artifacts from `/app/output`.

Artifacts, in `/app/output`:

| Artifact | What it is |
|---|---|
| `deck.pptx` | the editable deliverable, built **on the supplied template** |
| `deck.pdf`, `slide-XX.png` | real rendered previews (honestly absent if no renderer) |
| `plan.json` | page-level plan: role, claim, bullets, figure/table refs, notes |
| `evidence_map.json` | per-page mapping of every claim/number/figure → source locator |
| `verify_report.json` | gates V01–V15 with counts, verdict and defects |
| `result.json` | every declared `output_contract` key + honest status |
| `run.log`, `versions/` | per-stage timings, change records, diff, rollback |

## 2. Quickstart — one command does the work

```bash
python scripts/run.py --case /input/case.json --work /app/work --out /app/output
```

`run.py` runs the deterministic chain and writes partial deliverables as it goes, so a
failure still leaves an honest `status:"partial"` `result.json` plus whatever succeeded:

```
ingest.py   -> material.json      (pages, sections, figures+captions, tables, numbers with verbatim quotes)
plan.py     -> plan.json          (narrative arc, one claim per page, notes, layout choice)
build_deck  -> deck.pptx          (template-bound; master/theme/footer untouched)
render.py   -> deck.pdf + PNGs    (real previews, or an honest skip)
evidence.py -> evidence_map.json  (per-page provenance, quote-verified)
verify_deck -> verify_report.json (V01-V15; exit 0 iff verdict PASS)
result_contract -> result.json    (contract-driven)
```

Every stage is separately runnable (`--case/--work` plus optional `--in/--out`), which is
what makes the pipeline debuggable and cheap to re-run inside the platform budget.

## 3. How to plan the pages (the part only you can do)

1. **Read the material first.** `ingest.py` output is the ground truth; if a number is not
   in `material.json.numbers`, it does not go on a slide.
2. **Pick the spine, then cut.** Fix the narrative: problem → gap → method → key result(s)
   → validation → limitation → conclusion. Then decide what the audience must be able to
   repeat afterwards, and delete everything else. Record the deletions in
   `plan.json.excluded` with a reason — the reviewer wants to see the tradeoff, not the
   leftovers.
3. **One page, one claim.** A page whose claim cannot be said in one sentence is two pages.
   ≤5 bullets, each ≤120 characters. A paragraph pasted into a text box is a defect.
4. **Give the decisive result its own page**, with the figure and the number on it.
5. **Respect the template's real layouts.** Choose `layout_hint` values that exist in the
   supplied template (the builder reads them for you); never invent a layout name.
6. **Write the speaker notes as what you would actually say**, ≥40 characters per slide.
   Notes are a deliverable, not an afterthought.
7. **Slide count is a constraint, not a suggestion.** Honour the case's count/range.

## 4. Evidence mapping is not optional

Every page records, in `evidence_map.json`, each claim / number / figure / table **with its
source locator and a verbatim quote**. The rule is one-way:

> If you cannot produce a verbatim `quote` for it, it does not ship.

Consequences that reviewers check first, because they are the recorded failure modes of
existing agents: numbers copied wrong or **rounded by impression**, captions that do not
match the figure, a placeholder or invented curve where a figure belongs, a template's
fonts/colours/footer overwritten, a master broken while editing one page, text overflowing
after export, images cropped, and no way back to the previous version.

## 5. Figures: what may be redrawn and what may only be referenced

| Situation | Mode | Rule |
|---|---|---|
| The numbers are in a real data table in the material | `redraw-from-data` | Redraw with the same values; the numbers on the slide must equal the table's |
| Only a picture of the figure exists | `source-crop` | Crop the real figure region from the source; never re-draw what you cannot read |
| The figure is referenced but the material has no usable image | `reference-only` | Cite it in text (「见论文 Figure 3」); do **not** insert a placeholder or an invented curve |
| Nothing usable at all | — | Say so in `uncertainty`; a missing figure is reported, never faked |

## 6. Template discipline

- Open the supplied `.pptx`; **fill placeholders on existing layouts**. The builder asserts
  and reports `master_unchanged`, `theme_unchanged`, `footer_preserved`,
  `layouts_added=0`, `layouts_deleted=0`.
- Never restyle the master, never overwrite theme fonts/colours, never fill the footer or
  slide-number placeholder.
- Fonts: honour whatever the template names (organizer templates in this family name
  **Liberation Serif**, OFL-1.1). If a font cannot be resolved, report the substitution —
  never hide it.
- If a layout has no usable content placeholder, place text inside the content
  placeholder's rectangle rather than inventing a new slide geometry.

## 7. Iterating without losing the thread

Multi-round editing is a first-class requirement, and its failure mode is
「说不清哪一版是交付版」.

- Each run writes a version record under `versions/` with: what changed, **why**, and the
  sha256 of the artifact.
- A bad round must be **revertible**: `scripts/run.py` re-run from an earlier `plan.json`
  reproduces the earlier deck deterministically — the plan is the versioned source of
  truth, the `.pptx` is a build product.
- Never edit the `.pptx` in place; never keep two "final" files.

## 8. Self-check (run this before believing anything)

```bash
python scripts/verify_deck.py --case /input/case.json --work /app/work   # exit 0 iff verdict PASS
python scripts/selftest.py                                               # end-to-end + injected-defect tests
```

`verify_report.json` carries V01–V19 (V16 duplicate content, V17 constraint leakage into
content, V18 build-report invariants, V19 minimum content coverage). Its semantics are strict
on purpose:

- `verdict = PASS` requires **zero `FAIL` and zero `UNVERIFIED`**; a `SKIP` is allowed only
  with a recorded reason.
- **A check that could not actually inspect the thing reports `UNVERIFIED`, not `PASS`.**
  If the report is green, that proves exactly the invariants it names and nothing more —
  a green file check does **not** certify that the presentation is good science.
- The gate must be able to fail: `selftest.py` injects defects (fabricated number, missing
  notes, wrong slide count, modified master/theme, placeholder figure, missing
  `output_contract` key, overflowing text) and **exits non-zero if any is not detected**.

## 9. Failure handling — degrade honestly, never fabricate

| Situation | Behaviour |
|---|---|
| No PDF text layer (scanned) | Try OCR; if unavailable, `status:"partial"`, name the affected pages |
| Figure is a bitmap with no extractable data | `reference-only`, do not invent values |
| Template missing or unreadable | Build from a minimal internal layout, mark the deck "template not applied" in `uncertainty` |
| Font not installed | Use the closest resolved font, report the substitution and the missing family |
| Slide count cannot be met with the material | Deliver the closest honest count, record the constraint conflict |
| No renderer available | `render.status:"skipped"`, `available:false`, no preview paths claimed |
| Any stage fails | Write partial deliverables first, then `status:"partial"` with the limits |

## 10. Boundaries — what this Skill does not do

- It does not write, translate or polish the paper's prose, and does not do literature
  review.
- It does not design the study, the statistics, or the visualization algorithm.
- It does not author a figure from raw data as a standalone plotting job.
- It does not develop presentation software, nor guarantee a specific visual style beyond
  the supplied template.
- It does not certify scientific quality. Its gates cover structure, traceability,
  template invariants and export integrity — the science remains the author's.
- Supported inputs: PDF (text layer), DOCX, MD/TXT, and `.pptx` templates with a real
  master/layout set. Known limits and their detections are in
  `references/failure-modes.md`.

## 11. Reference files

| File | Read it for |
|---|---|
| `references/runbook.md` | full invocation, stage-by-stage, debugging a failing stage |
| `references/verification_gates.md` | **the authoritative V01–V19 table**: what each gate inspects, what makes it FAIL vs UNVERIFIED vs N/A, severity, failing example |
| `references/prior-art.md` | what existing tools already do, and the deliberately narrow boundary of this Skill's contribution |
| `references/spec-schema.md` | the exact JSON schema of every artifact |
| `references/provenance-and-evidence.md` | how locators and verbatim quotes are produced and verified |
| `references/exports-fonts-colour.md` | measured export/font/colour behaviour, renderer discovery |
| `references/failure-modes.md` | known failure modes, detection, and honest limits |
| `references/rubric-mapping.md` | rubric line → the artifact/check that addresses it |
| `references/measured-gain.md` | what was measured, on which fixture, and what is *not* claimed |
| `scripts/contract_matrix.py` | machine-checkable version of the rubric mapping |
