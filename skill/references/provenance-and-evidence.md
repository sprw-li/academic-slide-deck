# Provenance, evidence and attribution

This file answers two different questions that are easy to confuse:

1. **Evidence integrity** — can a reviewer trace every fact on a slide back to the source,
   and can they tell when a claim is *not* traceable?
2. **Attribution** — where did the code, fonts, icons and other material in this Skill come
   from, and what obligations come with them?

Both are recorded because the rubric asks for 来源 (sources) and because misattributing a
licence is a real, cheap-to-trigger failure.

---

## Part 1 — Evidence integrity

### 1.1 The one-way rule

> **If ingest cannot produce a verbatim quote for a fact, that fact does not enter
> `material.json` — so it cannot reach a slide.**

This is enforced *upstream* on purpose. Filtering invented numbers out late is unreliable
(the wrong number may look plausible); refusing to admit them at all removes the class of
error instead of shrinking it.

Practical consequence: `material.json.warnings[]` is a **first-class output**, not noise. A
dropped figure, an unmatched number, a caption that could not be tied to a page all appear
there with the reason. A reviewer who disagrees with what was dropped can see exactly what
was dropped and re-run ingest.

### 1.2 What a locator contains

Every figure, table, number and claim carries a `locator`:

```jsonc
{ "source": "paper",                 // which input file, by id
  "page": 3,                          // 1-based page (PDF) or section index (docx/md)
  "section": "3.2 Dose response",     // when determinable
  "quote": "…IC50 of 12.3 μM (95% CI 9.8-15.1)…",
  "char_start": 4120, "char_end": 4153 }
```

The `quote` is the load-bearing field: `evidence.py` re-reads the source text and marks an
entry `verified: true` only when the quote really occurs there. A locator with a page number
and no quote is not evidence.

### 1.3 How the numbers are checked twice, independently

| Stage | Who checks | Against what |
|---|---|---|
| ingest | `ingest.py` | the source text it is reading |
| plan | `plan.py` | `material.json.numbers` (may only reuse `raw` strings verbatim) |
| gate | `verify_deck.py` **V04** | re-derives the slide numbers from **`material.json`**, not from `plan.json` |
| map | `evidence.py` | re-verifies each quote character-by-character against the source |

V04 deliberately does **not** trust the planner's own report. If the planner and the
material disagree, V04 fails — a planner that lies cannot hide behind its own summary.

### 1.4 The per-page evidence map

`evidence_map.json` is a **deliverable**, not a log:

```jsonc
{ "pages": [
    { "slide": 4, "role": "result",
      "claims":  [ { "claim": "…", "locator": {...}, "match": "exact", "verified": true } ],
      "numbers": [ { "raw": "12.3", "locator": {...} } ],
      "figures": [ { "ref": "fig2", "caption_locator": {...} } ],
      "unverified": [] } ],
  "coverage": { "slides_total": 10, "slides_with_evidence": 10 } }
```

Anything not verifiable lands in `unverified[]` and is reported in `result.json.uncertainty`.
An empty `unverified[]` plus `slides_with_evidence == slides_total` is the claim gate **V14**
checks — and it is a claim about *traceability*, never about scientific validity.

### 1.5 What evidence integrity does NOT establish

- It does not establish that the paper's own numbers are correct.
- It does not establish that the selected narrative is the best one, or even a good one.
- It does not establish that a paraphrase preserves meaning — a `match: "paraphrase"` entry
  is labelled as such precisely because it is weaker than a verbatim match.
- It does not establish that a figure cropped from the source is the *right* figure for the
  claim; that is a human judgement.

## Part 2 — Attribution and licensing

### 2.1 This Skill's own code

All scripts, references and schemas in this submission were written for this challenge.
No third-party source files are vendored into the shipped tree.

### 2.2 Regenerated development fixtures (not shipped as cases)

The development fixtures under `work/fixtures/` are **synthetic and self-generated**:
the template is produced by `build_template.py` from python-pptx's own default template
plus theme/stamp edits, and the papers are generated locally. They are **not** official
organizer cases, and each fixture's `case.json.note` and `task.md` say so.

### 2.3 Fonts

The Skill **refers to font families by name and never ships font binaries.**

- Organizer templates in this challenge family name **Liberation Serif**
  (Liberation fonts: SIL Open Font License 1.1 — OFL-1.1).
- Locally, the development fixtures reference the same family; the font files themselves are
  read from the platform image / the local material bank, never copied into the payload.
- If a template names a family the machine cannot resolve, the substitution is reported
  (`render_report.json` → `missingFonts`, plus `result.json.uncertainty`). Silent
  substitution is the failure mode this is designed to make visible.

> **If you extend this Skill with a bundled font, you must carry its licence and attribution
> into your own documentation.** OFL requires the licence text to accompany the font, and
> CC-BY/OFL icon sets require attribution. Do not assume the platform's font availability
> licenses redistribution.

### 2.4 Icons, palettes and other third-party material

None are bundled in this submission. If you add any:

| Source family | Obligation |
|---|---|
| OFL fonts (Liberation, Source Han, Inter, IBM Plex, academicons) | include the OFL text; do not sell the font by itself; rename if modified |
| CC-BY icons (Font Awesome icons, colorcet, bioicons in part) | **attribution required** in the delivered documentation |
| CC0 / MIT / BSD / Apache-2.0 | attribution appreciated, not required (keep the notice for MIT/BSD) |
| Copyleft tools (GPL/AGPL, e.g. some PDF/LaTeX converters) | *invoke* the tool; do **not** copy its source into the payload |
| No licence found in the upstream repo | **do not use** |

The local material bank used during development lives **outside** this payload, in the
contest workspace under `SKILL_Downloaded/` (a workspace-internal reference library with
per-entry download receipts and per-entry licence fields; its own index is in that
directory's `README.md` and `SKILL.md`). It is a **reference bank, not an environment**, and
it is not part of this submission. No path from it is referenced by any script here.

### 2.5 Credentials

No credentials, tokens or private keys appear anywhere in this submission — verified by
`work/verify/check.py` check **S05** (credential-like string scan over the shipped tree).
Development used a workspace-local CLI credential file that is outside the payload; no value
from it is reproduced in any artifact, log or report.

### 2.6 Honesty vocabulary used in the artifacts

Three status layers are kept separate on purpose, because collapsing them is how a green
report starts lying:

| Layer | Values | Meaning |
|---|---|---|
| check | `PASS` / `FAIL` / `SKIP` / `UNVERIFIED` / `N/A` | the outcome of one inspection |
| run | `ok` / `partial` / `failed` | how complete the deliverable is |
| claim | measured / not measured | whether a gain was actually observed, and on what |

`UNVERIFIED` is treated as a **gate failure**, so an inspection that could not happen cannot
masquerade as success. And a green gate certifies the invariants it names — structure,
traceability, template preservation, export integrity — not scientific quality.
