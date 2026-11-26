# Artifact schemas

Every stage of the pipeline writes one JSON artifact. This file is the human-readable
contract for those artifacts; `schemas/*.json` are the machine-readable JSON Schemas, and
`scripts/verify_deck.py` is what actually enforces the parts that matter at run time.

All artifacts carry `"schema": "<name>/v1"` and `"case_id"` so a reviewer can join them.

Pipeline (see `references/runbook.md` for invocation):

```
material.json → plan.json → deck.pptx + build_report.json
              → render_report.json (+ deck.pdf, slide-XX.png)
              → evidence_map.json → verify_report.json → result.json
```

---

## `material.json` — `material/v1`

The only artifact produced by reading the paper. **Nothing not present here may appear on a
slide**, which is why `warnings[]` matters: it records what was dropped and why.

| Field | Type | Meaning |
|---|---|---|
| `sources[]` | array | one per input file |
| `sources[].id` | str | handle used by every locator |
| `sources[].kind` | `pdf`\|`docx`\|`md`\|`txt` | how it was read |
| `sources[].path`, `.sha256`, `.bytes` | str/int | identity of the exact bytes read |
| `sources[].pages[]` | array | `{n, text}` 1-based |
| `sources[].sections[]` | array | `{id, title, page, char_start, char_end}` |
| `figures[]` | array | `{id, label, page, bbox, caption, locator, kind, data_table}` |
| `figures[].kind` | `raster`\|`vector`\|`unknown` | decides whether redraw is even possible |
| `tables[]` | array | `{id, label, page, caption, locator, cells, n_rows, n_cols}` |
| `numbers[]` | array | `{value, raw, unit, context, kind, locator}` — **`raw` is the verbatim string** |
| `claims[]` | array | `{text, locator}` |
| `warnings[]` | array | dropped/unresolved items **with reasons** |

Invariant: every `figures[]/tables[]/numbers[]/claims[]` entry has a `locator` containing a
verbatim `quote`. No quote → the item is absent and a `warnings[]` entry explains it.

## `plan.json` — `plan/v1`

The **versioned source of truth**. Deck quality changes here, not in the `.pptx`.

| Field | Meaning |
|---|---|
| `constraints` | audience, duration, `slide_count`/`slide_count_range`, language, template & starting-deck paths |
| `narrative.arc`, `narrative.main_thread[]` | the spine the pages implement |
| `slides[].index` | 1-based page order |
| `slides[].role` | `title`\|`outline`\|`problem`\|`method`\|`result`\|`figure`\|`table`\|`limitation`\|`conclusion`\|`ack`\|`backup` |
| `slides[].title`, `.claim` | the page's one-sentence claim |
| `slides[].bullets[]` | ≤5 items, each ≤120 chars |
| `slides[].figure_ref` / `.table_ref` | must resolve in `material.json` (gate V05) |
| `slides[].evidence[]` | `{type, ref, locator, note}` |
| `slides[].numbers_used[]` | `{value, raw, unit, locator}` — `raw` must be copied verbatim |
| `slides[].notes` | speaker notes, ≥40 chars (gate V03) |
| `slides[].layout_hint` | must be a layout that **exists** in the template |
| `excluded[]` | `{item, reason}` — what was deliberately left off |
| `uncertainty[]` | what could not be determined |

## `build_report.json` — `build_report/v1`

| Field | Meaning |
|---|---|
| `pptx`, `sha256` | the artifact built |
| `template` | path, sha256, masters, layouts, and the theme's major/minor latin+east-asian fonts |
| `invariants` | `master_unchanged`, `theme_unchanged`, `footer_preserved`, `layouts_added`, `layouts_deleted` |
| `slides[]` | `{index, layout, placeholders_filled, overflow_risk, notes_chars, images[]}` |
| `slides[].images[].mode` | `source-crop` \| `redraw-from-data` \| `reference-only` — nothing else |
| `fonts_used[]`, `warnings[]` | what was actually written, and anything suspicious |

`invariants` is the builder's *claim*; gates V07–V09 re-derive it from the package XML and
do not take this field's word for it.

## `render_report.json` — `render_report/v1`

| Field | Meaning |
|---|---|
| `renderer.name` | `libreoffice-kit` \| `soffice` \| `none` |
| `renderer.available`, `.version`, `.command` | what was actually executed |
| `pdf` | `{path, pages, bytes, sha256}` |
| `pngs[]` | `{slide, path, bytes, width, height}` — one per slide |
| `status` | `ok` \| `skipped` \| `failed` |
| `missingFonts[]` | font families the renderer had to substitute — **reported, never hidden** |
| `uncertainty[]` | including why rendering was skipped |

Hard rule: `status:"skipped"` with `available:false` is the honest outcome when no renderer
exists. No preview path may be claimed that was not produced.

## `evidence_map.json` — `evidence_map/v1`

| Field | Meaning |
|---|---|
| `pages[]` | one entry per slide: `slide`, `role`, `claims[]`, `figures[]`, `tables[]`, `numbers[]`, `unverified[]` |
| `claims[].match` | `exact` \| `paraphrase` — labelled because a paraphrase is weaker evidence |
| `claims[].verified` | true only when the quote really occurs in the source text |
| `coverage.slides_total` / `.slides_with_evidence` | gate V14 |

## `verify_report.json` — `verify_report/v1`

| Field | Meaning |
|---|---|
| `checks[]` | `{id, name, status, detail, evidence}` with `status ∈ {PASS, FAIL, SKIP, UNVERIFIED}` |
| `counts` | per-status totals |
| `verdict` | `PASS` iff zero `FAIL` **and** zero `UNVERIFIED` |
| `defects[]` | `{severity, check, detail, slide}` |

Gate ids are fixed and are referenced by `references/rubric-mapping.md`:

| id | Checks |
|---|---|
| V01 | slide count within the case constraint |
| V02 | every slide has a title and a claim |
| V03 | every slide has ≥40 chars of speaker notes |
| V04 | every slide number appears verbatim in `material.json` (re-derived, not trusted) |
| V05 | every `figure_ref`/`table_ref` resolves |
| V06 | every embedded image is `source-crop` or `redraw-from-data` (never a placeholder) |
| V07 | template masters/layouts unmodified (XML level) |
| V08 | theme fonts unchanged |
| V09 | footer / slide-number placeholders preserved |
| V10 | no text overflow |
| V11 | images inside their bounds (not cropped) |
| V12 | fonts resolvable; substitution reported |
| V13 | `result.json` contains every `output_contract` key |
| V14 | evidence map covers every slide |
| V15 | version records exist |
| V16 | no duplicate slide content (no shared titles, no ≥80% contained bullet text) |
| V17 | no constraint/input leakage into content (audience/duration/slide count/template path) |
| V18 | `build_report.json` invariants are internally consistent |
| V19 | minimum content coverage — a near-empty extraction cannot vacuously pass every `all()` |
| V20 | run produced deliverables — a non-empty readable deck in the output dir; `partial`/`failed` can never read as PASS |

The authoritative per-gate table (input, FAIL vs UNVERIFIED vs N/A, severity, failing example)
is `references/verification_gates.md`.

`UNVERIFIED` deliberately fails the gate: a check that could not inspect its target must not
be able to report success.

## `result.json`

Contract-driven. Fills **every** key declared in `case.json.output_contract` (tolerant
synonym mapping, e.g. `deck_pptx|pptx|presentation` → the `.pptx` path), and always adds:

| Field | Meaning |
|---|---|
| `case_id` | echoed from the case |
| `status` | `ok` \| `partial` \| `failed` |
| `sources[]` | `{path, sha256, role}` for every input actually read |
| `uncertainty[]` | everything not established, in plain language |
| `change_records` | what changed between rounds and why |
| `_self_check` | `{verify_verdict, checks_failed, artifacts[]}` |
| `_unmapped_contract_keys[]` | declared keys with no artifact — present and explained, never dropped |

A declared key that cannot be delivered is emitted as `null` with an `uncertainty` entry.
Silently omitting it is a defect (gate V13).
