# Failure modes and what this skill does not certify

This is the honest limitations ledger. Every "detected" claim below was produced by injecting
the defect and re-running the gate (`work/tmp-verify-render/negtest.py`, 18 injected defects,
**18/18 detected by the expected check, all with non-zero exit**, 2026-11-23). Where a row is
reasoned rather than measured, it says so.

The rule that governs all of this: **a check never reports PASS on something it did not
inspect.** When a check cannot read its evidence it reports `UNVERIFIED`, which fails the gate
by design (`CONTRACT.md:196`). Do not "fix" a red gate by relaxing it.

Gate-by-gate semantics — what each of V01–V19 reads, when it may pass, the N/A / WARN /
BLOCKING distinction, and V19's measured threshold — live in
[`verification_gates.md`](verification_gates.md). This file is the *failure-mode* ledger.

---

## A. Rendering and export

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| A1 | No renderer reachable | discovery returns nothing → `render_report.status:"skipped"`, `renderer.available:false` | writes the deck and a matching `uncertainty` entry; **claims no preview path**; `result.json.status` degrades to `partial` | that the deck looks right — nothing was rendered |
| A2 | `convert` refuses to overwrite | measured `EEXIST … open 'deck.pdf'` on the 2nd run | unlinks the target before every attempt | — |
| A3 | `render` refuses an existing `--output-dir` | measured `EEXIST … mkdir` | renders into a scratch dir it clears, then normalises to `slide-XX.png` and deletes the scratch | — |
| A4 | Uniformly blank/transparent preview (a renderer failure, not an empty slide) | PNG pixel probe: `nonbackground_ratio < 5e-4` → `blank_slides_suspected` + uncertainty | flags the slides; if *every* preview is uniform it says to treat the whole render as failed and rely on geometry checks | nothing about *why* it is blank |
| A5 | Rendered page count ≠ slide count | `len(pngs) != len(pptx.slides)` → uncertainty entry | records both counts | that the pages map 1:1 onto slides |
| A6 | Renderer reports missing fonts | `missingFonts` copied into `render_report.json` (+ uncertainty) | surfaces the list; V12 fails if a substituted face is reported nowhere | that the substituted glyphs match the real font |
| A7 | Only PDF produced, no PNG (or vice versa) | `status:"partial"` | keeps the artefact that exists, records the gap | the missing preview |
| A8 | Render media/animation fidelity | not detectable from a static export | — | **any** claim about animation, transitions, embedded media or speaker-note audio |
| A9 | **The renderer hangs and a naive subprocess timeout does not return** | measured: a `render` attempt stalled for **>10 min** with an idle `node` worker; the mechanism was then reproduced (the CLI spawns a worker **grandchild** that inherits stdout, so killing Node alone leaves the pipe open and `communicate()` blocks forever) | stdout/stderr go to files (not pipes), the whole **process tree** is killed on expiry (`taskkill /F /T`, `killpg`), and a total `--budget` (default 300 s) caps all attempts; the hang is reported as `exit 124` + an `error` string, never as a fake preview | that the same stall cannot happen for a *different* reason — but it can no longer block the run, and `timeout_test.py` proves the timeout returns (7.8 s against a 5 s deadline) |
| A10 | Concurrent LibreOffice use | observed once (a chain measuring 9.5 s + 2.2 s stalled >600 s while other agents ran LibreOffice) | the budget above bounds the damage | the correlation is observed, not proven — run the render stage **serially**, never in parallel with another LibreOffice job |

## B. Geometry and layout

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| B1 | Text overflows a shape with **no autofit** (text is clipped) | V10: estimated extent > box, from deck geometry using real font advances | **FAIL**, with the shape, slide and both extents | the exact rendered pixel — it is an estimate (see B4) |
| B2 | `spAutoFit` shape grows off the slide | V10: grown extent vs the **slide** box | **FAIL** ("runs off the slide") | whether it merely collides with a neighbour instead |
| B3 | `normAutofit` shrink-to-fit | V10 | passes, but the shape is listed in the detail as relying on shrink-to-fit | that the shrunk text is still legible — nothing measures font size on screen |
| B4 | Estimate vs reality | V10 additionally cross-checks the rendered PDF with `pdfplumber` (no word may fall outside the page box) | reports the cross-check result in the same detail string | that the renderer's own text metrics equal PowerPoint's |
| B5 | Cropped image (`<a:srcRect>` non-zero) | V11 reads the picture's OOXML, namespace-agnostically (`<p:blipFill>` carries `<a:srcRect>`) | **FAIL** with the crop values | that an uncropped image is *the right* image (that is D3/E1) |
| B6 | Picture frame outside the slide box | V11 compares frame geometry to the slide size | **FAIL** | overlap between two in-bounds shapes — deliberately not attempted |
| B7 | Deck with zero picture shapes | V11 | passes **vacuously** and says so: "0 images … proves nothing about figures that were not embedded" | anything about figures that should have been embedded |
| B8 | Unreadable deck / no `python-pptx` | every geometry check | `UNVERIFIED` (fails the gate) | nothing |

## C. Fonts and colour

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| C1 | Font named by the deck is not installed | V12 resolves each attributable family against platform font dirs + matplotlib faces | if it is reported as substituted anywhere → **PASS** (substitution visible); if not → **FAIL** | that the substituted face is a good match (measured: Liberation Serif → **Times New Roman**) |
| C2 | A substituted font is reported nowhere | V12 | **FAIL** "font `X` substituted silently" | — |
| C3 | Theme's supplemental script-font table mistaken for "fonts used" | guarded: the theme's `<a:fontScheme>` fallback map and `+mj-lt`/`+mn-ea` refs are excluded | (this was a **false positive found and fixed** on the first run) | — |
| C4 | Preview treated as a font-fidelity certificate | — | documented in `exports-fonts-colour.md` §4 | **Never** claim a preview shows the organizer's real typeface when `missingFonts` is non-empty |
| C5 | Theme font changed | V08 compares `<a:fontScheme>` major/minor latin+ea to the template | **FAIL** naming each changed slot | that a colour or effect in the theme is still right (same check covers the part, but only font slots are read out and diffed) |
| C6 | Colour shift through PDF/PNG export | measured: solid `#0072B2` and `#D55E00` round-trip **exactly** at 144 dpi | `render.py` records pdf sha256; swatch verification is a harness measurement, not a shipped gate | per-slot theme **colour** palette behaviour, and colour under a different renderer |

## D. Template fidelity

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| D1 | Master or layout modified | V07 C14N-compares every `slideMaster`/`slideLayout` part against the template's own XML (12 parts on the organizer fixture) | **FAIL** listing the parts | that the deck *uses* the right layout for a slide's role — only that the parts are untouched |
| D2 | Template part deleted | V07 notices a part present in the template and missing from the deck | **FAIL** | — |
| D3 | Builder's `invariants.master_unchanged` lies | V07 never reads `build_report.json`'s boolean; it re-derives from OOXML | — | — |
| D4 | Template cannot be located | V07/V08/V09 | **UNVERIFIED** (fails the gate) stating every path tried | — |
| D5 | Case declares the template with a **case-relative** path | resolution tries the value as given and relative to the case file's dir and the work dir (measured: `materials/template-organizer.pptx` resolved) | — | — |
| D6 | Footer / slide-number placeholder filled by the builder | V09 finds `ftr`/`sldNum`/`dt` placeholders on slides with non-empty text | **FAIL** | that a *deleted* placeholder in a layout is caught — that is D1 |
| D7 | Template has no footer placeholder at all | V09 | **SKIP** with a recorded reason (nothing to preserve) | — |
| D8 | Master/layout changed but re-serialised identically | V07 uses C14N, so prefix/whitespace churn is not a false alarm | — | changes that are semantically invisible in the part but visible at runtime (none known) |

## E. Traceability, numbers and evidence

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| E1 | **Fabricated or rounded number on a slide** — the single worst rubric failure (`task/task.md:42`) | V04 extracts numeric tokens from the **text actually on the slides** (deck text, not `plan.json`), strips figure/table/section labels, and requires a verbatim match in `material.numbers` | **FAIL** listing slide:token | that the number is *used correctly* in the argument — only that it is attested in the source |
| E2 | `material.numbers` is empty | V04 | **UNVERIFIED** — it refuses to pass by default | — |
| E3 | Figure/table reference does not resolve | V05 checks `figure_ref`/`table_ref` and evidence refs against `material.json` | **FAIL** | that the embedded picture is that figure (D-side of E4) |
| E4 | Embedded image with no real provenance (placeholder, invented curve) | V06 counts pictures **in the deck** and compares with the provenance modes declared in `build_report.json`; only `source-crop` and `redraw-from-data` are allowed | **FAIL** on an undeclared embedded picture, or on a declared mode that is not an allowed one | that the pixels equal the paper's figure — it proves the declared *mode*, and says so in the PASS detail |
| E5 | A figure is referenced but never embedded | V06 (declared with an allowed mode but no picture found) | **FAIL** | whether the omission was intentional (record it in `plan.excluded` instead) |
| E6 | Locator quote not present in the source | `evidence.py` classifies `exact` / `paraphrase` / `unsupported`; only `exact` sets `verified: true` | `paraphrase` and `unsupported` items go into the page's `unverified[]` with a reason | that a paraphrase is wrong — it is simply not a quote, so it is not marked verified |
| E7 | Figure/table caption text absent from the source page text | measured on the fixture: the locator quote `"Apparatus"` was **not** in `material.sources[].pages[].text`, so the figure was reported unsupported | records the gap honestly | — (the fix belongs in `ingest.py`: include caption text in the page text) |
| E8 | Evidence map missing a slide, or a slide's entry is empty | V14 requires every slide to have a non-empty claims/figures/tables/numbers entry | **FAIL** | that each entry is *strong* evidence — that is E6's job |
| E9 | **Two slides reuse the same content** (a "filled" page count hides a repeated page) | V16 (a) groups every bullet by a **fingerprint** that strips a leading `"Label:"` prefix, and FAILs on any fingerprint occurring twice or more **anywhere in the deck** — including several times on one slide; (b) compares titles/bullets across slides for identical/contained text and ≥70 % **content-word** overlap (stopwords ignored); notes only for **exact** repeats | **FAIL**, naming the slides and the repeated text | that the remaining slides are *distinct in meaning* — only that their wording is not reused. It does **not** gate role-prefixed titles or heading+sentence concatenations: those have no mechanical signature that separates them from legitimate prose, and the pathological form (a repeated stub) is what rule (a) catches |
| E10 | **Constraint/input leakage into slide content** (the brief read back to the audience: `Audience: …`, `Duration: 10.0`) | V17 matches a *labelled* brief field (`audience\|duration\|slide_count\|template\|constraints` + `:`/`=`), the declared audience string, the template file name, and a `<duration> min` phrasing | **FAIL** per slide/text unit | nothing about prose that merely mentions the same subject — the label separator is required, so "the duration of the reaction" is not flagged |
| E11 | **Empty-but-green: the source was gutted, so every check passes vacuously** (`all()` over an empty list is true). A scanned PDF with no text layer is the canonical trigger | V19 (R1) notices a declared source file with bytes but < 200 extracted chars and no figures/tables/numbers/claims; R2 fires when the source has numbers but none reach a slide; R3 when the source had content and the deck has no bullets; R4 when a content-role page has nothing but a title; R5 when every figure/table is dropped | **FAIL**, naming the rule that fired, with the measured coverage in the detail | that a *short* paper is a gutted one — the thresholds were derived from the two real fixtures (≈13 000 extracted chars, 14–17 bullets, 15–21 numbers on slides) and are stated in `verification_gates.md` §4 |
| E12 | Partial figure/table curation that the audit trail does not explain | V19 emits a **WARN** (never a failure) when a source figure/table is never referenced and its id is not found — as a **word-boundary substring** — in any `plan.excluded[].item` **or** `.reason` | advisory note in `verify_report.warnings[]` and the check's detail; the status is unchanged | that every figure *should* be used — a deck need not, but the reason should be recorded. **History:** the first revision compared the bare id against a *set of whole item strings*, so an entry like `"fig2 (figure Figure 2) not placed: …"` was not recognised and the advisory asserted something untrue about our own artifact. It now matches inside the item/reason text (`(?<![\w-])fig2(?![\w-])`, so `fig2` is not satisfied by `fig20`), and the advisory is proven in both directions: silent where explained (caseA), present where genuinely unexplained (a variant with fig3's explanation removed) |

## F. Result contract

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| F1 | A declared `output_contract` key is silently dropped | V13 compares the key set of `case.json.output_contract` against `result.json` | **FAIL** listing the missing keys (measured: dropping `summary_audio` → V13 FAIL, exit 1) | — |
| F1b | A declared key is present but null/empty **with no explanation** | V13 checks that each null/empty declared key is named by some `uncertainty` entry | **FAIL**; CONTRACT.md:33-34 requires every undeliverable key to carry a reason rather than just being blank | — |
| F1c | `result.json` does not exist because the producing stage has not run | V13 looks in **both** `--out` and `--work`; absence is reported as **UNVERIFIED** ("producer stage did not run"), never as a FAIL against the skill's output | fails the gate either way (CONTRACT.md:196), but the status names the real situation | that a downstream stage can be assumed — the fix is to run `result_contract.py` before `verify_deck.py` |
| F2 | A declared key has no counterpart in the skill | `result_contract.py` emits `null` (or `[]`/`{}` when the case's own default declares that type), adds it to `_unmapped_contract_keys` **and** writes one `uncertainty` entry per key | measured on a synthetic contract: 3 unmatched keys → `null`/`[]`/`{}` + `_unmapped_contract_keys` + 3 explanations | the value the platform wanted |
| F3 | Contract uses unexpected spellings | tolerant synonym table; measured with two deliberately disjoint fixtures (`deck_pptx/deck_pdf/slide_previews/speaker_notes/page_evidence_map/plan` **and** `pptx/pdf/previews/notes/evidence/outline/status/uncertainty`) | both filled with zero absent and zero empty keys | a key whose intent cannot be inferred — it goes to F2 rather than being guessed |
| F4 | A mapped key's artefact was not produced | the key is present with `null`/`[]` and gets an `uncertainty` entry naming the slot it mapped to | — | — |
| F5 | `build_report.json` describes a *different* deck | V16 compares `build_report.pptx` with the deck actually inspected | **FAIL** — and the work-dir deck is what gets inspected, never the foreign path | — |
| F6 | `result.json` missing entirely | V13 | **FAIL** | — |

## G. Evidence of process

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| G1 | No version records | V15 scans matrix names (`versions.json`, `change_records.json`, `diff*.md`, `rollback*.md` — pinned with `deck-core`) in **`--out` first, then `--work`** | **UNVERIFIED** when nothing exists; **FAIL** when present but missing v1 / changes / diff / rollback | — |
| G2 | Slide count outside the case constraint | V01 counts slides **from the deck file** and reads the constraint from the case (several spellings) | **FAIL**; if no constraint exists anywhere → **UNVERIFIED**, never a pass | — |
| G3 | Slide without title/claim, or notes < 40 chars | V02 counts a real deck title; V03 measures notes on the deck's **notes slides** (not `plan.json`) | **FAIL** per slide | whether the notes are *good* |
| G4 | A check reports SKIP with no reason | `Gate.add()` rewrites a reasonless SKIP into **UNVERIFIED** | the gate cannot be weakened by accident | — |

## H. Environment and input

| id | Failure mode | How it is detected | What the skill does | What it does NOT certify |
|---|---|---|---|---|
| H1 | `pip` unusable (measured: `tempfile.mkdtemp` 0o700 → WinError 5) | — | stdlib + already-installed packages only; every optional import is guarded (`pptx`, `PIL`, `lxml`, `pypdf`, `pdfplumber`, `fontTools`) | that the eval container has the same packages — guards degrade instead |
| H2 | `git`/`curl` fail natively | — | `urllib` only | — |
| H3 | Author-machine absolute path leaks into the tree | — | renderer discovery is env/PATH/bounded-search; scratch harnesses stay outside `work/submission` | — |
| H4 | `__pycache__` in the submission tree | — | `PYTHONDONTWRITEBYTECODE=1` everywhere; verified absent after all runs | — |
| H5 | `pdfplumber`/`pypdf` unavailable | V10's render cross-check becomes unavailable | the check still runs on geometry and the detail says the cross-check did not happen; PDF page count falls back to a stdlib `/Type /Page` count | — |

---

## What this skill explicitly does NOT certify

1. **Scientific correctness or presentation quality.** `task/task.md:97` says it plainly:
   technical file checks do not certify the science. Our checks prove structure, traceability and
   layout mechanics — nothing about whether the argument is right or the talk is good.
2. **A benchmark score.** The platform scores. This skill never invents a number for it.
3. **Font fidelity of a preview.** See C4 and `exports-fonts-colour.md` §4: when
   `missingFonts` is non-empty (measured normal here for **Liberation Serif**), the PDF/PNG show
   substituted glyphs — measured substitution was Liberation Serif → **Times New Roman**.
4. **Native-application fidelity.** LibreOffice pagination, line breaking and shaping can differ
   from PowerPoint; a preview is evidence about the file, not a promise about the deck.
5. **Media, animation, transitions, audio.**
6. **That a PASS means the deck is good.** A green gate means no *known* defect was found by the
   checks that actually ran, and it names the checks. It is a floor, not a grade.
7. **Numbers the source itself contradicts.** V04 proves a number is *attested* in
   `material.numbers`; it cannot know the paper's own numbers are internally consistent.
8. **That an estimate equals a measurement.** V10 is an estimate with a rendered cross-check
   where a PDF exists; B1–B4 are the limits of that.

## How this ledger was produced

| Harness (scratch, not shipped) | Measured result |
|---|---|
| `work/tmp-verify-render/negtest.py` | 18 injected defects; 18/18 detected by the expected check; every case exit ≠ 0 |
| `work/tmp-verify-render/contract_test.py` | caseA + caseB contract spellings fully filled; 3 unmatched keys → `null`/`[]`/`{}` + `_unmapped_contract_keys` + 3 uncertainty entries; dropped key → V13 FAIL; case-relative template resolved |
| `work/tmp-verify-render/organizer_test.py` | real organizer template: `missingFonts ["Liberation Serif"]`, PDF embeds `TimesNewRomanPSMT`; V07/V08/V09/V12 PASS |
| `work/tmp-verify-render/timeout_test.py` | pipe-holding grandchild → `rc=124` in 7.8 s against a 5 s deadline (previously an indefinite hang); normal child output still captured |
| `work/tmp-verify-render/v19_gutted_test.py` | real `ingest.py` on a PDF whose page has no text layer → `chars=0 figures=0 tables=0 numbers=0 claims=0`, V19 **FAIL** (R1 + R4), gate exit 1 |
| `work/tmp-verify-render/measure_coverage.py` | the caseA/caseB coverage numbers V19's thresholds were derived from |
| `work/tmp-verify-render/mkfixture.py` + `verify_deck.py` | good fixture: verdict **PASS**, 19/19 checks (V01–V19), exit 0 |
| `render.py` idempotence run | two consecutive runs into one work dir, both `status ok` |
| `render.py` with no renderer | `status skipped`, `available false`, exit 1, no paths claimed |

### V16 / V17 were seen red before they were trusted

A gate added only after the defect is fixed has never been observed failing, so it is not
evidence. Both new gates were run against a **real caseA deck that still had the defects**, and
against the good fixture:

| Deck | V16 | V17 | Overall |
|---|---|---|---|
| lead's real caseA deck (before `deck-core`'s fix) | **FAIL** — 3–4 reused pairs (slide 2 title ≡ slide 4 title; slide 8 title contained in slide 7 bullet) | **FAIL** — slide 1 bullets `Audience: …` and `Duration: 10.0` | 16 PASS / 2 FAIL / 0 UNVERIFIED, exit 1 |
| real caseA after `deck-core` fixed the leakage | **FAIL** — 2 reused pairs remain | **PASS** | 16 PASS / 2 FAIL (V10, V16), exit 3 |
| good fixture (no defects) | PASS | PASS | **PASS**, 19/19, exit 0 |

### V19 was seen red before it was trusted

| Run | V19 | Whole-gate verdict |
|---|---|---|
| gutted no-text-layer input (`v19_gutted_test.py`) | **FAIL** (R1 "source read but nothing extracted" + R4 title-only pages) | FAIL, exit 1 |
| real caseA (8-slide talk) | **PASS** (+ 1 advisory WARN: `fig2, fig3, fig4, tab2` unreferenced and not in `plan.excluded`) | FAIL 17/2 (V10, V16) |
| real caseB (12-slide review) | **FAIL** — R4: slides 4, 11, 12 have roles `method`/`limitation`/`conclusion` but only a title | FAIL 17/2 (V10, V19) |
| synthetic good fixture | **PASS** | **PASS**, 19/19, exit 0 |

A raw figure-drop-rate threshold was **rejected**: caseA uses 1 of 4 figures (drop rate 0.75)
because it is an 8-page talk with one result figure, so such a rule would fail a good deck. The
chosen rules and the reason for each are in `verification_gates.md` §4.
