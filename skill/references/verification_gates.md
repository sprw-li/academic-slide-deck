# Verification gates V01–V19 — authoritative semantics

`scripts/verify_deck.py` is the skill's self-check gate. It writes
`verify_report.json` (`schema: verify_report/v1`) and **exits 0 iff `verdict == "PASS"`**.

This file is the contract for what each gate reads, when it is allowed to pass, and what a
red result means. If the code and this file disagree, the **code is wrong** — fix the code and
re-measure, never the other way round.

---

## 1. PASS semantics, and the three outcomes that are *not* PASS

`CONTRACT.md:196`: **`verdict=PASS` requires zero `FAIL` and zero `UNVERIFIED`.** `SKIP` is
legal only with a recorded reason. So there are three distinct non-PASS outcomes and they must
not be conflated:

| Outcome | Status(s) in `checks[]` | Fails the gate? | Meaning | Recorded in |
|---|---|---|---|---|
| **N/A — not applicable** | `SKIP` + a reason starting `N/A:` | **No** | the check has nothing to inspect *because the situation does not apply* (e.g. the template declares no footer, so "preserve the footer" is vacuous) | `checks[].detail` |
| **WARN — advisory** | none; **the status is unchanged** | **No** | the gate noticed something worth a reviewer's eye but is not allowed to fail the submission on it | `warnings[]` (+ the check's `detail`) |
| **BLOCKING — inspected and wrong** | `FAIL` | **Yes** | the gate read the artefact and the property does not hold | `defects[]` with severity |
| **BLOCKING — could not inspect** | `UNVERIFIED` | **Yes** | the gate could not read its evidence. This is deliberate: an unverifiable claim is not a verified claim | `defects[]` with severity |

A `SKIP` with **no** reason is silently rewritten to `UNVERIFIED` by `Gate.add()`, so the gate
cannot be weakened by omitting an explanation by accident.

### Which gates may ever be N/A

| Gate | May be N/A? | Why |
|---|---|---|
| V01–V18 | **No** | each has a well-defined subject (the deck, the plan, the material, the template, the contract, the artefact set). Missing input is `UNVERIFIED`, not N/A, because the property *should* be checkable. |
| **V19** | **Yes** | only when the case declares no source material at all *and* no `material.json` exists — there is then genuinely nothing to measure coverage against. Recorded as `SKIP` with reason `N/A: …`. |

V09 additionally uses `SKIP` (with a reason, **not** the `N/A:` prefix) when the template itself
has no footer/slide-number placeholder: there is nothing to preserve. That is a recorded skip,
not an N/A, because the check's subject exists but is empty.

---

## 2. The gate table

`evidence` column = the artefact each check actually reads (its `checks[].evidence` pointer
follows the same item). "severity" is the severity written into `defects[]`.

**The deck reader has two backends.** `verify_deck.py` reads the deck through `zipfile` +
`xml.etree.ElementTree` (both stdlib) into a model; python-pptx, when importable, is only an
equivalent adapter. Measured with the imports blocked the way a thin container would
(`BLOCK_IMPORTS=pptx`, `pptx,lxml`, and `pptx,lxml,pypdf,pdfplumber,PIL,matplotlib`), the
verifier returns the **same verdict** `PASS 20/20` on a real deck in every case — it neither
crashes nor degrades to UNVERIFIED. V07/V08/V09 use ElementTree structural comparison when lxml
is absent, so the template checks survive too.

| id | What it checks | Reads | FAIL when | UNVERIFIED when | Sev |
|---|---|---|---|---|---|
| **V01** | slide count within the case constraint | **the deck file** (slide count) + `case.json`/`plan.constraints` (constraint) | count is outside the declared `slide_count` / `slide_count_range` | no constraint found in any accepted spelling; or no readable deck | high |
| **V02** | every slide has a non-empty title and a claim | deck title placeholders + `plan.json` | a slide has no title in deck *or* plan, or the plan title is absent from the deck, or no claim | neither plan slides nor a readable deck | high |
| **V03** | every slide has ≥ 40 chars of speaker notes | the deck's **notes slides** (not `plan.json`) | any slide's deck notes < 40 chars (the detail says if the plan had enough but the deck did not) | no deck and no plan slides | high |
| **V04** | every number on a slide is traceable to the source | **deck slide text** + `material.numbers` | a numeric token (caption labels stripped) has no verbatim match in `material.numbers` | `material.json` missing, or `material.numbers` empty (refuses to pass by default) | high |
| **V05** | every `figure_ref`/`table_ref` resolves | `plan.json` vs `material.json` | a ref (including evidence refs) names an id not in `material.figures`/`tables` | `material.json` missing | high |
| **V06** | every embedded image has real provenance | **pictures counted in the deck** vs `build_report.json` modes | an embedded picture has no allowed mode (`source-crop`, `redraw-from-data`); or a declared embed is `placeholder`/`reference-only` while a picture IS present; or a declared allowed-mode image is not embedded | neither deck nor build_report | high |
| **V07** | template masters/layouts unmodified | **OOXML parts** of deck vs template (C14N) | any part differs, or a template part is missing from the deck | deck parts unreadable; **template not locatable** | high |
| **V08** | theme fonts unchanged | `ppt/theme/themeN.xml` `<a:fontScheme>` in deck vs template | any of major/minor latin/ea differs | template not locatable, or its theme has no readable fontScheme | high |
| **V09** | footer / slide-number placeholders survive, unfilled | masters/layouts/slides of deck vs template | placeholders gone while the template had them; or the builder wrote text into one | deck parts unreadable; template not locatable | medium |
| **V10** | no text overflow | **deck geometry** (real font advances) + rendered PDF via `pdfplumber` | a no-autofit shape's estimated extent exceeds its box; or an `spAutoFit` shape would grow off the slide; or a rendered word lies outside the page box | no readable deck | medium |
| **V11** | images not cropped, inside bounds | picture OOXML (`srcRect`, namespace-agnostic) + frame geometry | non-zero `srcRect`; or a frame outside the slide box | no readable deck | medium |
| **V12** | fonts resolvable, substitution reported | `typeface=` values on slides/notes/layouts/masters + theme major/minor, resolved against platform fonts; vs `render_report.missing_fonts`/warnings | an unresolvable family is reported **nowhere** | no attributable typeface at all | medium |
| **V13** | `result.json` fills every `output_contract` key | `case.json.output_contract` vs `result.json` in `--out` then `--work` | result.json exists but a declared key is missing; or a key is null/empty with **no** `uncertainty` entry naming it | **result.json absent** → "producer stage did not run" | high |
| **V14** | evidence map covers every slide | `evidence_map.json` pages vs deck slide count | a slide has no entry, or its entry has zero claims/figures/tables/numbers | `evidence_map.json` missing/empty | medium |
| **V15** | version records exist (v1 + changes + diff + rollback) | `versions.json`, `change_records.json`, `diff*.md`, `rollback*.md` in `--out` then `--work`, plus `result.json.change_records` | artefacts exist but v1 / changes / diff / rollback is missing | nothing found anywhere | medium |
| **V16** | no duplicate slide content | titles/bullets/notes across slides | a **bullet fingerprint repeats anywhere in the deck** — a leading `"Label:"` prefix is stripped first, so `"Problem: what this page has to establish."` … `"Conclusion: …"` all collapse to one fingerprint (this fires whether the repeats sit on two slides **or several times on one slide**); or two units are identical / one contains the other / share ≥ 70 % of the shorter side's **content words** (stopwords ignored); notes only for exact repeats | no plan bullets and no readable deck text | high |
| **V17** | no constraint/input leakage into content | slide/notes text vs declared metadata | a *labelled* brief field (`Audience:`, `Duration:`, `slide_count =`), the declared audience string, the template file name, or a `<duration> min` phrasing appears | no slide text; or no constraint metadata declared at all | medium |
| **V18** | `build_report` describes the deck under inspection | `build_report.pptx` vs the deck actually inspected | the two disagree (the **work-dir deck** is always the one inspected) | — | high |
| **V19** | minimum content coverage (anti false-green) | declared source files, `material.json`, `plan.json`, slide text | see §4 — any of R1–R5 | — (may be **N/A → SKIP**) | high |
| **V20** | the run produced deliverables | `--out` contents (a readable `.pptx`, `slide-*.png`), `result.json.status`, `render_report.status` | no readable deck in `--out` and none declared by `result.json`; or `status=="failed"`; or render says `ok` while fewer previews than slides exist | `status=="partial"` — a deck was delivered but the run declares itself incomplete, so completeness cannot be certified | high |

### V20's exact rule (direction matters)

A missing artefact must never read as success, but a partial-but-honest run must still report its
real verdict. The implemented rule, in order:

1. **no readable `.pptx`** in `--out` (and none valid at a path `result.json` declares) → `FAIL`.
   A run that produces no deck cannot be graded as a complete solution. A file that exists but is
   not a readable deck (not a zip / no slide parts / 0 bytes) is also `FAIL`.
2. `result.json.status == "failed"` → `FAIL`.
3. `result.json.status == "partial"` → **`UNVERIFIED`** — deliberately *not* PASS (incomplete is
   not success) and deliberately *not* FAIL (the deck is real, and every other check keeps its own
   verdict; one incomplete run must not erase the rest of the report).
4. `render_report.status == "ok"` but fewer non-empty previews than slides → `FAIL` (the render
   advertised itself as complete and is not).
5. otherwise → `PASS`, with the deck size, slide-part count, preview count and both statuses in
   the detail. If the deck was only found at a path declared by `result.json` rather than inside
   `--out`, a **WARN** notes that the platform reads the output directory.

**Why `status` is not derived from the verdict:** V20 reads `result.json.status`, so if
`result_contract.py` computed `status` from the gate verdict the two would be circular and a run
could never climb out of a provisional `partial`. `status` therefore describes **deliverable
completeness only**; correctness is published separately in `result.json._self_check`
(`verify_verdict`, `checks_failed`) and repeated in `uncertainty`. Measured: before the split, a
complete run reported `partial` forever because the first `result_contract` pass ran before the
gate.

Two structural notes:

* **V07/V08/V09 are OOXML comparisons against the template's own parts.** They never read
  `build_report.invariants`, which is the builder's self-report. A gate that trusted those
  booleans would pass anything.
* **V04 is re-derived** from `material.json` plus the text that is actually on the slides —
  never from `plan.json.numbers_used`, which is also written by the builder.

---

## 3. Concrete failing examples (all reproduced by injection)

Each row was produced by mutating a good fixture and re-running the gate; the harness is
`work/tmp-verify-render/negtest.py` (18 injected defects, **18/18 detected by the expected
check, every case non-zero exit**).

| Gate | Injected defect | Observed result |
|---|---|---|
| V01 | `case.constraints.slide_count` changed 3 → 7 | `FAIL` "deck has 3 slides, case requires exactly 7" |
| V02 | blanked one `plan.slides[i].claim` | `FAIL`, names the slide |
| V03 | set one deck notes slide to `"too short"` | `FAIL` "deck notes 9 chars" |
| V04 | rewrote a slide run `12.3` → `99.9` | `FAIL` "1 numeric token(s) … no verbatim match: slide 2:99.9" |
| V05 | set `plan.slides[2].figure_ref = "fig9"` | `FAIL` "figure_ref 'fig9' not in material.figures ['fig1']" |
| V06 | embedded a picture on slide 2 that `build_report` does not declare | `FAIL` "1 embedded image(s) but only 0 declared with an allowed provenance mode" |
| V06 | declared mode `placeholder` while the image is embedded | `FAIL` "declared mode 'placeholder' … is not an allowed provenance mode" |
| V07 | injected an attribute into `ppt/slideMasters/slideMaster1.xml` | `FAIL` "1 template part(s) differ from the template's own XML" |
| V08 | rewrote a theme `typeface="Calibri"` | `FAIL` "theme font major_latin changed from 'Calibri' to 'Injected Font'" |
| V09 | injected a `ftr` placeholder carrying text on slide 1 | `FAIL` "the builder wrote text into footer/slide-number placeholders" |
| V10 | put ~550 pt of text in a 0.5 in `spAutoFit` textbox at y=5 in | `FAIL` "grows to … bottom / slide is … tall — the content runs off the slide" |
| V11 | inserted `<a:srcRect l="12000" …/>` into a picture's `<p:blipFill>` | `FAIL` "cropped via srcRect l=\"12000\"…" |
| V12 | named `NoSuchFontZzz` on a run **and deleted** `render_report.json` | `FAIL` "font 'NoSuchFontZzz' substituted silently" |
| V13 | removed a declared key (`summary_audio`) from `result.json` | `FAIL` "result.json is missing 1 declared key(s): ['summary_audio']" |
| V13 | removed `result.json` entirely | `UNVERIFIED` "producer stage did not run" |
| V14 | deleted slide 3 from `evidence_map.pages` | `FAIL` "no evidence entry for slide(s) [3]" |
| V15 | deleted `versions.json` + `rollback.md` and emptied `result.json.change_records` | `FAIL` "incomplete: ['no v1 version record', 'no rollback instruction']" |
| V16 | real caseA deck with slide 2 title ≡ slide 4 title, slide 8 title contained in slide 7 bullet | `FAIL` "3 pair(s) of slides reuse content: title of slide 2 is identical to title of slide 4" |
| V16 | real caseB deck, slide 2 carrying `Problem: what this page has to establish.` / `Method: …` / `Key result: …` / `Limitations: …` / `Conclusion: …` | `FAIL` "the same bullet text appears 5 times on slide 2 (after stripping any 'Label:' prefix), so the page carries boilerplate rather than content: 'what this page has to establish'" — this is the case that a character-containment rule **missed** (the shared tail was ~77 % of each bullet, just under an 80 % character threshold) |
| V17 | real caseA deck with bullets `Audience: 课题组组会 …` and `Duration: 10.0` | `FAIL` "bullet repeats the brief's own 'Audience' field" / "'Duration' field" |
| V18 | `build_report.pptx` pointing at a different deck | `FAIL` "points at … but the work dir holds …" |
| V19 | real `ingest.py` run on a PDF whose page has **no text layer** | `FAIL` R1 + R4 — see §4 |
| V04/V05 | deleted `material.json` | `UNVERIFIED` (not FAIL, not PASS) |
| V07/V08/V09 | removed `template_path` from case **and** `plan.constraints` **and** `build_report.template.path` | `UNVERIFIED` "template .pptx could not be located" |

---

## 4. V19 — minimum content coverage (and the threshold, derived)

### The false-green it kills

If a scanned or oddly-encoded PDF makes `ingest.py` drop most of the source, `material.json`
becomes nearly empty. Every downstream check is then satisfied **vacuously** — `all()` of an
empty list is `True` — so a deck of template boilerplate can reach `verdict: PASS`. V19 refuses
to pass a run that had too little to work with.

### Thresholds, and how they were chosen

Measured on the two real fixtures **before** choosing anything
(`work/tmp-verify-render/measure_coverage.py`):

| Fixture | slides | bullets | extracted chars | figures | tables | numbers | distinct numbers on slides | title-only |
|---|---|---|---|---|---|---|---|---|
| caseA (8-slide talk) | 8 | 14 | 13 880 | 4 | 2 | 370 | 21 | slide 1 (role `title`) |
| caseB (12-slide review) | 12 | 17 | 13 091 | 3 | 2 | 143 | 15–22 | slides 1, 4, 11, 12 |

The first data point killed the obvious rule. **A raw figure-drop-rate threshold does not
work**: caseA legitimately uses 1 of 4 figures — a drop rate of **0.75** — because it is an
8-page talk with one result figure. A "drop rate > 0.5 ⇒ FAIL" rule would fail a perfectly good
deck, so V19 does **not** use one. Instead:

| Rule | Condition (all must hold to fire) | Why this shape | caseA | caseB | gutted |
|---|---|---|---|---|---|
| **R1** the gutted-input detector | ≥ 1 declared source file with bytes > 0, **and** extracted text < **200 chars**, **and** zero figures/tables/numbers/claims | a real file that yields nothing is a decode failure, not a short paper. 200 chars is far below any real paper page (caseA/caseB extract ~13 000) and above nothing | pass (13 880) | pass (13 091) | **FAIL** (0 chars) |
| **R2** numbers never reach the deck | ≥ 4 slides, `material.numbers` non-empty, **0** distinct numeric tokens on any slide | the vacuous-`all()` class in its purest form | pass (21) | pass (15+) | n/a (no numbers) |
| **R3** boilerplate deck | ≥ 3 slides, source supplied ≥ 1 figure/table/number, **0 bullets** across the deck | a deck with no bullets has no narrative, whatever the page count | pass (14) | pass (17) | **FAIL** (0 bullets, but source empty → R1 carries it) |
| **R4** title-only content page | ≥ 5 slides, and a slide whose role is **not** a divider role has no bullets, no figure/table ref and no numbers | a page whose whole content is its title. Divider roles (title/section/outline/ack/backup/references/qa/…) are exempt so real dividers are not punished | pass (slide 1 is role `title`) | **FAIL** — slides 4, 11, 12 have roles `method`/`limitation`/`conclusion` but only a title | **FAIL** (all 5) |
| **R5** complete figure/table loss | source has ≥ 2 figures (or ≥ 2 tables) and the deck references **none** | complete loss is a defect; partial curation is a judgement call | pass (1 of 4) | pass (2 of 3) | n/a |
| **WARN** unaccounted drops | source figures/tables never referenced and their id is not found — as a **word-boundary substring** — in any `plan.excluded[].item` **or** `.reason` | **advisory only** — a deck need not use every figure, but the audit trail should say why. deck-core writes items like `"fig2 (figure Figure 2) not placed: …"`, so the id is matched inside the item/reason rather than by whole-string equality | — | — | — |

R4 is why caseB is red and caseA is green: it fires on *role*, not on page count. The fix is
either to give those pages content or to label them `section`. That is a real finding, not a
threshold artefact — and it is exactly the sort of page a non-expert reviewer notices instantly.

### V19 seen red before being trusted

A threshold nobody has seen fail is not a threshold. `v19_gutted_test.py` builds a **real** PDF
whose page carries no text layer (`PdfWriter().add_blank_page(...)`, 431 B), feeds it to the
**real** `ingest.py`, and shows the failure end to end:

```
blank pdf: 431 B (has a page, no text layer)
ingest exit: 0
ingested: sources=1 chars=0 figures=0 tables=0 numbers=0 claims=0
gate exit: 1 | verdict: FAIL {'PASS': 14, 'FAIL': 3, 'SKIP': 0, 'UNVERIFIED': 2}
V19: FAIL
V19 detail: R1 source read but nothing extracted: 1 declared source file(s) totalling 431 B
            yielded only 0 chars of text and 0 figures/tables/numbers/claims. A scanned PDF
            with no text layer or a failed decode presents exactly this way, and every other
            gate would pass vacuously over the empty lists;
            R4 slide(s) [1, 2, 3, 4, 5] have a content role but nothing but a title …
```

Two independent rules (R1 and R4) fire on the same gutted run, and the gate exits non-zero.
On the real fixtures V19 then goes **green on caseA** and red on caseB for the role reason
stated above:

| Run | V19 | Whole-gate verdict |
|---|---|---|
| gutted no-text-layer input | **FAIL** (R1 + R4) | FAIL, exit 1 |
| real caseA (8-slide talk) | **PASS** (+ 1 advisory WARN about `fig2-fig4`/`tab2`) | FAIL 17 PASS / 2 FAIL (V10, V16) |
| real caseB (12-slide review) | **FAIL** (R4: slides 4, 11, 12) | FAIL 17 PASS / 2 FAIL (V10, V19) |
| synthetic good fixture (3 slides, no defects) | **PASS** | **PASS**, 19/19, exit 0 |

---

## 4b. V16 — the repetition threshold, and why the first one was wrong

V16's first revision compared **characters**: identical, contained, or ≥ 80 % character overlap.
That let a real defect through. On the real caseB deck, slide 2 carried five bullets whose role
prefixes differ but whose tail is shared:

```
- Problem: what this page has to establish.
- Method: what this page has to establish.
- Key result: what this page has to establish.
- Limitations: what this page has to establish.
- Conclusion: what this page has to establish.
```

The shared tail is 31 of ~40 characters ≈ **77 %** — just under the 80 % line — and the pairwise
comparison also skipped same-slide pairs, so the deck reached `verdict: PASS`
(**PASS 19/19** on the current caseB). Two changes close the class rather than the instance:

| Change | Rule | Why this shape |
|---|---|---|
| **Repeated bullet fingerprint** | strip a leading `"Label:"`/`"Label："` prefix (≤ 24 chars), lowercase, collapse whitespace, drop trailing punctuation; **two or more occurrences of the same fingerprint anywhere in the deck FAIL** | boilerplate repeats *by construction*; the label is exactly the part that differs. The occurrence counter deliberately ignores slide boundaries, because this real defect had all five on **one** slide — a "same bullet on ≥ 2 slides" rule would have missed it too |
| **Content-word overlap** | ≥ **70 %** of the shorter side's **content words** (stopwords removed), both sides ≥ 3 content words | the same sentence with a different label cannot hide behind shared filler words. This replaces the character percentage |

Guards against false positives, all measured on the good fixture (which stays **PASS 19/19**):

* fingerprints shorter than **20 characters** are ignored, so short repeated navigation phrases
  do not fire;
* notes are compared for **exact** equality only and never against a title/bullet — house-style
  speaker notes that differ only by slide number are legitimate, and a fuzzy test on them would
  fail every deck built from a template;
* the content-word test needs ≥ 3 content words on both sides, so short headings are not matched
  on one shared noun.

**Before → after, on the real deck (same artefacts, only the gate changed):**

| Run | V16 | Whole gate |
|---|---|---|
| caseB, V16 revision 1 (character threshold) | **PASS** — the gap | **PASS 19/19, exit 0** |
| caseB, V16 revision 2 (fingerprint + content words) | **FAIL** — "the same bullet text appears 5 times on slide 2 …" | **FAIL 18/1, exit 1** |
| good fixture, revision 2 | PASS | **PASS 19/19, exit 0** |
| real caseA, revision 2 | FAIL (reused slide 2/4 title etc.) | FAIL 17/2 (V10, V16) |

### Titles and concatenated heading+sentence: deliberately **not** gated

Observed on the same deck: titles shaped `"Key result: Introduction"`, `"Problem: Abstract"`,
`"Figure: Synthetic screening flow"`, and a bullet that concatenates a heading with its sentence
without punctuation (`"Scope and Literature Selection A synthetic search returned 412 records."`).

These are **not** gates, on purpose:

* a role-prefixed title can be perfectly reasonable, and `"Figure: Synthetic screening flow"` is a
  normal descriptive title — failing it would punish legitimate outline pages;
* "heading concatenated with a lead sentence" has no mechanical signature that separates it from a
  legitimate hand-authored heading followed by a sentence. Any rule strict enough to catch the
  observed case would also fail ordinary prose;
* the *pathological* form of this class — the same stub repeated — **is** caught, by V16 rule 1
  above, which is the mechanically sound part of the complaint.

Handled as documented review guidance (`README.md`, rubric mapping) instead. If a future revision
wants a signal without failing anything, the natural home is a **WARN** entry (see §1), not a
gate.

---

## 5. What the gate does **not** certify

Restated here because it is the first thing a hostile reviewer should hold us to
(full ledger: `failure-modes.md`):

1. Not scientific correctness and not presentation quality — `task/task.md:97`.
2. Not a benchmark score; the platform scores.
3. Not font fidelity of a preview: measured substitution is Liberation Serif →
   **Times New Roman**, and `render_report.missing_fonts` is what says so
   (`exports-fonts-colour.md` §4).
4. Not native-application fidelity (LibreOffice ≠ PowerPoint).
5. Not media/animation/audio.
6. `PASS` means *no known defect was found by the checks that ran*. It is a floor, not a grade.

## 6. Reproducing every claim in this file

| Harness (scratch, not shipped) | Command result |
|---|---|
| `negtest.py` | 18/18 injected defects detected by the expected check, all non-zero exit |
| `v19_gutted_test.py` | V19 FAIL on the no-text-layer input, gate exit 1 |
| `measure_coverage.py` | the coverage numbers in §4's threshold table |
| `contract_test.py` | caseA + caseB contract spellings filled; unmapped keys → `null` + `_unmapped_contract_keys`; dropped key → V13 FAIL |
| `timeout_test.py` | renderer timeout enforced (rc=124 in 7.8 s against a 5 s deadline) |
| `mkfixture.py` | good fixture → **PASS 19/19**, exit 0 |
