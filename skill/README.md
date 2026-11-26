# academic-slide-deck — academic presentation generation & editing with retained evidence

A **DSH Skill** for Bohrium challenge `ppt-6ad93034` / `ARMI:2611.000074v1`
(「学术 PPT 生成与编辑」).

It turns one paper or research material into a presentation a researcher could actually
stand up and give: pages that carry the argument, numbers that match the source, the
supplied `.pptx` template intact, an editable deck plus a real rendered export, and a
per-page evidence map that someone else can audit.

> **The one line:** the LLM decides *what the pages say*; the scripts do every read,
> every write into the file, and every check. `deck.pptx` is a build product — the plan is
> the source of truth.

---

## 1. What problem it solves

Researchers repeatedly re-translate the same material into slides — group meeting, proposal
and mid-term review, defence, project review, grant application. The difficulty is not
copying paragraphs into text boxes; it is **取舍与重排**: which result earns a page of its
own, how far the method can be compressed, whether Figure 3 or Figure 4 comes first, how
many points an audience can follow on one page.

The recorded failure modes of existing agents and tools are narrow and repeatable:

| Failure it is designed against | How this Skill prevents it |
|---|---|
| Abstract or paragraph pasted into a text box; page packed full | `plan.py` enforces one claim per page, ≤5 bullets, ≤120 chars each |
| Key numbers copied wrong or rounded by impression | every slide number must appear **verbatim** in `material.json`; gate V04 re-derives it from the source |
| Caption and in-figure annotation disagree | captions are extracted with a verbatim quote locator, not retyped |
| Placeholder image or an invented curve where a figure belongs | figures are `source-crop` or `redraw-from-data` only; otherwise `reference-only` (cited, not faked); gate V06 |
| Template fonts / colours / footer overwritten | the master is never mutated; gates V07–V09 read the OOXML directly |
| Editing one page breaks the master and every page after it | the deck is rebuilt from the plan, never edited in place |
| After export: font substituted, text overflowing, image cropped | gates V10–V12 on geometry, cross-checked against the render |
| After several rounds you cannot return to the previous version | `versions/` change records + deterministic rebuild from `plan.json` |

## 2. Install

The evaluator installs the Skill at `/skills/weekly-skill`. To use it elsewhere, put this
directory wherever your harness loads skills and invoke `SKILL.md`.

Requirements: Python ≥ 3.10 with `python-pptx`, `lxml`, `pypdf`, `pdfplumber`, `Pillow`,
`python-docx` (optional PDF rasterisation via `pypdfium2` or `PyMuPDF`). A renderer is
needed only for PDF/PNG previews; see §7 for discovery and honest degradation.

Nothing here installs packages at runtime, and no absolute machine path is hard-coded —
renderer and font locations come from environment variables, `PATH` and bounded known
candidates.

## 3. Use it — one command

```bash
python scripts/run.py --case /input/case.json --work /app/work --out /app/output
```

Then check it:

```bash
python scripts/verify_deck.py --case /input/case.json --work /app/work   # exit 0 iff verdict PASS
python scripts/selftest.py                                               # end-to-end + injected defects
```

Each stage also runs on its own (`--case/--work`, optional `--in/--out`), which is what
makes a failing stage debuggable and cheap to re-run:

```bash
python scripts/ingest.py  --case case.json --work w    # -> w/material.json
python scripts/plan.py    --case case.json --work w    # -> w/plan.json
python scripts/build_deck.py --case case.json --work w # -> w/deck.pptx + w/build_report.json
python scripts/render.py  --case case.json --work w    # -> w/render/deck.pdf + slide-XX.png + w/render_report.json
python scripts/evidence.py --case case.json --work w   # -> w/evidence_map.json
python scripts/verify_deck.py --case case.json --work w# -> w/verify_report.json (exit 0 iff PASS)
python scripts/result_contract.py --case case.json --work w --out o  # -> o/result.json
```

## 4. Inputs

| Input | Notes |
|---|---|
| `case.json` | declares `case_id` and an **`output_contract`** dict. |
| `task.md` | audience, duration, slide count, and any extra constraint |
| material | paper as PDF (text layer) / DOCX / MD / TXT; sometimes several files |
| template | a **master/layout-based editable `.pptx`** — the deck is built on it, not beside it |
| starting deck | for edit/adapt cases, an existing deck to modify rather than restart |

**The `output_contract` is read, not assumed.** Its key names differ between cases
(`deck_pptx` in one, `pptx` in another; `slide_previews` vs `previews`; …). The Skill maps
tolerantly, fills every declared key, and always adds `case_id`, `status`, `sources`,
`uncertainty`, `change_records` and `_self_check`. A declared key it cannot deliver is
emitted as `null` with an `uncertainty` entry and listed in `_unmapped_contract_keys` —
**never silently dropped** (gate V13).

## 5. Outputs

| Artifact | Meaning |
|---|---|
| `deck.pptx` | the editable deliverable, built on the supplied template |
| `deck.pdf`, `slide-XX.png` | real previews for inspection (honestly absent/skipped if no renderer) |
| `plan.json` | page roles, claims, bullets, figure/table refs, speaker notes, and what was **excluded** and why |
| `evidence_map.json` | per page: every claim / number / figure / table with its source locator and a verified verbatim quote |
| `verify_report.json` | gates V01–V15, counts, verdict, defects |
| `result.json` | every `output_contract` key plus honest status and uncertainty |
| `run.log`, `versions/` | per-stage timings, change records, diffs, rollback notes |

## 6. Method

```
material  ──ingest──▶ figures/tables/numbers each with a verbatim locator
          ──plan────▶ narrative arc, one claim per page, notes, layout chosen
                      from the template's REAL layouts, exclusions recorded
          ──build───▶ placeholders filled on existing layouts only;
                      master / theme / footer asserted unchanged
          ──render──▶ PDF + per-slide PNG (or an honest skip)
          ──evidence▶ per-page provenance, quotes re-verified against the source
          ──verify──▶ V01-V15; PASS requires zero FAIL and zero UNVERIFIED
          ──result──▶ contract-driven result.json
```

Two properties are load-bearing:

1. **Traceability is upstream.** If a fact cannot be given a verbatim quote during ingest,
   it never reaches the plan, so the slide cannot contain it. Nothing downstream has to
   guess whether a number is real. (Failures are still detected independently in V04.)
2. **Template safety is structural, not stylistic.** The deck is generated onto the
   template's own layouts, and V07–V09 re-open the package and compare its XML against the
   template's, so "the master was preserved" is an observation rather than a promise.

### Figures: what may be redrawn and what may only be referenced

| Situation | Mode | Rule |
|---|---|---|
| numbers exist in a real data table in the material | `redraw-from-data` | redraw with the same values; slide numbers must equal the table |
| only a picture of the figure exists | `source-crop` | crop the real figure region from the source |
| figure referenced but no usable image | `reference-only` | cite it in text; never a placeholder or an invented curve |
| nothing usable | — | report it in `uncertainty`; a missing figure is stated, not faked |

## 7. Rendering, fonts and honest degradation

PDF/PNG export is required by the task ("Editable Office/slide/diagram outputs should also
have real rendered PDF/PNG previews for inspection"). Renderer discovery tries, in order:
`LO_KIT_CLI`/`LO_KIT_NODE`, `SOFFICE_BIN`, `soffice`/`soffice.bin` on `PATH`, a bundled
LibreOffice Kit, then a bounded filesystem search. Windows runs use a private
`-env:UserInstallation` profile.

If nothing is found, `render_report.json` reports `status:"skipped"`, `available:false`
and an `uncertainty` entry — **no preview path is ever claimed that was not produced**.

Fonts are taken from the template. Organizer templates in this family name **Liberation
Serif** (OFL-1.1). If a family cannot be resolved the substitution is reported in
`render_report.json` (`missingFonts`) and in `result.json.uncertainty`, never hidden.
Measured behaviour per renderer is in `references/exports-fonts-colour.md`.

## 8. Self-check and what it does *not* certify

`verify_report.json` carries gates **V01–V20** (see `references/verification_gates.md` for the
authoritative per-gate table). `verdict = PASS` requires **zero `FAIL` and
zero `UNVERIFIED`**; a `SKIP` is allowed only with a recorded reason. A check that could not
actually inspect its target reports `UNVERIFIED` — which fails the gate on purpose. The gate
is adversarially tested by `scripts/selftest.py`, which injects defects (fabricated number,
missing notes, wrong slide count, modified master/theme, placeholder figure, missing
`output_contract` key, overflowing text) and **exits non-zero if any is not detected**.

**A green report proves exactly the invariants it names and nothing more.** It does not
certify that the presentation is good science, that the narrative is the best one, or that
a reviewer will agree with the content selection. Those remain authorial judgements.

## 9. Limitations

- **Text-layer PDFs.** A scanned PDF needs OCR; without it the run is honestly `partial`
  with the affected pages named.
- **Bitmap figures** yield no data to redraw — they are referenced or cropped, never
  re-derived.
- **No template → degraded deck.** It builds on an internal minimal layout and marks the
  deck "template not applied" rather than pretending.
- **Slide count is negotiated, not forced.** If the material cannot support the requested
  count, the closest honest count is delivered and the conflict is recorded.
- **Visual fidelity is not certified.** Overflow/cropping are checked on geometry (and
  against the render when available); exact visual appearance depends on the renderer.
- Full per-mode detail and detections: `references/failure-modes.md`.

## 10. Repository layout

```
SKILL.md                     the Skill entry point (what the agent reads first)
README.md                    this file — usage, method, limits
scripts/
  common.py                  case loading, output_contract mapping, io, sha256
  ingest.py                  material -> material.json (locators are mandatory)
  plan.py                    narrative + page plan  (the versioned source of truth)
  build_deck.py              template-bound .pptx, master/theme/footer preserved
  render.py                  PDF/PNG previews with renderer discovery
  evidence.py                per-page provenance, quotes re-verified
  verify_deck.py             gates V01-V15  (exit 0 iff PASS)
  result_contract.py         contract-driven result.json
  run.py                     the deterministic chain, partial-safe
  selftest.py                end-to-end + injected-defect tests
  gain_check.py              measured rework metrics on our own fixture
  contract_matrix.py         rubric line -> artifact/check traceability
references/                  runbook, verification_gates (V01-V19), prior-art,
                             schemas, provenance, exports/fonts, failure modes,
                             rubric mapping, measured gain
schemas/                     JSON Schema for every artifact
examples/                    a clearly-labelled minimal call example (示意, not official)
```

## 11. Licensing and attribution

Skill code is original to this submission. Where third-party material was consulted or
adapted, the licence and the attribution obligation are recorded in
`references/provenance-and-evidence.md`. Fonts are **not** redistributed: the Skill refers
to font families by name and relies on the platform image or the template. If you reuse
CC-BY/Font-Awesome icons or OFL fonts in a deck, carry the attribution into your own
documentation — see the licence traps noted in that reference.

## 12. Status vocabulary used in this repository

Deliberately uncollapsed, because collapsing them is how false greens happen:

- check status: `PASS` / `FAIL` / `SKIP` / `UNVERIFIED` / `N/A`
- run status (`result.json.status`): `ok` / `partial` / `failed`
- gates checked a **structure or an invariant**; they never certify scientific quality.
- Gains are reported only as measured on the repository's own **synthetic** fixture. No
  official score is claimed anywhere in this Skill.
