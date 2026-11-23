# Prior art and the honest boundary of this Skill's contribution

Written because a reviewer is entitled to ask "what does this do that an existing tool or a
well-prompted agent does not?", and because overstating that answer is the fastest way to
lose the 创新增益 points entirely.

**The short version:** this Skill is **not a new method for generating presentations**. It is
an *auditing layer* — verbatim provenance, template preservation asserted at OOXML level, a
contract-driven result emitter, and a gate set that can fail — applied inside one
contest-shaped pipeline. Everything else in it is engineering that others have published
before. The claim below is deliberately narrow.

---

## 1. What already exists (delimitation)

The comparison set is not hypothetical: these were the tools available in this workspace's
reference material bank while the Skill was written (the bank is a separate offline library
with per-entry download receipts and licence fields; descriptions below come from its
generated catalog, not from a fresh reading of each repository's source).

| Prior art | What it already does | What it therefore does **not** let us claim |
|---|---|---|
| **python-pptx** | reads/writes PPTX: shapes, placeholders, layouts, speaker notes | "template-based deck generation" — trivial, decades-old library capability |
| **pptx-template** | fills an existing `.pptx` template's slots from data | "slot filling into a real template" |
| **pptx-mcp** | `template_manager` / `layout_manager`: placeholder and master inspection | "layout/master awareness" |
| **Office-PowerPoint-MCP-Server** (`ppt-mcp`) | add/edit slides, *preserve template*, insert images | "template-preserving editing via a tool interface" |
| **PptAgent** | LLM agent that plans an outline, generates, and then **refines** real `.pptx` decks | "an LLM plans pages and produces a deck" |
| **deepPPT** (Chinese academic PPT skill) | audience-specific styling rules (组会/答辩/Journal Club) + QA/revision loop | "academic-presentation-specific skill with self-QA" |
| **anthropic document skills (pptx)** | `html2pptx` plus OOXML-level Python helpers for PPTX generation/editing | "OOXML-level manipulation beyond the python-pptx API" |
| **md2pptx / pandoc (`--reference-doc`) / marp / reveal.js / slidev** | text-markup → slide deck, some with theming | "source-to-deck conversion" and "themed output" |
| **pptxgenjs / pptx2md** | deck generation in JS; deck → markdown incl. notes | "programmatic deck I/O" |
| **generic schema-driven JSON emission + CI gates** | standard software engineering | "contract-driven output" or "verification gates" as ideas |

## 2. What a well-prompted base agent can already do

A capable agent given the same inputs can be asked to read the PDF, pick pages, fill a
template, cite sources, and emit a `result.json`. Nothing here is *impossible* for it.

The difference is not capability, it is **enforcement and auditability**:

| | prompted base agent | this Skill |
|---|---|---|
| cites the paper | usually | only locatable facts are admitted at all (`ingest.py` drops the rest) |
| numbers match the source | checked by eye, if at all | re-derived from the material by gate **V04**, downstream of the planner |
| template preserved | asserted in prose | told apart from the template at OOXML level by **V07–V09** |
| contract fields | whatever it remembers to write | every declared key, with unmatched keys surfaced rather than dropped (**V13**) |
| result | plausible narrative | a gate that reports `FAIL`/`UNVERIFIED` and exits non-zero |

That is a real difference in reliability. It is **not** a new way to make slides.

## 3. The narrow claim this Skill does make

> Given (a) a paper, (b) a real master/layout-based `.pptx` template, (c) a case file that
> declares its own `output_contract`, produce a deck in which **every on-slide fact carries a
> re-verifiable verbatim locator**, **the template is provably intact**, and **every declared
> result field is accounted for** — and do it through a chain that is deterministic,
> re-runnable, partial-safe, and able to say no.

Specifically, these are the pieces that are not simply "what the library does":

1. **Drop-don't-guess at ingest.** Unlocatable figures/tables/numbers never enter
   `material.json`; they are recorded in `warnings[]`. Removal-beats-filtering is a design
   choice with a testable consequence, not a library feature.
2. **Numbers verified downstream of the planner** (V04 re-derives from `material.json`,
   not from `plan.json`), so a planner that hallucinates a number cannot validate its own work.
3. **Template invariants judged from the produced package's XML** rather than from the
   builder's own report — the builder's claim and the gate's finding are separate objects.
4. **Contract-driven emission with an explicit unmatched-key channel**
   (`_unmapped_contract_keys`), so an unknown hidden-case field degrades visibly instead of
   vanishing.
5. **A gate that can fail**: `UNVERIFIED` blocks, and the gate set was adversarially tested by
   defect injection *before* being trusted — plus a minimum-content-coverage gate (V19) so an
   empty extraction cannot vacuously pass every `all()`.

## 4. What is explicitly *not* claimed

- Not a new layout algorithm, not a new planning model, not a better-looking deck.
- Not a demonstration that the output is **better** than a base agent's — that comparison is
  produced by the platform, and this repository does not have it. `references/measured-gain.md`
  states exactly what was measured (rework metrics on our own synthetic fixture) and what was not.
- Not a replacement for PptAgent-style iterative refinement, nor for the MCP servers' editing
  surface; those are complementary and could drive the same gates.
- Not evidence that the gates certify **scientific** quality. They cover structure,
  traceability, template preservation and export integrity. A green report is a statement about
  those invariants and nothing else.

## 5. Overlap with sibling submissions in this contest series

Closest siblings available locally: the graphical-abstract Skill and the academic
template-adaptation Skill (`ARMI:2611.000062v1` / `ARMI:2611.000072v1` in this workspace). The
template-adaptation Skill also fills templates and asserts content preservation, for DOCX
rather than PPTX, with its own fidelity gates; the graphical-abstract Skill also builds an
evidence map for a single figure.

**No byte-level diff or feature-level comparison against those submissions is included here,
so their marginal overlap is UNVERIFIED.** Stating the overlap exists is honest; claiming
independence without the diff would not be. If this Skill is compared against them by a
reviewer, the defensible distinction is the artifact set (`plan.json` / `evidence_map.json` /
`verify_report.json` / contract-driven `result.json`) and the medium (an editable,
template-bound `.pptx` with real rendered previews), not the idea of verification itself.
