# Rubric mapping — every scored line → the artifact or check that addresses it

Source of the rubric: `task/task.md:40-50` of this submission's challenge
(`ppt-6ad93034` / `ARMI:2611.000074v1`). Verbatim sub-item names are quoted; the right-hand
column names something a reviewer can **open and run**, not a promise.

This file is the prose twin of `scripts/contract_matrix.py`, which checks mechanically that
every rubric line below points at a path and a check id that actually exist. If they ever
disagree, `contract_matrix.py` is the one that fails.

> Grading for this challenge is `human_review_only` (`task/challenge.json.scoring`). Nothing
> in this repository claims an official score. What follows is traceability, not scoring.

---

## 1. 实际效果 — 40 分

### 1.1 任务完成度 (15)
> 「是否按任务要求交付一版可打开的学术演示稿，包括页数、页序、每页承担的叙事角色、演讲备注和导出件是否齐备」

| Addressed by | Where |
|---|---|
| Slide count / order / per-page narrative role | `plan.json` → `slides[].index`, `.role`, `.claim`; gate **V01** (count within the case constraint), **V02** (title + claim present) |
| Speaker notes are a real deliverable | `plan.json.slides[].notes`; gate **V03** (≥40 chars per slide) |
| A genuinely openable, editable deck | `deck.pptx` produced by `build_deck.py`; `build_report.json` records layout, filled placeholders and notes length per slide |
| Export件齐备 | `deck.pdf` + `slide-XX.png` from `render.py`; `render_report.json` reports the renderer actually used — or an explicit skip |
| Constraints actually honoured | `plan.json.constraints` parsed from the case; violations surface as **V01**, not as prose |

### 1.2 结果正确性 (15)
> 「页面上的数字、图表、图注、引用编号与论文原文、数据表是否一致，有无编造或按印象取整」

This is the line where existing agents fail hardest, so it has the strictest gates:

| Addressed by | Where |
|---|---|
| Every slide number appears **verbatim** in the source | gate **V04**, which re-derives the numbers from `material.json.numbers` (not from `plan.json`, so a lying planner is still caught) |
| No rounding by impression | `plan.py` may only use `raw` strings taken from `material.json`; a mismatched `raw` is a **V04** defect |
| Figure/table/caption consistency | captions are extracted *with a verbatim quote locator*, never retyped; `evidence_map.json.pages[].figures[]/tables[]` |
| Reference numbers / section references | `material.json.sections[]` + `claims[].locator`; `evidence.py` re-verifies each quote against the source text and marks it `verified` |
| **Ingest-side prevention** | an item without a verbatim locator is *dropped* from `material.json` and recorded in `warnings` — fabrication is structurally hard, not merely forbidden |

### 1.3 结果完整性 (5)
> 「是否覆盖论文主线、页结构、图表引用、讲解逻辑、模板约束、导出渲染六类要素」

| Element | Artifact |
|---|---|
| 论文主线 | `plan.json.narrative.arc` + `main_thread[]` |
| 页结构 | `plan.json.slides[].role`; gate **V02/V14** |
| 图表引用 | `plan.json.slides[].figure_ref/table_ref`; gate **V05** (must resolve in `material.json`) |
| 讲解逻辑 | `plan.json.slides[].notes` + `evidence[]`; gate **V03** |
| 模板约束 | `build_report.json.invariants`; gates **V07/V08/V09** |
| 导出渲染 | `render_report.json` + previews; gates **V10/V11/V12** |

Explicitly recorded so the reviewer can see the *negative* space: `plan.json.excluded[]`
lists material deliberately left off the deck, with a reason.

### 1.4 结果可核验性 (5)
> 「源文件、页级金标准、导出件、逐页 diff、运行日志和版本记录能否支撑他人独立复核并复跑」

| Addressed by | Where |
|---|---|
| Source identity | every `result.json.sources[]` entry carries path + sha256 |
| Page-level evidence | `evidence_map.json` — one entry per slide, quote-verified |
| Export件 | `deck.pdf` / `slide-XX.png` + `render_report.json` (renderer, command, sizes) |
| Per-page diff | `versions/` change records; the deck is reproducible from `plan.json` |
| Run log | `run.log` with per-stage timings and exit codes |
| Independent re-run | every stage is an independently runnable CLI; `verify_deck.py` exits 0 iff PASS; `selftest.py` re-runs the whole chain with injected defects |

## 2. 使用增益 — 25 分

### 2.1 质量增益 (15)
> 「在相同输入论文、相同模板、相同工具权限下，加载成果后页面质量相比基础 Agent 是否提升」

**What this repository can honestly claim:** the *mechanisms* that remove the failure modes
listed in §1.2 are implemented and gated (V04–V12). **What it cannot claim:** a measured
quality delta against a no-Skill baseline — the platform produces that comparison, and this
repository does not invent one. See `references/measured-gain.md`, which is explicit about
the boundary.

### 2.2 效率增益 (5)
> 「是否减少返工页数、重排次数、导出复查轮次或整体人工修改量」

| Addressed by | Where |
|---|---|
| Fewer reordering rounds | the page plan is a single artifact that can be revised before any file is written |
| Fewer export re-check rounds | `render.py` + V10–V12 make overflow/crop/font problems machine-detected instead of visually hunted |
| Fewer manual edits | `deck.pptx` is generated; the human edits `plan.json` and rebuilds |
| **Traceable mechanism, not a claim** | one command runs the whole chain (`run.py`), and each stage is re-runnable, so the platform's 64-request budget is not spent on rediscovery |
| Measured metric | `scripts/gain_check.py` counts pages needing reordering, export re-check rounds, numbers failing traceability, and template violations — **on our own synthetic fixture only** |

### 2.3 增益可验证性 (5)
> 「是否通过公平对照证明页面质量或制作效率的改善」

Gate **V13** + `gain_check.py` + `references/measured-gain.md`. `measured-gain.md` states
which numbers are measured, on which fixture, and which are *not* claimed. A reviewer can
re-run `gain_check.py` and get the same numbers.

## 3. 创新增益 — 15 分

### 3.1 新增能力 (5)
> 「是否形成基础 Agent 或现有方案不具备、或难以稳定实现的页序规划、模板槽位填充、图表数据绑定与局部编辑能力」

| Capability | Why a base agent does not have it *reliably* (see `references/prior-art.md` for the same capabilities as they exist in prior art) |
|---|---|
| **Contract-driven result emission** | the platform declares `output_contract` keys that differ per case; the Skill reads them and maps synonyms, rather than shipping a fixed field list |
| **Template-safe slot filling with asserted invariants** | the master/theme/footer are re-opened and compared at XML level (V07–V09), so preservation is observed |
| **Number-level data binding** | numbers are bound to verbatim source quotes at ingest; V04 re-derives them independently downstream |
| **Evidence map as a first-class artifact** | per-page provenance is a deliverable, not a log line |

### 3.2 方法改进 (5)
> 「内容取舍、页序组织、槽位约束、图表重绘、编辑定位、导出校验等方法是否带来实质改进」

- **Traceability upstream instead of checking downstream**: unlocatable facts never enter
  the plan, which removes a whole class of fabrication rather than filtering it afterwards.
- **Plan-as-versioned-source / deck-as-build-product**: makes every iteration revertible by
  construction, which answers 「改坏后如何回退」 structurally instead of procedurally.
- **Figures classified by what the material actually supports** (`redraw-from-data` /
  `source-crop` / `reference-only`) instead of assuming a figure can always be redrawn.
- **Gate semantics that can fail**: `UNVERIFIED` fails the gate, so the checker cannot
  silently degrade into a rubber stamp.

### 3.3 创新证据 (5)
> 「是否证明创新点相对科学空间现有能力、开源方案或商用插件的边界和价值」

| Evidence | Where |
|---|---|
| Boundary vs. existing tooling | `references/prior-art.md` — an explicit delimitation against python-pptx, pptx-template, pptx-mcp, Office-PowerPoint-MCP-Server, PptAgent, deepPPT, the Anthropic pptx helpers, md2pptx/pandoc/marp, and plain schema-driven emission. It also states the **narrow** claim (an audit layer, not a new generation method) and marks sibling-submission overlap as UNVERIFIED rather than claiming independence |
| What the Skill does *not* do or certify | `references/failure-modes.md` + the "does not certify" list |
| Adversarial test of the checker itself | `selftest.py` injects defects and exits non-zero if any is *not* detected; the gate set was seen red before it was trusted |
| Measured, reproducible gain metrics | `gain_check.py`, re-runnable |
| Known limits named | `README.md` §9, `references/failure-modes.md` |

> **Honesty note on 创新增益.** The council-style reading of this dimension is that
> "evidence-anchored planning + template-safe filling + contract-driven emission" is *not*
> beyond a well-prompted agent — only the *enforcement and auditability* are novel. The claim
> in this repository is deliberately narrowed to that, because an overstated novelty claim is
> worth less than a defensible narrow one. See `references/prior-art.md` §3–§4.

## 4. 可靠体验 — 10 分

### 4.1 调用与易用性 (3)
Trigger-rich frontmatter (`SKILL.md`) with explicit *use when* clauses and *do NOT use for*
clauses; a single documented command; `references/runbook.md` for stage-by-stage debugging;
`examples/` for a labelled minimal invocation.

### 4.2 运行稳定性 (4)
- Deterministic chain, no network dependency at run time.
- **Partial-safety**: any stage failure still writes the artifacts produced so far and a
  `status:"partial"` `result.json` naming the limits (`task.md:88` requires this).
- Works across paper types (experimental / review / theoretical fixtures) and slide-count
  constraints; optional dependencies are guarded with honest degradation.

### 4.3 异常处理 (3)
| Situation | Behaviour |
|---|---|
| no PDF text layer | OCR if available, else `partial` with the affected pages named |
| bitmap figure, no extractable data | `reference-only`, no invented values |
| template missing/unreadable | degraded internal layout, marked "template not applied" |
| font missing | substitution reported in `render_report.json.missingFonts` + `uncertainty` |
| slide count unreachable | closest honest count + recorded conflict |
| no renderer | `status:"skipped"`, `available:false`, no preview path claimed |
| any stage fails | partial deliverables first, then `status:"partial"` |

## 5. 生态贡献 — 10 分

### 5.1 文档完整性 (3)
`README.md` (usage/method/limits), `references/runbook.md`, `references/spec-schema.md`,
`references/provenance-and-evidence.md`, `references/exports-fonts-colour.md`,
`references/failure-modes.md`, this file, `references/measured-gain.md`.

### 5.2 资源可复用性 (4)
`schemas/` (JSON Schema per artifact), `examples/` (a labelled minimal case),
and the fixture builder so a third party can regenerate the development cases.
Every stage is a standalone CLI, so pieces can be reused without adopting the whole Skill.

### 5.3 分享推广价值 (3)
The design pattern — *plan as versioned source, evidence as a deliverable, template
invariants asserted at XML level, a gate that can fail* — is reusable well beyond
presentations. `references/failure-modes.md` is written as transferable experience.

## 6. GitHub 公开作品链接 — 5 分

`GITHUB_URL.txt` at the submission root (line 1 = repository link; line 2 = the file or
directory inside it). The link must open **without logging in**.

---

## How to check this file is not lying

```bash
python scripts/contract_matrix.py     # exit 0 iff every rubric line above resolves to a real path + check id
```

A rubric line mapped to a path that does not exist, or to a check id that `verify_deck.py`
never emits, is a failure — not a formatting nit.
