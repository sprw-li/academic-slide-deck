# Exports, fonts and colour

Every number in this file was **measured on this host on 2026-11-23** by running the shipped
scripts; anything that is arithmetic or a library claim rather than a measurement is labelled
as such. Sources of truth: `scripts/render.py` (discovery, export, raster, probes),
`scripts/verify_deck.py` (V07–V12 read the deck's own OOXML), and
`work/tmp-verify-render/*_test.py` (the measurement harnesses, which are scratch and do not
ship).

---

## 1. The renderer that worked, and the exact invocation

`office-renderer` is a provided platform tool, so the rubric requires real rendered previews
(`task/task.md:86`). The renderer that works here is the **DSH-bundled LibreOffice Kit, driven
through its Node CLI wrapper** — not the raw protocol binary.

```text
<node> <...>/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js convert --input <deck.pptx> --output <deck.pdf>
<node> <...>/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js render  --input <deck.pptx> --output-dir <newdir> --pages "1,2"
```

Measured behaviour (both exit 0):

| Property | Measured value |
|---|---|
| `convert` stdout | `{"backend":"native","missingFonts":[],"outputPath":"..."}` |
| `convert` latency, 2-slide deck | **4.3 s** |
| `convert` output | 33 171 B, header `%PDF-1.7` |
| `render` stdout | JSON manifest: `backend`, `rasterEngine:"libreoffice"`, `dpi:144`, `pageCount`, `images[]` (`path,width,height,byteLength,rectangle`), `missingFonts` |
| `render` output, 13.333 × 7.5 in slide | `page-0001.png` **1920 × 1081** px, 22 564 B (probe deck) |
| End-to-end `render.py` on a 1-slide deck | **15.2 s** (convert + raster + probes) |

### Four measured quirks that the code has to absorb

1. **`--help` is not supported** — `cli.js --help` exits 1 with `Unknown option '--help'`. Never
   probe the CLI that way.
2. **`convert` refuses to overwrite.** A second run into the same work dir failed with
   `EEXIST: file already exists, open '.../deck.pdf'`. `render.py` unlinks the target first.
3. **`render` refuses an existing `--output-dir`.** `EEXIST: file already exists, mkdir
   '.../png'`. `render.py` renders into a scratch directory it clears first, then normalises the
   pages to `slide-XX.png` and removes the scratch dir.
4. **The renderer can hang indefinitely, and a plain subprocess timeout does not save you.**
   Measured (2026-11-23): a `render` attempt on a real case sat for **more than 10 minutes** with
   an idle `node` worker, while the same work measured 9.5 s on an idle machine. The mechanism was
   then reproduced deliberately: the CLI spawns a worker **grandchild** that inherits stdout, so
   killing only the Node process leaves the grandchild holding the pipe and `communicate()`
   blocks forever — the timeout fires but never returns. `render.py` therefore

   * sends child stdout/stderr to **files** instead of pipes, so `wait(timeout=…)` can return;
   * kills the whole **process tree** (`taskkill /F /T`, `killpg` elsewhere) on expiry;
   * caps the *total* render effort with `--budget` (default 300 s) so a hung renderer cannot
     consume the platform's 900 s execution budget;
   * reports the hang honestly (`exit 124`, `error` string naming the timeout, `status` degraded)
     instead of blocking or fabricating a preview.

   Reproduced with `work/tmp-verify-render/timeout_test.py`: a `node` parent that spawns a
   long-lived pipe-holding child now returns `rc=124` in **7.8 s** against a 5 s deadline, and a
   normal child still returns its output. Before the fix that same shape hung.

There is **no `soffice` anywhere** on this host: not on `PATH`, no `%PROGRAMFILES%\LibreOffice`
install, and the kit's `program\` tree contains data only. The `soffice` branch in `render.py`
is a portability fallback for hosts that do have LibreOffice; it was **not** exercised here, and
that is why it is written to degrade rather than assumed to work.

## 2. Renderer discovery (no author-machine path in the shipped file)

`render.py` never hard-codes a machine path. It tries, in order:

| # | Strategy | What it looks at |
|---|---|---|
| 0 | `--renderer-cli` / `--node` | explicit operator override |
| 1 | `LO_KIT_CLI` / `LIBREOFFICE_KIT_CLI` / `DSH_LO_KIT_CLI` | env |
| 2 | `SOFFICE_BIN` / `LIBREOFFICE_KIT_BIN` / `LO_KIT_BIN` / `LIBREOFFICE_BIN` | env |
| 3 | `PATH` | `soffice`, `soffice.exe`, `soffice.bin`, `libreoffice`, `impress`, `localc` |
| 4 | known install dirs | `%PROGRAMFILES%`, `%PROGRAMFILES(X86)%`, `%LOCALAPPDATA%`, `/usr/bin`, `/usr/lib/libreoffice`, `/opt/libreoffice*`, `/Applications/...` |
| 5 | bounded search | ancestors of the `node` executable, of every `PATH` entry, of every env var whose value names a `node_modules`/`app.asar` location, and of the cwd — each probed for `node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js` |

Strategy 5 is what actually fires on this host, and it fires **without any path being written
into the script**: walking up from an env var that points inside the app's `node_modules`
reaches the `node_modules` parent, and the kit's relative path is then found there. Measured:
`strategy: "bounded-search"`, `renderer.name: "libreoffice-kit"`, `status: "ok"`.

Honest degradation was measured by emptying `PATH` of the runtime entries and dropping the
env var: `status: "skipped"`, `renderer.available: false`, `pdf.path: ""`, `pngs: []`, one
`uncertainty` entry, exit 1 — and **no fabricated preview path**.

## 3. PDF → PNG

The kit rasterises directly, so Poppler is only a fallback (`pdftoppm`, found on `PATH` here).
`PIL` cannot rasterise a PDF vector page and `pypdf` cannot either, so neither is used for that
job. **ImageMagick's `convert` is deliberately never invoked**: on Windows `convert.exe` is the
filesystem-conversion utility in `system32`, and calling it would be a silent trap rather than a
renderer.

## 4. Fonts — what is measured, and what a preview therefore proves

The organizer template names **Liberation Serif** as both the major and minor latin theme font,
and that font is **not installed on this host** (no match in `%WINDIR%\Fonts` or the per-user
font dir). That is the important case, and it was measured end to end on the real fixture
`fixtures/caseA/materials/template-organizer.pptx`:

| Observation | Measured value |
|---|---|
| Deck built from the organizer template, theme font | `majorFont latin = Liberation Serif`, `minorFont latin = Liberation Serif` |
| Kit `render`/`convert` `missingFonts` | `["Liberation Serif"]` |
| Fonts **actually embedded in the produced PDF** (read back out of `/Resources /Font`, via `pypdf`) | `/BAAAAA+TimesNewRomanPSMT`, `/CAAAAA+TimesNewRomanPSMT`, `/DAAAAA+TimesNewRomanPSMT`, `/EAAAAA+ArialMT` |
| `verify_deck.py` V12 verdict | **PASS** — "1 declared typeface(s) are not installed on this host (`['Liberation Serif']`) and every one of them is openly reported as substituted" |
| Second data point, a deck naming two fonts | `missingFonts: ["Liberation Serif", "DefinitelyNotAFont Xyz"]` |

So the substitution is **reported, not hidden** — the behaviour the docs promise is real. But
read the middle row carefully: the PDF does **not** contain Liberation Serif; LibreOffice
substituted **Times New Roman**, which is metric-compatible with Liberation Serif and therefore
a plausible-looking substitute.

### The font fields in `render_report.json` (one measurement, one field)

`missing_fonts` is the **single** canonical field: a list, always present (possibly empty), and it
is the union of the renderer's own reports from the export call and the raster call. The raw
per-call values and the read-back embedding list live together in `_font_evidence`:

```jsonc
{ "missing_fonts": ["Liberation Serif"],
  "_font_evidence": {
     "kit_convert_missingFonts": ["Liberation Serif"],
     "kit_render_missingFonts":  ["Liberation Serif"],
     "pdf_fonts_embedded": ["/BAAAAA+TimesNewRomanPSMT", "/CAAAAA+TimesNewRomanPSMT",
                            "/DAAAAA+TimesNewRomanPSMT", "/EAAAAA+ArialMT"] } }
```

An earlier revision emitted *two* fields for this (`missing_fonts` from the raster manifest and
`_missing_fonts` from the export), which read as a self-contradiction — one of them looked like
"nothing missing" while `uncertainty` said a font was substituted. Both now come from one
measurement, and the `uncertainty` sentence is generated from that same list, so the fields can no
longer disagree.

The kit reports *which* families it could not find; it does **not** say what it used instead. That
is why `render.py` reads the answer back out of the produced PDF with `pypdf`
(`_font_evidence.pdf_fonts_embedded`), turning "the renderer says a font is missing" into
measured evidence of what was actually embedded. When `pypdf` is unavailable the field is `null`
and `_font_evidence.pdf_fonts_embedded_note` says the embedding was not checked — an unchecked
property is labelled, not assumed.

> **Can a preview-only render be trusted for font fidelity? No.**
> The preview is a faithful picture of how *LibreOffice* laid the deck out with the
> *substituted* face. It is trustworthy for structure, page count, geometry, colour and gross
> overflow, and it is the only cheap way to see the deck at all — but it is **not** a
> certificate that PowerPoint will render the same glyphs or the same line breaks with the
> organizer's real font. `render_report.json.missing_fonts` is exactly the field that tells you
> the preview is a substitute render, and V12 fails if a substituted face is *not* reported
> anywhere.

Font resolution inside `verify_deck.py` scans the platform font directories plus matplotlib's
bundled faces, by family name. V12 deliberately ignores two things that produced a **false
failure on the first run**: the theme's *supplemental* script-font table (Angsana New, Cordia
New, Batang, … — a fallback map, not a font in use) and `+mj-lt`/`+mn-ea` style references,
which are pointers to the theme font rather than font names. On the good fixture the
attributable set is exactly `['Arial', 'Calibri']`.

**Licensing (for allowlist/attribution purposes, from `SKILL_Downloaded/repos/06-fonts-icons/INDEX.md`).**
Liberation Serif/Sans/Mono ship under **OFL-1.1**; OFL permits embedding in a PPTX and
redistribution with the work, but forbids selling the font on its own and forbids distributing
a modified version under the Reserved Font Name. The skill does not bundle or subset any font —
it only names families and reports whether they resolve — so nothing here redistributes a font
binary.

## 5. Colour and theme preservation

Measured by round-tripping solid swatches of known RGB through PPTX → PDF → PNG at 144 dpi and
reading the pixel back with `PIL`:

| Source RGB | Pixel in rendered PNG | Result |
|---|---|---|
| `#0072B2` (0, 114, 178) | (0, 114, 178) | **exact** |
| `#D55E00` (213, 94, 0) | (213, 94, 0) | **exact** |

No colour management was applied, so an sRGB fill survives the export unchanged at this dpi.
Theme preservation is checked structurally rather than by pixel: V08 compares the deck's
`ppt/theme/theme1.xml` `<a:fontScheme>` against the template's, and V07 compares every
`slideMaster`/`slideLayout` part C14N-normalised (12 parts on the organizer fixture,
**identical**). Measured on the organizer run: `major_latin`/`minor_latin` stayed
`Liberation Serif` and V08 passed.

**Scope note (asserted, not measured):** theme *colour* preservation is covered by the same
part-level comparison as V08 — the theme XML is compared, so a changed theme colour is caught —
but only one solid-swatch pixel round-trip was measured, so no per-slot theme-colour palette
result is claimed here.

## 6. What was actually run

| Command (shape) | Result |
|---|---|
| `render.py --work W --in probe.pptx --renderer-cli <cli> --node <node>` | `status ok`, `deck.pdf` + `slide-01/02.png` |
| `render.py --work W --in probe.pptx` (no overrides) | `status ok`, `renderer libreoffice-kit`, `strategy bounded-search` |
| `render.py` twice into the **same** work dir | both runs `status ok` (the EEXIST quirks are fixed; before the fix the second run degraded to `partial`) |
| `render.py` with no renderer reachable | `status skipped`, `available false`, uncertainty recorded, exit 1 |
| `timeout_test.py`: parent spawning a pipe-holding grandchild | `rc=124` in **7.8 s** against a 5 s deadline; normal child still returns its output; no leftover scratch files |
| `render.py` + `verify_deck.py` on the real organizer template | `render ok`; V07 PASS, V08 PASS, V09 PASS, V12 PASS |
| `run.py` on fixture caseA (full pipeline, own scratch) | **16 PASS / 2 FAIL (V10, V16)**, exit 3; render stage completed in the 44 s run |
| `verify_deck.py` on the good fixture | verdict **PASS**, 18/18 checks (V01–V18) |
| `negtest.py` (18 injected defects) | 18/18 detected by the expected check, all non-zero exit |

### Latency and concurrency

| Measurement | Value |
|---|---|
| `convert`, 2-slide probe | 4.3 s |
| `render.py` on a 1-slide deck (export + raster + probes) | 15.2 s |
| `run.py` end to end on fixture caseA (ingest → … → result, 8 slides, incl. render) | **44.1 s** |
| `verify_deck.py` on an 8-slide deck | 2.2 s |
| Once observed, while other agents ran LibreOffice concurrently | a render chain that measures 9.5 s stalled for **>600 s** |

The stall is the reason for the tree-kill + `--budget` design above. **Do not run this pipeline's
render stage in parallel with another LibreOffice job.**

## 7. Failure modes

Each row below is injected and detected in `negtest.py` unless noted; the ledger with the
"does not certify" column is in [`failure-modes.md`](failure-modes.md).

| Failure | Detection |
|---|---|
| no renderer reachable | `status:"skipped"`, `available:false`, `uncertainty[]`, no paths claimed |
| re-render into a used output dir | target/scratch cleared before calling the kit |
| uniformly blank preview | `nonbackground_ratio < 5e-4` → `blank_slides_suspected` + uncertainty |
| rendered page count ≠ slide count | `uncertainty` entry comparing the two counts |
| substituted/missing font | `missingFonts` copied into `render_report.json`; V12 fails if it is not reported |
| unreadable deck / missing `python-pptx` | geometry checks report **UNVERIFIED** (fails the gate) rather than passing |
