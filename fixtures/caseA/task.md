# SYNTHETIC-EXP-001 — SYNTHETIC TASK CARD (NOT AN OFFICIAL CASE)

**SYNTHETIC / NOT AN OFFICIAL CASE.** The paper in `materials/` is invented for this fixture. Its numbers are
internally consistent but they are not measurements and must never be cited as such.

## Task
Turn the supplied research material into a presentation for a group meeting, using the
organizer template in `materials/template-organizer.pptx`.

## Fixed constraints
- Audience: 课题组组会 (supervisor + peer graduate students), mixed theoretical/experimental.
- Duration: 10 minutes speaking time.
- Slide count: exactly 8 slides (title through conclusion, references allowed as backup).
- Language: English slide text; speaker notes may be English.
- Template: `materials/template-organizer.pptx` (master, layouts, theme fonts and footer
  must be preserved — they are part of the deliverable).
- Sources: `materials/paper.pdf` and `materials/paper.docx` (same paper, two formats).

## Requirements
1. One page = one claim; at most 5 bullets per slide.
2. Every number on a slide must appear verbatim in the source material; no rounding and
   no invented values.
3. Every figure used must be taken from the source (crop) or redrawn from supplied data —
   placeholder or invented curves are not acceptable.
4. Speaker notes of at least 40 characters on every content slide.
5. Keep a page-level evidence map (slide -> figure/table/section/number) and a render of
   the deck (PDF + PNG previews).
6. Return every field named in `case.json` → `output_contract`.
