# SYNTHETIC-REV-002 — SYNTHETIC TASK CARD (NOT AN OFFICIAL CASE)

**SYNTHETIC / NOT AN OFFICIAL CASE.** The paper in `materials/` is invented for this fixture, including all
literature counts and reported error metrics. Nothing here is a real meta-analysis.

## Task
Turn the supplied review material into a seminar presentation for a theory group, using
the organizer template in `materials/template-organizer.pptx`.

## Fixed constraints
- Audience: 理论组 seminar, including one external reviewer.
- Duration: 20 minutes speaking time.
- Slide count: 12 slides (min 11, max 12 — a range is acceptable).
- Language: English slide text.
- Template: `materials/template-organizer.pptx` (master, layouts, theme fonts and footer
  must be preserved).
- Sources: `materials/paper.pdf` and `materials/paper.docx`.

## Requirements
1. One page = one claim; at most 5 bullets per slide. Do not paste abstract paragraphs.
2. Every reported metric on a slide must be traceable to the source text or to Table 1/2;
   the review must not present synthetic literature values as real measurements.
3. Figures must come from the source or be redrawn from supplied numbers.
4. Speaker notes of at least 40 characters on every content slide.
5. Keep a page-level evidence map and a render of the deck (PDF + PNG previews).
6. Return every field named in `case.json` → `output_contract`.
